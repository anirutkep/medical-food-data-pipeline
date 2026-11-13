#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cross-Platform Hybrid OCR & Dual-Pass Verification Tool.

Extracts text tokens with bounding boxes from nutrition label photographs,
groups them into horizontal Y-bands, sorts columns left-to-right, and performs
dual-pass numeric verification against Pass 1 JSON or canonical brand data.py.

Supported Backends:
  - 'vision'      : macOS native Apple Vision (VNRecognizeTextRequest)
  - 'winrt'       : Windows native WinRT OCR (Windows.Media.Ocr via PowerShell)
  - 'tesseract'   : Cross-platform Tesseract OCR (via pytesseract / Pillow)
  - 'vlm_fallback': Structured fallback returning instructions for multimodal VLM
  - 'auto'        : Platform-aware native OCR -> Tesseract -> VLM fallback
"""

from __future__ import annotations

import argparse
import base64
import importlib.util
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple


def group_tokens_by_y(
    tokens: List[Tuple[float, float, str]],
    y_thresh: float = 0.012,
) -> List[Dict[str, Any]]:
    """
    Group OCR text observations by vertical (Y-coordinate) row bands.

    Parameters:
      tokens: List of (yc, xc, text) tuples where:
              yc: normalized Y center (0.0 = bottom, 1.0 = top)
              xc: normalized X left (0.0 .. 1.0)
              text: recognized string token
      y_thresh: normalized vertical band threshold (default 0.012)

    Returns:
      List of dicts: {"y": round(avg_y, 4), "columns": List[str], "numbers": List[float]}
      sorted top-to-bottom, columns left-to-right.
    """
    if not tokens:
        return []

    # Sort top-to-bottom (-yc), then left-to-right (xc)
    sorted_tokens = sorted(tokens, key=lambda item: (-item[0], item[1]))

    grouped: List[List[Tuple[float, float, str]]] = []
    for yc, xc, text in sorted_tokens:
        if grouped and abs(yc - grouped[-1][0][0]) <= y_thresh:
            grouped[-1].append((yc, xc, text))
        else:
            grouped.append([(yc, xc, text)])

    rows: List[Dict[str, Any]] = []
    for band in grouped:
        # Sort columns left-to-right
        band.sort(key=lambda item: item[1])
        avg_y = sum(item[0] for item in band) / len(band)
        cols = [item[2] for item in band]
        nums = [
            float(m.replace(",", ""))
            for col in cols
            for m in re.findall(r"\b\d+(?:,\d{3})*(?:\.\d+)?\b", col)
        ]
        rows.append({"y": round(avg_y, 4), "columns": cols, "numbers": nums})
    return rows


def _find_powershell_executable() -> str:
    """Find available PowerShell binary on Windows or cross-platform."""
    for exe in ("powershell.exe", "powershell", "pwsh"):
        if shutil.which(exe):
            return exe
    raise RuntimeError("No PowerShell executable found on system (tried powershell.exe, powershell, pwsh).")


def _run_winrt_ocr(
    image_path: str | Path,
    y_thresh: float = 0.012,
    languages: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """
    Execute native Windows WinRT OCR via PowerShell -EncodedCommand with UTF-8 environment variables.
    """
    img_path = Path(image_path).expanduser().resolve()
    if not img_path.exists():
        raise FileNotFoundError(f"Image not found: {img_path}")

    ps_exe = _find_powershell_executable()
    lang_str = ",".join(languages) if languages else "th-TH,en-US"

    # PowerShell WinRT OCR script
    # Uses $env:OCR_IMAGE_PATH and $env:OCR_LANG to guarantee unicode/space safety
    ps_script = r"""
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$imagePath = $env:OCR_IMAGE_PATH
$reqLang = $env:OCR_LANG
if (-not $reqLang) { $reqLang = "th-TH,en-US" }

try {
    Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
    [Windows.Media.Ocr.OcrEngine, Windows.Foundation.UniversalApiContract, ContentType=WindowsRuntime] | Out-Null
    [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation.UniversalApiContract, ContentType=WindowsRuntime] | Out-Null
    [Windows.Storage.StorageFile, Windows.Foundation.UniversalApiContract, ContentType=WindowsRuntime] | Out-Null
} catch {
    Write-Output '{"error": "WINRT_UNAVAILABLE", "message": "Windows WinRT OCR API not available on this system."}'
    exit 0
}

$availLangs = [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages
if ($reqLang -like "*th*") {
    $foundTh = $false
    foreach ($l in $availLangs) {
        if ($l.LanguageTag -like "th*" -or $l.LanguageTag -eq "th") {
            $foundTh = $true
            break
        }
    }
    if (-not $foundTh) {
        Write-Output '{"error": "MISSING_THAI_PACK", "message": "Windows WinRT Thai (th) OCR language pack not installed. Please install Thai OCR via Windows Settings -> Time & Language -> Language -> Thai -> Options -> Optical Character Recognition."}'
        exit 0
    }
}

$asTaskGeneric = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
}

function AwaitOperation($asyncOp, $type) {
    $m = $asTaskGeneric.MakeGenericMethod($type)
    $netTask = $m.Invoke($null, @($asyncOp))
    $netTask.Wait(-1) | Out-Null
    return $netTask.Result
}

try {
    $fileOp = [Windows.Storage.StorageFile]::GetFileFromPathAsync($imagePath)
    $file = AwaitOperation $fileOp ([Windows.Storage.StorageFile])

    $streamOp = $file.OpenAsync([Windows.Storage.FileAccessMode]::Read)
    $stream = AwaitOperation $streamOp ([Windows.Storage.Streams.IRandomAccessStream])

    $decoderOp = [Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)
    $decoder = AwaitOperation $decoderOp ([Windows.Graphics.Imaging.BitmapDecoder])

    $bitmapOp = $decoder.GetSoftwareBitmapAsync()
    $bitmap = AwaitOperation $bitmapOp ([Windows.Graphics.Imaging.SoftwareBitmap])

    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
    if (-not $engine -and $availLangs.Count -gt 0) {
        $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($availLangs[0])
    }
    if (-not $engine) {
        Write-Output '{"error": "ENGINE_INIT_FAILED", "message": "Failed to initialize WinRT OcrEngine."}'
        exit 0
    }

    $ocrOp = $engine.RecognizeAsync($bitmap)
    $ocrResult = AwaitOperation $ocrOp ([Windows.Media.Ocr.OcrResult])

    $words = @()
    foreach ($line in $ocrResult.Lines) {
        foreach ($w in $line.Words) {
            $words += @{
                text = $w.Text
                x = [double]$w.BoundingRect.X
                y = [double]$w.BoundingRect.Y
                w = [double]$w.BoundingRect.Width
                h = [double]$w.BoundingRect.Height
            }
        }
    }

    $result = @{
        width = [double]$decoder.PixelWidth
        height = [double]$decoder.PixelHeight
        words = $words
    }

    $json = ConvertTo-Json -InputObject $result -Compress -Depth 5
    Write-Output $json
} catch {
    $errMsg = $_.Exception.Message
    Write-Output ('{"error": "EXECUTION_ERROR", "message": ' + (ConvertTo-Json $errMsg) + '}')
}
"""

    encoded_cmd = base64.b64encode(ps_script.encode("utf-16le")).decode("ascii")

    env = os.environ.copy()
    env["OCR_IMAGE_PATH"] = str(img_path)
    env["OCR_LANG"] = lang_str

    try:
        proc = subprocess.run(
            [ps_exe, "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded_cmd],
            capture_output=True,
            text=True,
            encoding="utf-8",
            env=env,
        )
    except Exception as exc:
        raise RuntimeError(f"Failed to execute PowerShell WinRT runner: {exc}") from exc

    if proc.returncode != 0:
        raise RuntimeError(f"PowerShell WinRT OCR execution failed (code {proc.returncode}): {proc.stderr}")

    stdout_text = proc.stdout.strip()
    if not stdout_text:
        raise RuntimeError(f"PowerShell WinRT OCR produced no output. Stderr: {proc.stderr}")

    try:
        payload = json.loads(stdout_text)
    except Exception as exc:
        raise RuntimeError(f"Failed to parse WinRT OCR output as JSON: {stdout_text}") from exc

    if "error" in payload:
        err_type = payload.get("error")
        err_msg = payload.get("message", err_type)
        if err_type == "MISSING_THAI_PACK":
            raise RuntimeError(f"Windows WinRT Thai (th) OCR language pack not installed: {err_msg}")
        raise RuntimeError(f"Windows WinRT OCR error ({err_type}): {err_msg}")

    w_img = float(payload.get("width", 1.0))
    h_img = float(payload.get("height", 1.0))
    if w_img <= 0.0:
        w_img = 1.0
    if h_img <= 0.0:
        h_img = 1.0

    tokens: List[Tuple[float, float, str]] = []
    for word in payload.get("words", []):
        x = float(word.get("x", 0.0))
        y = float(word.get("y", 0.0))
        w = float(word.get("w", 0.0))
        h = float(word.get("h", 0.0))
        text = str(word.get("text", "")).strip()
        if text:
            # Normalized bottom-left matching Apple Vision coordinates
            yc = 1.0 - (y + h / 2.0) / h_img
            xc = x / w_img
            tokens.append((yc, xc, text))

    return group_tokens_by_y(tokens, y_thresh=y_thresh)


def _run_vision_ocr(
    image_path: str | Path,
    y_thresh: float = 0.012,
    languages: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """
    Execute macOS Apple Vision OCR (VNRecognizeTextRequest).
    """
    try:
        import Vision  # type: ignore
        from Foundation import NSURL  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "Apple Vision framework (pyobjc-framework-Vision / Foundation) is not available."
        ) from exc

    img_path = Path(image_path).expanduser().resolve()
    if not img_path.exists():
        raise FileNotFoundError(f"Image not found: {img_path}")

    req = Vision.VNRecognizeTextRequest.alloc().init()
    req.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
    req.setRecognitionLanguages_(languages or ["th-TH", "en-US"])
    req.setUsesLanguageCorrection_(False)

    handler = Vision.VNImageRequestHandler.alloc().initWithURL_options_(
        NSURL.fileURLWithPath_(str(img_path)), None
    )
    ok, err = handler.performRequests_error_([req], None)
    if not ok:
        raise RuntimeError(f"Vision OCR failed for {img_path}: {err}")

    tokens: List[Tuple[float, float, str]] = []
    for obs in req.results() or []:
        candidates = obs.topCandidates_(1)
        if not candidates:
            continue
        bb = obs.boundingBox()
        yc = float(bb.origin.y + bb.size.height / 2.0)
        xc = float(bb.origin.x)
        text = str(candidates[0].string()).strip()
        if text:
            tokens.append((yc, xc, text))

    return group_tokens_by_y(tokens, y_thresh=y_thresh)


def _run_tesseract_ocr(
    image_path: str | Path,
    y_thresh: float = 0.012,
    languages: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    """
    Execute Tesseract OCR via pytesseract and Pillow (PIL).
    """
    try:
        import pytesseract  # type: ignore
        from pytesseract import Output  # type: ignore
        from PIL import Image  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Tesseract OCR requires 'pytesseract' and 'Pillow' (PIL).") from exc

    img_path = Path(image_path).expanduser().resolve()
    if not img_path.exists():
        raise FileNotFoundError(f"Image not found: {img_path}")

    tess_langs = []
    if languages:
        for l in languages:
            l_lower = l.lower()
            if "th" in l_lower and "tha" not in tess_langs:
                tess_langs.append("tha")
            elif "en" in l_lower and "eng" not in tess_langs:
                tess_langs.append("eng")
    lang_arg = "+".join(tess_langs) if tess_langs else "tha+eng"

    try:
        with Image.open(img_path) as img:
            img_w, img_h = img.size
            data = pytesseract.image_to_data(img, lang=lang_arg, output_type=Output.DICT)
    except Exception as exc:
        raise RuntimeError(f"Tesseract OCR failed: {exc}") from exc

    tokens: List[Tuple[float, float, str]] = []
    n_boxes = len(data.get("text", []))
    for i in range(n_boxes):
        text = str(data["text"][i]).strip()
        if not text:
            continue
        x = float(data["left"][i])
        y = float(data["top"][i])
        w = float(data["width"][i])
        h = float(data["height"][i])
        yc = 1.0 - (y + h / 2.0) / float(img_h)
        xc = x / float(img_w)
        tokens.append((yc, xc, text))

    return group_tokens_by_y(tokens, y_thresh=y_thresh)


def extract_rows_from_image(
    image_path: str | Path,
    y_thresh: float = 0.012,
    languages: Optional[List[str]] = None,
    backend: str = "auto",
) -> Dict[str, Any]:
    """
    Run OCR using the specified or auto-detected backend.

    Parameters:
      image_path: Path to label image
      y_thresh: Vertical band grouping threshold
      languages: Optional list of language tags (e.g. ['th-TH', 'en-US'])
      backend: 'auto', 'vision', 'winrt', 'tesseract', or 'vlm_fallback'

    Returns:
      Dict with status, backend, image path, and extracted rows or fallback instructions.
    """
    img_path = Path(image_path).expanduser().resolve()
    if not img_path.exists():
        raise FileNotFoundError(f"Image not found: {img_path}")

    valid_backends = {"auto", "vision", "winrt", "tesseract", "vlm_fallback"}
    if backend not in valid_backends:
        raise ValueError(f"Invalid backend '{backend}'. Must be one of {sorted(valid_backends)}")

    if backend == "vlm_fallback":
        return {
            "status": "vlm_fallback",
            "backend": "vlm_fallback",
            "image": str(img_path),
            "rows": [],
            "reason": "Explicitly requested vlm_fallback",
            "instructions": "Inspect the image directly with multimodal vision, extract Pass 1 rows/JSON, and verify via --compare-json or --data-py.",
        }

    if backend == "auto":
        last_reason = ""
        if sys.platform == "darwin":
            try:
                rows = _run_vision_ocr(img_path, y_thresh=y_thresh, languages=languages)
                return {"status": "ok", "backend": "vision", "image": str(img_path), "rows": rows}
            except Exception as exc:
                last_reason = f"Apple Vision failed: {exc}"
        elif sys.platform == "win32":
            try:
                rows = _run_winrt_ocr(img_path, y_thresh=y_thresh, languages=languages)
                return {"status": "ok", "backend": "winrt", "image": str(img_path), "rows": rows}
            except Exception as exc:
                last_reason = f"Windows WinRT failed: {exc}"
        else:
            last_reason = f"Platform '{sys.platform}' has no native OCR"

        # Fallback to Tesseract
        try:
            rows = _run_tesseract_ocr(img_path, y_thresh=y_thresh, languages=languages)
            return {"status": "ok", "backend": "tesseract", "image": str(img_path), "rows": rows}
        except Exception as exc:
            last_reason += f"; Tesseract failed: {exc}"

        # Final fallback to VLM
        return {
            "status": "vlm_fallback",
            "backend": "vlm_fallback",
            "image": str(img_path),
            "rows": [],
            "reason": f"Auto backends exhausted ({last_reason})",
            "instructions": "Inspect the image directly with multimodal vision, extract Pass 1 rows/JSON, and verify via --compare-json or --data-py.",
        }

    # Explicit backends
    if backend == "vision":
        rows = _run_vision_ocr(img_path, y_thresh=y_thresh, languages=languages)
        return {"status": "ok", "backend": "vision", "image": str(img_path), "rows": rows}
    elif backend == "winrt":
        rows = _run_winrt_ocr(img_path, y_thresh=y_thresh, languages=languages)
        return {"status": "ok", "backend": "winrt", "image": str(img_path), "rows": rows}
    elif backend == "tesseract":
        rows = _run_tesseract_ocr(img_path, y_thresh=y_thresh, languages=languages)
        return {"status": "ok", "backend": "tesseract", "image": str(img_path), "rows": rows}

    raise ValueError(f"Unhandled backend: {backend}")


def _parse_row_numbers_from_json(data: Any) -> Set[float]:
    """Parse numeric values from all 3 JSON shapes or raise ValueError."""
    nums: Set[float] = set()
    found_structure = False

    def _extract_from_row_list(row_list: List[Any]) -> None:
        nonlocal found_structure
        found_structure = True
        for row in row_list:
            if not isinstance(row, dict):
                raise ValueError("Expected row entry to be a dict")
            for n in row.get("numbers", []):
                try:
                    nums.add(round(float(n), 3))
                except (ValueError, TypeError):
                    pass

    if isinstance(data, list):
        _extract_from_row_list(data)
    elif isinstance(data, dict):
        if "rows" in data:
            if not isinstance(data["rows"], list):
                raise ValueError("Expected 'rows' to be a list")
            _extract_from_row_list(data["rows"])
        else:
            # Mapping {image_path: rows_list | {"rows": rows_list, ...}}
            if not data:
                found_structure = True
            for k, v in data.items():
                if isinstance(v, list):
                    _extract_from_row_list(v)
                elif isinstance(v, dict) and "rows" in v:
                    if not isinstance(v["rows"], list):
                        raise ValueError(f"Expected 'rows' for '{k}' to be a list")
                    _extract_from_row_list(v["rows"])
                else:
                    raise ValueError(f"Unrecognized structure in JSON map for key '{k}'")
    else:
        raise ValueError("JSON data must be a list or dict")

    if not found_structure:
        raise ValueError("Could not parse any valid row structures from JSON")

    return nums


def compare_with_pass1(rows: List[Dict[str, Any]], pass1_json_path: str | Path) -> List[str]:
    """
    Compare OCR-extracted numeric tokens per row against a prior JSON pass.

    Accepts flat list, top-level dict with 'rows', and format json output mapping.
    """
    json_path = Path(pass1_json_path).expanduser().resolve()
    if not json_path.exists():
        raise FileNotFoundError(f"Pass 1 JSON file not found: {json_path}")

    try:
        data = json.loads(json_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise ValueError(f"Failed to load or parse JSON from {json_path}: {exc}") from exc

    ref_nums = _parse_row_numbers_from_json(data)
    ocr_nums: Set[float] = set()
    for r in rows:
        for n in r.get("numbers", []):
            try:
                ocr_nums.add(round(float(n), 3))
            except (ValueError, TypeError):
                pass

    only_ocr = sorted(ocr_nums - ref_nums)
    only_ref = sorted(ref_nums - ocr_nums)
    diffs: List[str] = []
    if only_ocr:
        diffs.append(f"Numbers only in OCR pass: {only_ocr}")
    if only_ref:
        diffs.append(f"Numbers only in reference pass: {only_ref}")
    return diffs


def _load_and_validate_data_py(
    data_py_path: str | Path,
    product_id: Optional[str] = None,
) -> Tuple[Any, List[Dict[str, Any]]]:
    """Load data.py module and validate the PRODUCTS schema contract."""
    py_path = Path(data_py_path).expanduser().resolve()
    if not py_path.exists():
        raise FileNotFoundError(f"data.py file not found: {py_path}")

    spec = importlib.util.spec_from_file_location("custom_brand_data", py_path)
    if spec is None or spec.loader is None:
        raise ValueError(f"Could not load Python module spec from {py_path}")

    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as exc:
        raise ValueError(f"Failed to execute Python module {py_path}: {exc}") from exc

    if not hasattr(module, "PRODUCTS"):
        raise ValueError(f"Missing required 'PRODUCTS' attribute in {py_path}")

    products = getattr(module, "PRODUCTS")
    if not isinstance(products, list) or len(products) == 0:
        raise ValueError(f"'PRODUCTS' in {py_path} must be a non-empty list of dicts")

    for idx, prod in enumerate(products):
        if not isinstance(prod, dict):
            raise ValueError(f"Product at index {idx} in {py_path} is not a dict")
        if "nutrients" not in prod or not isinstance(prod["nutrients"], list):
            raise ValueError(f"Product at index {idx} in {py_path} is missing 'nutrients' list")

    if product_id is not None:
        matched = [p for p in products if p.get("id") == product_id]
        if not matched:
            raise ValueError(f"Product id '{product_id}' not found in PRODUCTS of {py_path}")
        target_products = matched
    else:
        target_products = products

    return module, target_products


def compare_with_data_py(
    rows: List[Dict[str, Any]],
    data_py_path: str | Path,
    product_id: Optional[str] = None,
) -> List[str]:
    """
    Compare OCR-extracted numeric tokens against canonical nutrients in brand data.py.
    """
    _, target_products = _load_and_validate_data_py(data_py_path, product_id=product_id)

    ref_nums: Set[float] = set()
    for prod in target_products:
        for nut in prod["nutrients"]:
            if isinstance(nut, (tuple, list)) and len(nut) >= 4:
                # 4-tuple: (name, unit, value_base, value_serving)
                for val in (nut[2], nut[3]):
                    if val is not None:
                        try:
                            ref_nums.add(round(float(val), 3))
                        except (ValueError, TypeError):
                            pass

    ocr_nums: Set[float] = set()
    for r in rows:
        for n in r.get("numbers", []):
            try:
                ocr_nums.add(round(float(n), 3))
            except (ValueError, TypeError):
                pass

    only_ocr = sorted(ocr_nums - ref_nums)
    only_ref = sorted(ref_nums - ocr_nums)
    diffs: List[str] = []
    if only_ocr:
        diffs.append(f"Numbers only in OCR rows: {only_ocr}")
    if only_ref:
        diffs.append(f"Numbers only in data.py: {only_ref}")
    return diffs


def main(argv: Optional[List[str]] = None) -> int:
    """CLI entrypoint for OCR row extraction and dual-pass verification."""
    parser = argparse.ArgumentParser(
        description="Cross-platform OCR row extractor and dual-pass verification tool."
    )
    parser.add_argument("images", nargs="+", help="Path(s) to nutrition label images (.jpg/.png)")
    parser.add_argument(
        "--y-thresh",
        type=float,
        default=0.012,
        help="Normalized Y-band threshold (default: 0.012)",
    )
    parser.add_argument(
        "--backend",
        choices=("auto", "vision", "winrt", "tesseract", "vlm_fallback"),
        default="auto",
        help="OCR backend to use (default: auto)",
    )
    parser.add_argument(
        "--format",
        choices=("table", "markdown", "json"),
        default="table",
        help="Output format (default: table)",
    )
    parser.add_argument(
        "--lang",
        help="Comma-separated language codes, e.g. th-TH,en-US (default: platform defaults)",
    )
    parser.add_argument(
        "--compare-json",
        help="Optional JSON file from Pass 1 for dual-pass numeric diff",
    )
    parser.add_argument(
        "--data-py",
        help="Optional canonical brand data.py file for dual-pass numeric diff",
    )
    parser.add_argument(
        "--product-id",
        help="Optional product ID to filter when using --data-py",
    )

    args = parser.parse_args(argv)

    languages = [l.strip() for l in args.lang.split(",") if l.strip()] if args.lang else None

    # Early validation of compare-json and data-py files & schemas
    if args.compare_json:
        cmp_path = Path(args.compare_json).expanduser().resolve()
        if not cmp_path.exists():
            print(f"Error: --compare-json file not found: {args.compare_json}", file=sys.stderr)
            return 1
        try:
            raw_data = json.loads(cmp_path.read_text(encoding="utf-8"))
            _parse_row_numbers_from_json(raw_data)
        except Exception as exc:
            print(f"Error: Invalid --compare-json file: {exc}", file=sys.stderr)
            return 1

    if args.data_py:
        try:
            _load_and_validate_data_py(args.data_py, product_id=args.product_id)
        except Exception as exc:
            print(f"Error: Invalid --data-py file or schema: {exc}", file=sys.stderr)
            return 1

    all_output: Dict[str, Any] = {}
    has_failure = False

    for img in args.images:
        try:
            res = extract_rows_from_image(
                img,
                y_thresh=args.y_thresh,
                languages=languages,
                backend=args.backend,
            )
        except (FileNotFoundError, RuntimeError, ValueError) as exc:
            print(f"Error processing {img}: {exc}", file=sys.stderr)
            has_failure = True
            continue

        all_output[str(img)] = res
        rows = res.get("rows", [])

        if args.format == "table":
            if res.get("status") == "vlm_fallback":
                print(f"=== {img} (vlm_fallback: {res.get('reason')}) ===")
                print(f"Instructions: {res.get('instructions')}")
            else:
                print(f"=== {img} ({len(rows)} rows, backend={res.get('backend')}) ===")
                for r in rows:
                    print(f"[y={r['y']:.3f}] " + " | ".join(r["columns"]))
        elif args.format == "markdown":
            if res.get("status") == "vlm_fallback":
                print(f"### {Path(img).name} (VLM Fallback)")
                print(f"> {res.get('instructions')}")
            else:
                print(f"### {Path(img).name}")
                print("| Y | Columns |")
                print("| :---: | :--- |")
                for r in rows:
                    print(f"| {r['y']:.3f} | " + " / ".join(r["columns"]) + " |")

        # Compare operations (only on successfully extracted rows)
        if res.get("status") == "ok":
            if args.compare_json:
                try:
                    diffs = compare_with_pass1(rows, args.compare_json)
                    if diffs:
                        print(f"DUAL-PASS DIFF for {img}: " + "; ".join(diffs), file=sys.stderr)
                        has_failure = True
                    else:
                        print(f"DUAL-PASS MATCH for {img}: 100% numeric overlap")
                except Exception as exc:
                    print(f"Error during --compare-json: {exc}", file=sys.stderr)
                    has_failure = True

            if args.data_py:
                try:
                    diffs = compare_with_data_py(rows, args.data_py, product_id=args.product_id)
                    if diffs:
                        print(f"DATA-PY DIFF for {img}: " + "; ".join(diffs), file=sys.stderr)
                        has_failure = True
                    else:
                        print(f"DATA-PY MATCH for {img}: 100% numeric overlap")
                except Exception as exc:
                    print(f"Error during --data-py: {exc}", file=sys.stderr)
                    has_failure = True

    if args.format == "json":
        print(json.dumps(all_output, ensure_ascii=False, indent=2))

    return 1 if has_failure else 0


if __name__ == "__main__":
    raise SystemExit(main())
