# -*- coding: utf-8 -*-
"""
Idempotent Cross-Platform Unit Normalization Engine.

Applies MOPH Notification No. 445 (B.E. 2566) Thai RDI & Section 2.5 Unit Normalization
to Set 2 Markdowns (02_*.md) and Excel Workbooks (*.xlsx) with dynamic formulas.
"""

from __future__ import annotations
import argparse
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from unit_normalizer import (
    THAI_RDI_445_STANDARDS,
    canonicalize_nutrient_name,
    normalize_nutrient,
)

NORM_HDR_1 = "หน่วยมาตรฐาน (Thai RDI 445)"
NORM_HDR_2 = "ปริมาณมาตรฐานต่อ 1,000 kcal"
NORM_NOTE_MARKER = "<!-- UNIT_NORMALIZATION_NOTE_445 -->"


def _get_styles(font_name: str = "TH Sarabun New") -> Dict[str, Any]:
    """Generate corporate styling palette keyed to the specified font family."""
    thin_side = Side(border_style="thin", color="BFBFBF")
    double_side = Side(border_style="double", color="1B365D")
    return {
        "title": Font(name=font_name, size=15, bold=True, color="1B365D"),
        "hdr": Font(name=font_name, size=12, bold=True, color="FFFFFF"),
        "data": Font(name=font_name, size=11, color="000000"),
        "bold": Font(name=font_name, size=11, bold=True, color="1B365D"),
        "fill_hdr": PatternFill("solid", fgColor="1B365D"),
        "fill_sub": PatternFill("solid", fgColor="D9E1F2"),
        "fill_zebra": PatternFill("solid", fgColor="F2F2F2"),
        "fill_total": PatternFill("solid", fgColor="E2EFDA"),
        "border_data": Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side),
        "border_total": Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side),
    }


def _parse_float(cell_text: str) -> Optional[float]:
    cleaned = cell_text.replace("*", "").replace(",", "").strip()
    if not cleaned or cleaned in ("-", "N/A", "ไม่ระบุ", "None"):
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def _is_separator_row(line: str) -> bool:
    stripped = line.strip().strip("|").replace("|", "").replace(":", "").replace("-", "").replace(" ", "")
    return stripped == ""


def _text_display_width(text: str) -> float:
    """Calculate UTF-8 glyph-aware visual width with 1.6x factor for non-ASCII Thai text."""
    w = 0.0
    for ch in text:
        if ord(ch) > 127:
            w += 1.6
        else:
            w += 1.0
    return w


def detect_workbook_type(wb: openpyxl.Workbook) -> str:
    """
    Detect whether an Excel workbook is a multi-product Brand Master or a Single Standalone.

    Returns "master" if overview/comparison sheets exist, else "standalone".
    """
    for name in wb.sheetnames:
        if "ภาพรวมเปรียบเทียบ" in name or "สรุปพลังงาน" in name:
            return "master"
        if name in ("00_ภาพรวมเปรียบเทียบ_1000kcal", "01_สรุปพลังงาน_1000kcal"):
            return "master"
    return "standalone"


def detect_workbook_font(wb: openpyxl.Workbook) -> str:
    """
    Scan non-normalized product sheets for the most frequent populated cell font name.

    Falls back to 'TH Sarabun New' if no populated font is detected.
    """
    font_counts: Dict[str, int] = {}
    for sname in wb.sheetnames:
        if sname.startswith(("00_", "01_", "02_Normalized", "Normalized_")) or "Normalized" in sname:
            continue
        ws = wb[sname]
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value is not None and cell.font and cell.font.name:
                    fname = str(cell.font.name).strip()
                    if fname:
                        font_counts[fname] = font_counts.get(fname, 0) + 1

    if font_counts:
        return max(font_counts.items(), key=lambda x: x[1])[0]
    return "TH Sarabun New"


def process_set2_markdown(md_path: str | Path, dry_run: bool = False) -> bool:
    """
    Add or update normalized unit & quantity columns in Section 4 Nutrition Facts table.

    Idempotent: Replaces prior normalized columns and note blocks without duplication.
    Returns False untouched if Section 4 table is absent.
    """
    md_p = Path(md_path).resolve()
    if not md_p.is_file():
        return False

    content = md_p.read_text(encoding="utf-8")

    # Strip existing normalization note block if re-running
    content = re.sub(
        rf"\n*{re.escape(NORM_NOTE_MARKER)}.*?{re.escape(NORM_NOTE_MARKER)}\n*",
        "\n",
        content,
        flags=re.DOTALL,
    )

    lines = content.splitlines()
    ingredients_text = "\n".join(lines[:70])

    hdr_idx = None
    for i, line in enumerate(lines):
        if line.strip().startswith("|") and ("หน่วย" in line) and ("1,000 kcal" in line or "1000 kcal" in line):
            if "สารอาหารหลัก" in line or "ตัวคูณพลังงาน" in line or "% w/w" in line:
                continue
            hdr_idx = i
            break

    if hdr_idx is None:
        return False

    # Find end of table
    end_idx = hdr_idx + 1
    while end_idx < len(lines) and lines[end_idx].strip().startswith("|"):
        end_idx += 1

    hdr_cols = [c.strip() for c in lines[hdr_idx].strip().strip("|").split("|")]
    already_norm = len(hdr_cols) >= 2 and hdr_cols[-2] == NORM_HDR_1 and hdr_cols[-1] == NORM_HDR_2
    if already_norm:
        hdr_cols = hdr_cols[:-2]

    name_col = next((idx for idx, h in enumerate(hdr_cols) if any(k in h for k in ("สารอาหาร", "รายการ"))), 1)
    unit_col = next(
        (idx for idx, h in enumerate(hdr_cols) if h in ("หน่วย", "หน่วย (Unit)", "หน่วยตามฉลาก") or "หน่วย" in h),
        None,
    )
    kcal_col = next((idx for idx, h in enumerate(hdr_cols) if "1,000 kcal" in h or "1000 kcal" in h), None)

    if unit_col is None or kcal_col is None:
        return False

    new_lines: List[str] = lines[:hdr_idx]
    notes_seen: List[str] = []

    # Header row
    new_hdr = "| " + " | ".join(hdr_cols + [NORM_HDR_1, NORM_HDR_2]) + " |"
    new_lines.append(new_hdr)

    # Separator row
    sep_cols = [c.strip() for c in lines[hdr_idx + 1].strip().strip("|").split("|")]
    if already_norm and len(sep_cols) >= 2:
        sep_cols = sep_cols[:-2]
    new_sep = "| " + " | ".join(sep_cols + [":---:", ":---:"]) + " |"
    new_lines.append(new_sep)

    # Data rows
    for r_idx in range(hdr_idx + 2, end_idx):
        row_str = lines[r_idx].strip()
        if not row_str.startswith("|") or _is_separator_row(row_str):
            continue
        cols = [c.strip() for c in row_str.strip("|").split("|")]
        if already_norm and len(cols) >= 2:
            cols = cols[:-2]
        while len(cols) < len(hdr_cols):
            cols.append("")

        raw_name = cols[name_col]
        raw_unit = cols[unit_col]
        raw_val_str = cols[kcal_col]
        val_num = _parse_float(raw_val_str)

        if not raw_unit or val_num is None:
            norm_unit_cell = "-" if raw_unit else ""
            norm_val_cell = raw_val_str if ("<" in raw_val_str or ">" in raw_val_str) else ""
        else:
            res = normalize_nutrient(raw_name, raw_unit, val_num, ingredients_text)
            norm_unit_cell = str(res["target_unit"])
            norm_val_cell = f"**{res['normalized_1000kcal']:,.2f}**"
            note = str(res["note"])
            if note and note not in notes_seen:
                notes_seen.append(note)

        new_lines.append("| " + " | ".join(cols + [norm_unit_cell, norm_val_cell]) + " |")

    if notes_seen:
        new_lines.append("")
        new_lines.append(NORM_NOTE_MARKER)
        new_lines.append("> [!NOTE]")
        new_lines.append("> **เกณฑ์การแปลงหน่วยมาตรฐาน (ประกาศกระทรวงสาธารณสุข ฉบับที่ 445 พ.ศ. 2566 บัญชีหมายเลข 3 และข้อ 2.5):**")
        for n in notes_seen:
            new_lines.append(f"> - {n}")
        new_lines.append(NORM_NOTE_MARKER)

    new_lines.extend(lines[end_idx:])
    new_content = "\n".join(new_lines) + ("\n" if content.endswith("\n") else "")

    if not dry_run:
        md_p.write_text(new_content, encoding="utf-8")

    return True


def _find_excel_nutrition_table(ws: Any) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int]]:
    """Return (header_row, name_col, unit_col, kcal1000_col) in a product worksheet."""
    for r in range(1, ws.max_row + 1):
        row_vals = [str(ws.cell(r, c).value or "").strip() for c in range(1, ws.max_column + 1)]
        has_unit = any(v in ("หน่วย", "หน่วย (Unit)", "หน่วยตามฉลาก") or v == "หน่วย" for v in row_vals)
        has_1000 = any("1,000 kcal" in v or "1000 kcal" in v for v in row_vals)
        if not (has_unit and has_1000):
            continue
        if any("ตัวคูณพลังงาน" in v or "% w/w" in v or "สัดส่วนการกระจาย" in v for v in row_vals):
            continue

        name_col = next((i + 1 for i, v in enumerate(row_vals) if any(k in v for k in ("สารอาหาร", "รายการ"))), None)
        unit_col = next(
            (i + 1 for i, v in enumerate(row_vals) if v in ("หน่วย", "หน่วย (Unit)", "หน่วยตามฉลาก") or v == "หน่วย"),
            None,
        )
        kcal_col = next((i + 1 for i, v in enumerate(row_vals) if "1,000 kcal" in v or "1000 kcal" in v), None)

        if name_col and unit_col and kcal_col:
            return r, name_col, unit_col, kcal_col
    return None, None, None, None


def _extract_sheet_ingredients(ws: Any, before_row: int) -> str:
    texts: List[str] = []
    for r in range(1, before_row):
        for c in range(1, min(6, ws.max_column + 1)):
            v = ws.cell(r, c).value
            if v:
                texts.append(str(v))
    return " ".join(texts)


def process_excel_workbook(
    xlsx_path: str | Path,
    font_name: Optional[str] = None,
    dry_run: bool = False,
) -> bool:
    """
    Build or update normalized sheet with 100% dynamic Excel formulas and glyph-aware width.

    Target sheet name:
      - '02_Normalized_445' for Master workbooks
      - 'Normalized_1000kcal_445' for Standalone workbooks
    """
    xlsx_p = Path(xlsx_path).resolve()
    if not xlsx_p.is_file():
        return False

    wb = openpyxl.load_workbook(str(xlsx_p), data_only=False)
    wb_type = detect_workbook_type(wb)
    target_sheet_name = "02_Normalized_445" if wb_type == "master" else "Normalized_1000kcal_445"

    resolved_font = font_name or detect_workbook_font(wb)
    styles = _get_styles(resolved_font)

    # Remove previous normalized sheet if present (idempotent)
    if target_sheet_name in wb.sheetnames:
        del wb[target_sheet_name]

    records: List[Dict[str, Any]] = []
    for sname in wb.sheetnames:
        if sname.startswith(("00_", "01_", "02_Normalized", "Normalized_")) or "Normalized" in sname:
            continue
        ws = wb[sname]
        hdr_row, name_col, unit_col, kcal_col = _find_excel_nutrition_table(ws)
        if not hdr_row or not name_col or not unit_col or not kcal_col:
            continue

        ing_text = _extract_sheet_ingredients(ws, hdr_row)
        kcal_col_letter = get_column_letter(kcal_col)

        for r in range(hdr_row + 1, ws.max_row + 1):
            raw_name = ws.cell(r, name_col).value
            raw_unit = ws.cell(r, unit_col).value
            raw_val = ws.cell(r, kcal_col).value
            if not raw_name:
                continue

            name_str = str(raw_name).strip()
            if "สัดส่วนการกระจายพลังงาน" in name_str or name_str == "สารอาหารหลัก" or "พลังงานรวม" in name_str:
                break
            if not raw_unit or raw_val is None:
                continue

            unit_str = str(raw_unit).strip()
            if isinstance(raw_val, str) and not raw_val.startswith("="):
                continue

            res = normalize_nutrient(name_str, unit_str, 1.0, ing_text)
            canon = res["canonical_key"]
            rdi_val = THAI_RDI_445_STANDARDS[canon]["rdi"] if canon and canon in THAI_RDI_445_STANDARDS else None
            rdi_no = THAI_RDI_445_STANDARDS[canon]["no"] if canon and canon in THAI_RDI_445_STANDARDS else None

            records.append({
                "sheet": sname,
                "row": r,
                "col_letter": kcal_col_letter,
                "rdi_no": rdi_no,
                "name": name_str,
                "raw_unit": unit_str,
                "factor": float(res["factor"]),
                "target_unit": str(res["target_unit"]),
                "rdi_val": rdi_val,
                "note": str(res["note"]),
            })

    if not records:
        return False

    ws_norm = wb.create_sheet(title=target_sheet_name)
    ws_norm.views.sheetView[0].showGridLines = True

    ws_norm["A1"] = "ตารางค่าสารอาหารมาตรฐานต่อ 1,000 kcal (Normalized ตามประกาศกระทรวงสาธารณสุข ฉบับที่ 445 พ.ศ. 2566)"
    ws_norm["A1"].font = styles["title"]
    ws_norm["A2"] = "อ้างอิงบัญชีหมายเลข 3 (Thai RDIs 33 รายการ) และข้อ 2.5 ค่าการแปลงหน่วยสำหรับวิตามินและแร่ธาตุ"
    ws_norm["A2"].font = styles["data"]

    for c_idx in range(1, 12):
        cell_sub = ws_norm.cell(3, c_idx, "หมวด: ข้อมูลสารอาหารที่แปลงเป็นหน่วยมาตรฐานกลาง (ASCII-safe)" if c_idx == 1 else "")
        cell_sub.font = styles["bold"]
        cell_sub.fill = styles["fill_sub"]
        cell_sub.border = styles["border_data"]

    headers = [
        "ผลิตภัณฑ์ (Sheet)",
        "ลำดับ Thai RDI",
        "สารอาหาร",
        "ค่าเดิมต่อ 1,000 kcal",
        "หน่วยเดิม",
        "ตัวคูณแปลงหน่วย (ข้อ 2.5/SI)",
        "ปริมาณมาตรฐาน (1,000 kcal)",
        "หน่วยมาตรฐาน (445)",
        "Thai RDI (445)",
        "% Thai RDI (ต่อ 1,000 kcal)",
        "หมายเหตุสมมติฐาน (ข้อ 2.5)",
    ]
    hdr_r = 4
    for c_idx, h in enumerate(headers, 1):
        cell = ws_norm.cell(hdr_r, c_idx, h)
        cell.font = styles["hdr"]
        cell.fill = styles["fill_hdr"]
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = styles["border_data"]
    ws_norm.row_dimensions[hdr_r].height = 28

    for idx, rec in enumerate(records):
        r = hdr_r + 1 + idx
        sheet_ref = f"'{rec['sheet']}'!{rec['col_letter']}{rec['row']}"
        vals = [
            rec["sheet"],
            rec["rdi_no"] if rec["rdi_no"] is not None else "-",
            rec["name"],
            f"=ROUND({sheet_ref}, 2)",
            rec["raw_unit"],
            rec["factor"],
            f"=ROUND(D{r}*F{r}, 2)",
            rec["target_unit"],
            rec["rdi_val"] if rec["rdi_val"] is not None else "-",
            f"=ROUND((G{r}/I{r})*100, 1)" if rec["rdi_val"] is not None else "-",
            rec["note"],
        ]
        for c_idx, v in enumerate(vals, 1):
            cell = ws_norm.cell(r, c_idx, v)
            cell.font = styles["bold"] if c_idx == 7 else styles["data"]
            cell.border = styles["border_data"]
            if idx % 2 == 1:
                cell.fill = styles["fill_zebra"]
            if c_idx in (2, 5, 8):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in (4, 6, 7, 9, 10):
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if c_idx in (4, 7):
                    cell.number_format = "#,##0.00"
                elif c_idx == 6:
                    cell.number_format = "0.######"
                elif c_idx == 10:
                    cell.number_format = "0.0"

    # Total / Summary footer row
    total_r = hdr_r + 1 + len(records)
    for c_idx in range(1, 12):
        tcell = ws_norm.cell(total_r, c_idx)
        if c_idx == 1:
            tcell.value = "รวมรายการสารอาหารที่ Normalize แล้ว"
        elif c_idx == 3:
            tcell.value = f"=ROUND(COUNTA(C{hdr_r+1}:C{total_r-1}), 0)"
        elif c_idx == 4:
            tcell.value = "รายการ"
        tcell.font = styles["bold"]
        tcell.fill = styles["fill_total"]
        tcell.border = styles["border_total"]

    # UTF-8 glyph-aware column width computation
    base_min_widths = [24, 14, 34, 18, 16, 20, 22, 18, 14, 20, 44]
    for c_idx in range(1, 12):
        col_letter = get_column_letter(c_idx)
        max_len = float(base_min_widths[c_idx - 1])
        for r_idx in range(4, total_r + 1):
            cell_val = ws_norm.cell(r_idx, c_idx).value
            if cell_val is not None:
                s_val = str(cell_val)
                if s_val.startswith("="):
                    s_val = "123,456.78"
                w = _text_display_width(s_val) + 5.0
                if w > max_len:
                    max_len = w
        ws_norm.column_dimensions[col_letter].width = min(round(max_len, 1), 60)

    if not dry_run:
        wb.save(str(xlsx_p))

    return True


def normalize_workspace(
    root: str | Path,
    exclude: Optional[List[str]] = None,
    font_name: Optional[str] = None,
    dry_run: bool = False,
) -> Tuple[int, int]:
    """
    Recursively process all Set 2 Markdowns and Excel Workbooks in a workspace root.

    Returns (md_processed_count, xlsx_processed_count).
    """
    root_p = Path(root).resolve()
    exclude_set: Set[str] = set()
    if exclude:
        for item in exclude:
            for part in item.split(","):
                p_clean = part.strip()
                if p_clean:
                    exclude_set.add(p_clean)
    else:
        exclude_set = {"dist", ".git", ".agents", "tests", "node_modules", "__pycache__", ".venv"}

    md_ok = 0
    xlsx_ok = 0

    for dirpath, dirnames, filenames in os.walk(root_p, topdown=True):
        # Prune only child directory names matching exclude
        dirnames[:] = [d for d in dirnames if d not in exclude_set]

        for fname in sorted(filenames):
            if fname.startswith("02_") and fname.endswith(".md"):
                fpath = Path(dirpath) / fname
                if process_set2_markdown(fpath, dry_run=dry_run):
                    md_ok += 1
            elif fname.endswith(".xlsx") and not fname.startswith("~$"):
                fpath = Path(dirpath) / fname
                if process_excel_workbook(fpath, font_name=font_name, dry_run=dry_run):
                    xlsx_ok += 1

    return md_ok, xlsx_ok


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Cross-platform Unit Normalizer for Set 2 MD and Excel")
    parser.add_argument("--root", default=".", help="Root directory to normalize (default: current directory)")
    parser.add_argument("--exclude", action="append", help="Directories to exclude (comma-separated or repeatable)")
    parser.add_argument("--font", default=None, help="Explicit font name for Excel (e.g. 'Cordia New', 'TH Sarabun New')")
    parser.add_argument("--dry-run", action="store_true", help="Perform dry run without modifying files")
    args = parser.parse_args(argv)

    md_ok, xlsx_ok = normalize_workspace(
        root=args.root,
        exclude=args.exclude,
        font_name=args.font,
        dry_run=args.dry_run,
    )
    mode_str = "[DRY-RUN] Processed" if args.dry_run else "Updated"
    print(f"=== Unit Normalization: {mode_str} {md_ok} Set 2 Markdown(s), {xlsx_ok} Excel Workbook(s) ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
