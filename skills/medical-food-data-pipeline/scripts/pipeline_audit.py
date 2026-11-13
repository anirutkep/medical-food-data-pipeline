# -*- coding: utf-8 -*-
"""
Pipeline Audit Engine for Medical Food & Enteral Nutrition Deliverables.

Validates canonical data.py products and filesystem deliverables against clinical
and regulatory constraints:
  1. Dual-column (v1, v2) consistency (base vs serving ratio)
  2. Macronutrient caloric consistency (4*P + 9*F + 4*C vs base_kcal within 6%)
  3. Thai RDI 445 plausibility bounds [2.0%, 1000.0%] for micronutrients
  4. Cross-product cohort outlier detection (ratio > 10.0 or < 0.1 vs cohort median)
  5. Deliverables parity: Set 2 Markdown (02_*.md) and Standalone Excel (*_Nutrition_1000kcal.xlsx)
"""

from __future__ import annotations
import argparse
import importlib.util
import os
import statistics
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from unit_normalizer import (
    THAI_RDI_445_STANDARDS,
    canonicalize_nutrient_name,
    normalize_nutrient,
)


def load_canonical_products(data_py_path: str | Path) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
    """
    Dynamically import data.py module and validate BRAND_META and PRODUCTS contract.

    Raises ValueError if the module does not conform to the canonical schema.
    """
    p_path = Path(data_py_path).resolve()
    if not p_path.is_file():
        raise FileNotFoundError(f"Canonical data file not found: {p_path}")

    spec = importlib.util.spec_from_file_location("brand_data_module", str(p_path))
    if spec is None or spec.loader is None:
        raise ValueError(f"Unable to load Python module from {p_path}")

    mod = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(mod)
    except Exception as exc:
        raise ValueError(f"Error executing module {p_path}: {exc}") from exc

    if not hasattr(mod, "BRAND_META") or not isinstance(mod.BRAND_META, dict):
        raise ValueError("Missing or invalid 'BRAND_META' dict in data module")

    if not hasattr(mod, "PRODUCTS") or not isinstance(mod.PRODUCTS, list) or len(mod.PRODUCTS) == 0:
        raise ValueError("Missing or empty 'PRODUCTS' list in data module")

    for idx, prod in enumerate(mod.PRODUCTS):
        if not isinstance(prod, dict):
            raise ValueError(f"Item at index {idx} in PRODUCTS is not a dict")
        if not prod.get("id"):
            raise ValueError(f"Product at index {idx} missing required 'id'")
        if not prod.get("folder"):
            raise ValueError(f"Product '{prod.get('id')}' missing required 'folder'")
        base_kcal = prod.get("base_kcal")
        if not isinstance(base_kcal, (int, float)) or base_kcal <= 0:
            raise ValueError(f"Product '{prod.get('id')}' missing or invalid 'base_kcal' > 0")

        nutrients = prod.get("nutrients")
        if not isinstance(nutrients, list) or len(nutrients) == 0:
            raise ValueError(f"Product '{prod.get('id')}' missing or invalid 'nutrients' list")
        for n_idx, nut in enumerate(nutrients):
            if not isinstance(nut, (list, tuple)) or len(nut) < 4:
                raise ValueError(
                    f"Nutrient {n_idx} in '{prod.get('id')}' must be a 4-tuple (name, unit, val_base, val_serving)"
                )

        exemptions = prod.get("clinical_exemptions")
        if exemptions is not None:
            if not isinstance(exemptions, list) or any(not isinstance(e, str) for e in exemptions):
                raise ValueError(f"Product '{prod.get('id')}' 'clinical_exemptions' must be a list of strings")

    return mod.BRAND_META, mod.PRODUCTS


def audit_products(
    products: List[Dict[str, Any]],
    workspace_root: str | Path,
    cohort_min: int = 5,
) -> Dict[str, Any]:
    """
    Run 5 audit checks across the canonical product list and workspace deliverables.

    Returns dictionary containing pass/fail flag and detailed error lists.
    """
    ws_root = Path(workspace_root).resolve()

    ratio_errors: List[str] = []
    macro_errors: List[str] = []
    rdi_errors: List[str] = []
    outlier_errors: List[str] = []
    outlier_skipped: bool = False
    parity_errors: List[str] = []

    macro_keys = {"total_fat", "saturated_fat", "cholesterol", "protein", "total_carbohydrate", "dietary_fiber"}
    excluded_outlier_keys = {"saturated_fat", "cholesterol", "dietary_fiber"}

    # 1. Ratio Errors (Dual-column consistency)
    for p in products:
        p_id = p.get("id", "Unknown")
        expected_ratio: Optional[float] = None
        if p.get("base_value") and p.get("serving_g") and float(p["base_value"]) > 0:
            expected_ratio = float(p["serving_g"]) / float(p["base_value"])
        elif p.get("base_kcal") and p.get("serving_kcal") and float(p["base_kcal"]) > 0:
            expected_ratio = float(p["serving_kcal"]) / float(p["base_kcal"])

        if expected_ratio is not None and expected_ratio > 0:
            for item in p.get("nutrients", []):
                name, unit, val_base, val_serving = item[0], item[1], item[2], item[3]
                name_str = str(name).strip()
                if "น้ำ" in name_str or "water" in name_str.lower():
                    continue
                if val_base is None or val_serving is None or float(val_base) <= 0:
                    continue

                v_base = float(val_base)
                v_serving = float(val_serving)
                expected_val2 = v_base * expected_ratio
                abs_err = abs(v_serving - expected_val2)
                rel_err = abs_err / expected_val2 if expected_val2 > 0 else 0.0

                if rel_err > 0.08 and abs_err > 0.25:
                    ratio_errors.append(
                        f"[{p_id}] {name_str}: base={v_base}, serving={v_serving}, expected={expected_val2:.2f} "
                        f"(rel_err={rel_err:.3f} > 0.08, abs_err={abs_err:.3f} > 0.25)"
                    )

    # 2. Macronutrient Caloric Consistency
    for p in products:
        p_id = p.get("id", "Unknown")
        ener = float(p["base_kcal"])
        p_val: Optional[float] = None
        f_val: Optional[float] = None
        c_val: Optional[float] = None
        fiber_val: float = 0.0

        for item in p.get("nutrients", []):
            name, unit, val_base, _ = item[0], item[1], item[2], item[3]
            if val_base is None:
                continue
            name_str = str(name).strip()
            canon = canonicalize_nutrient_name(name_str)
            if canon == "protein" or (p_val is None and ("โปรตีน" in name_str or "protein" in name_str.lower())):
                p_val = float(val_base)
            elif canon == "total_fat" or (f_val is None and ("ไขมันทั้งหมด" in name_str or "total fat" in name_str.lower())):
                f_val = float(val_base)
            elif canon == "total_carbohydrate" or (c_val is None and ("คาร์โบไฮเดรต" in name_str or "carbohydrate" in name_str.lower())):
                c_val = float(val_base)
            elif canon == "dietary_fiber" or ("ใยอาหาร" in name_str or "fiber" in name_str.lower()):
                fiber_val = float(val_base)

        if p_val is not None and f_val is not None and c_val is not None:
            calc_gross = 4.0 * p_val + 9.0 * f_val + 4.0 * c_val
            calc_net = 4.0 * p_val + 9.0 * f_val + 4.0 * max(0.0, c_val - fiber_val) + 2.0 * fiber_val
            rel_err_gross = abs(calc_gross - ener) / ener
            rel_err_net = abs(calc_net - ener) / ener
            min_rel = min(rel_err_gross, rel_err_net)
            if min_rel > 0.06:
                macro_errors.append(
                    f"[{p_id}] Macro energy mismatch: gross={calc_gross:.1f} kcal, net={calc_net:.1f} kcal vs "
                    f"base_kcal={ener:.1f} kcal (min rel_err={min_rel*100:.1f}% > 6%)"
                )

    # 3. Thai RDI 445 Micronutrient Plausibility [2.0%, 1000.0%]
    for p in products:
        p_id = p.get("id", "Unknown")
        base_kcal = float(p["base_kcal"])
        exemptions = set(p.get("clinical_exemptions") or [])

        for item in p.get("nutrients", []):
            name, unit, val_base, _ = item[0], item[1], item[2], item[3]
            if val_base is None or float(val_base) <= 0:
                continue
            name_str = str(name).strip()
            canon = canonicalize_nutrient_name(name_str)
            if not canon or canon in macro_keys:
                continue
            if canon in exemptions:
                continue
            if canon in THAI_RDI_445_STANDARDS:
                val_1000 = float(val_base) * (1000.0 / base_kcal)
                res = normalize_nutrient(name_str, str(unit), val_1000)
                norm_val = res["normalized_1000kcal"]
                rdi = float(THAI_RDI_445_STANDARDS[canon]["rdi"])
                if norm_val is not None and rdi > 0:
                    pct_rdi = (norm_val / rdi) * 100.0
                    if pct_rdi < 2.0 or pct_rdi > 1000.0:
                        rdi_errors.append(
                            f"[{p_id}] {name_str} ({canon}): normalized={norm_val:.2f} {res['target_unit']}, "
                            f"rdi={rdi}, %rdi={pct_rdi:.1f}% outside plausible bounds [2.0%, 1000.0%]"
                        )

    # 4. Cross-product Cohort Outlier Detection
    if len(products) < cohort_min:
        outlier_skipped = True
        outlier_errors = []
    else:
        outlier_skipped = False
        cohort_values: Dict[str, List[Tuple[Dict[str, Any], float]]] = {}

        for p in products:
            base_kcal = float(p["base_kcal"])
            for item in p.get("nutrients", []):
                name, unit, val_base, _ = item[0], item[1], item[2], item[3]
                if val_base is None or float(val_base) <= 0:
                    continue
                name_str = str(name).strip()
                canon = canonicalize_nutrient_name(name_str)
                if not canon or canon in excluded_outlier_keys:
                    continue
                val_1000 = float(val_base) * (1000.0 / base_kcal)
                res = normalize_nutrient(name_str, str(unit), val_1000)
                norm_val = res["normalized_1000kcal"]
                if norm_val is not None and norm_val > 0:
                    cohort_values.setdefault(canon, []).append((p, norm_val))

        for canon, entries in cohort_values.items():
            if len(entries) >= 5:
                vals = [v for _, v in entries]
                med = statistics.median(vals)
                if med > 0:
                    for p, val in entries:
                        p_id = p.get("id", "Unknown")
                        exemptions = set(p.get("clinical_exemptions") or [])
                        if canon in exemptions:
                            continue
                        ratio = val / med
                        if ratio > 10.0 or ratio < 0.1:
                            outlier_errors.append(
                                f"[{p_id}] Outlier for {canon}: val={val:.2f}, cohort_median={med:.2f}, "
                                f"ratio={ratio:.2f} (outside [0.1, 10.0])"
                            )

    # 5. Deliverables Parity Checks
    for p in products:
        p_id = p.get("id", "Unknown")
        folder = p.get("folder", "")
        p_dir = ws_root / folder

        if not p_dir.is_dir():
            parity_errors.append(f"[{p_id}] Product directory not found: {p_dir}")
            continue

        # Set 2 Markdown verification
        md_files = sorted(p_dir.glob("02_*.md"))
        if not md_files:
            parity_errors.append(f"[{p_id}] Missing Set 2 Markdown (02_*.md) in {p_dir}")
        else:
            try:
                md_content = md_files[0].read_text(encoding="utf-8")
                if "1,000" not in md_content and "1000" not in md_content:
                    parity_errors.append(f"[{p_id}] Set 2 Markdown in {p_dir} missing '1,000' or '1000'")
            except Exception as exc:
                parity_errors.append(f"[{p_id}] Error reading Markdown file {md_files[0]}: {exc}")

        # Standalone Excel verification
        xlsx_files = sorted([f for f in p_dir.glob("*_Nutrition_1000kcal.xlsx") if not f.name.startswith("~$")])
        if not xlsx_files:
            parity_errors.append(f"[{p_id}] Missing *_Nutrition_1000kcal.xlsx in {p_dir}")
        else:
            try:
                wb = openpyxl.load_workbook(str(xlsx_files[0]), data_only=False)
                has_round = False
                for ws in wb.worksheets:
                    for row in ws.iter_rows(values_only=True):
                        for cell_val in row:
                            if isinstance(cell_val, str) and "=ROUND(" in cell_val.upper():
                                has_round = True
                                break
                        if has_round:
                            break
                    if has_round:
                        break
                if not has_round:
                    parity_errors.append(
                        f"[{p_id}] Excel workbook {xlsx_files[0].name} missing '=ROUND(' dynamic formulas"
                    )
            except Exception as exc:
                parity_errors.append(f"[{p_id}] Failed parsing Excel workbook {xlsx_files[0].name}: {exc}")

    passed = (
        len(ratio_errors) == 0
        and len(macro_errors) == 0
        and len(rdi_errors) == 0
        and len(outlier_errors) == 0
        and len(parity_errors) == 0
    )

    return {
        "passed": passed,
        "product_count": len(products),
        "ratio_errors": ratio_errors,
        "macro_errors": macro_errors,
        "rdi_errors": rdi_errors,
        "outlier_errors": outlier_errors,
        "outlier_skipped": outlier_skipped,
        "parity_errors": parity_errors,
    }


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Pipeline Audit Engine for Medical Food Products")
    parser.add_argument("--data", required=True, help="Path to data.py containing BRAND_META and PRODUCTS")
    parser.add_argument("--workspace", required=True, help="Path to workspace root directory")
    parser.add_argument("--cohort-min", type=int, default=5, help="Minimum cohort size for outlier check (default: 5)")
    args = parser.parse_args(argv)

    try:
        meta, products = load_canonical_products(args.data)
    except Exception as exc:
        print(f"Error loading canonical products: {exc}", file=sys.stderr)
        return 1

    res = audit_products(products, args.workspace, cohort_min=args.cohort_min)
    brand_name = meta.get("brand_name", "Unknown")
    print(f"=== Pipeline Audit Report: {brand_name} ({res['product_count']} product(s)) ===")
    print(f"Status: {'PASSED' if res['passed'] else 'FAILED'}")
    if res["outlier_skipped"]:
        print("Cohort Outlier Check: SKIPPED (cohort size < cohort_min)")
    else:
        err_count = len(res["outlier_errors"])
        print(f"Cohort Outlier Check: {'PASSED' if err_count == 0 else f'FAILED ({err_count} errors)'}")

    for err_type in ("ratio_errors", "macro_errors", "rdi_errors", "outlier_errors", "parity_errors"):
        errs = res[err_type]
        if errs:
            print(f"\n--- {err_type.upper()} ({len(errs)}) ---")
            for e in errs:
                print(f"  * {e}")

    return 0 if res["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
