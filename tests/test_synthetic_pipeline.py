# -*- coding: utf-8 -*-
"""
End-to-End Integration Suite for Synthetic Formula Worked Example.

Validates the complete synthetic medical food data pipeline:
  1. Canonical schema & mathematical consistency of synthetic_brand_data.py
  2. build_synthetic_example.py artifact generation (Set 1 MD, Set 2 MD, Standalone Excel, Brand Master Excel)
  3. Dynamic Excel =ROUND(...) formulas and Corporate Blue layout
  4. Idempotent MOPH 445 unit normalization (Markdown columns & dynamic Excel sheets)
  5. Regulatory & clinical audit passage via pipeline_audit.py (API and CLI)
"""

from __future__ import annotations
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

REPO_ROOT = Path(__file__).resolve().parent.parent
EXAMPLE_DIR = REPO_ROOT / "examples" / "synthetic_formula"
SCRIPTS_DIR = REPO_ROOT / "skills" / "medical-food-data-pipeline" / "scripts"

for p in (str(EXAMPLE_DIR), str(SCRIPTS_DIR)):
    if p not in sys.path:
        sys.path.insert(0, p)

# TDD: Will raise ModuleNotFoundError before implementation
import synthetic_brand_data
import build_synthetic_example
from apply_unit_normalization import normalize_workspace
from pipeline_audit import audit_products, load_canonical_products


class TestSyntheticPipeline(unittest.TestCase):
    """End-to-end integration tests for the synthetic formula worked example."""

    def test_01_canonical_schema_and_math_invariants(self):
        """Verify synthetic_brand_data conforms to Canonical Schema and mathematical invariants."""
        # Validate BRAND_META
        meta = getattr(synthetic_brand_data, "BRAND_META", None)
        self.assertIsInstance(meta, dict, "BRAND_META must be a dict")
        self.assertEqual(meta.get("brand_name"), "NutriSample_Co")
        self.assertEqual(meta.get("master_excel"), "NutriSample_Co_Nutrition_1000kcal.xlsx")

        # Validate PRODUCTS
        products = getattr(synthetic_brand_data, "PRODUCTS", None)
        self.assertIsInstance(products, list, "PRODUCTS must be a list")
        self.assertEqual(len(products), 1, "Expected exactly 1 synthetic product")

        p = products[0]
        self.assertEqual(p.get("id"), "Formula_Alpha")
        self.assertEqual(p.get("folder"), "Formula_Alpha")
        self.assertEqual(p.get("base_type"), "powder")
        self.assertEqual(float(p.get("base_value", 0)), 100.0)
        self.assertEqual(float(p.get("base_kcal", 0)), 450.0)
        self.assertAlmostEqual(float(p.get("serving_kcal", 0)), 250.0, places=1)
        self.assertEqual(p.get("clinical_exemptions"), [])

        # Validate dual-column consistency across all nutrients
        expected_ratio = float(p["serving_g"]) / float(p["base_value"])
        for item in p.get("nutrients", []):
            name, unit, v_base, v_serv = item[0], item[1], item[2], item[3]
            if v_base is not None and v_serv is not None and float(v_base) > 0:
                expected_serv = float(v_base) * expected_ratio
                abs_err = abs(float(v_serv) - expected_serv)
                rel_err = abs_err / expected_serv if expected_serv > 0 else 0.0
                self.assertFalse(
                    rel_err > 0.08 and abs_err > 0.25,
                    f"Nutrient ratio error in {name}: base={v_base}, serving={v_serv}, exp={expected_serv:.2f}",
                )

    def test_02_build_synthetic_example_creates_deliverables_in_tempdir(self):
        """Verify build_synthetic_example.build_all generates all 4 deliverables with normalization."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            out_path = build_synthetic_example.build_all(output_dir=tmp_dir)
            self.assertEqual(Path(out_path).resolve(), Path(tmp_dir).resolve())

            # 4 Expected Deliverables
            master_xlsx = Path(tmp_dir) / "NutriSample_Co_Nutrition_1000kcal.xlsx"
            p_dir = Path(tmp_dir) / "Formula_Alpha"
            set1_md = p_dir / "01_Formula_Alpha_Label.md"
            set2_md = p_dir / "02_Formula_Alpha_Ingredients_Nutrition_1000kcal.md"
            standalone_xlsx = p_dir / "Formula_Alpha_Nutrition_1000kcal.xlsx"

            for f in (master_xlsx, set1_md, set2_md, standalone_xlsx):
                self.assertTrue(f.is_file(), f"Expected deliverable missing: {f}")

            # Verify Set 2 Markdown has normalization columns and 1000 kcal
            set2_content = set2_md.read_text(encoding="utf-8")
            self.assertIn("1,000", set2_content)
            self.assertIn("หน่วยมาตรฐาน (Thai RDI 445)", set2_content)
            self.assertIn("ปริมาณมาตรฐานต่อ 1,000 kcal", set2_content)

            # Verify Standalone Excel sheets and dynamic formulas
            wb_std = openpyxl.load_workbook(str(standalone_xlsx), data_only=False)
            self.assertIn("Formula_Alpha", wb_std.sheetnames)
            self.assertIn("Normalized_1000kcal_445", wb_std.sheetnames)
            ws_std = wb_std["Formula_Alpha"]
            has_round_std = any(
                isinstance(cell.value, str) and "=ROUND(" in cell.value.upper()
                for row in ws_std.iter_rows()
                for cell in row
            )
            self.assertTrue(has_round_std, "Standalone Excel must contain '=ROUND(' dynamic formulas")

            # Verify Brand Master Excel sheets and dynamic formulas
            wb_mst = openpyxl.load_workbook(str(master_xlsx), data_only=False)
            self.assertIn("00_ภาพรวมเปรียบเทียบ_1000kcal", wb_mst.sheetnames)
            self.assertIn("01_สรุปพลังงาน_1000kcal", wb_mst.sheetnames)
            self.assertIn("Formula_Alpha", wb_mst.sheetnames)
            self.assertIn("02_Normalized_445", wb_mst.sheetnames)
            ws_mst = wb_mst["Formula_Alpha"]
            has_round_mst = any(
                isinstance(cell.value, str) and "=ROUND(" in cell.value.upper()
                for row in ws_mst.iter_rows()
                for cell in row
            )
            self.assertTrue(has_round_mst, "Master Excel must contain '=ROUND(' dynamic formulas")

    def test_03_checked_in_artifacts_exist_and_pass_audit(self):
        """Verify the checked-in artifacts in output/NutriSample_Co/ exist and pass audit."""
        checked_in_dir = EXAMPLE_DIR / "output" / "NutriSample_Co"
        master_xlsx = checked_in_dir / "NutriSample_Co_Nutrition_1000kcal.xlsx"
        set1_md = checked_in_dir / "Formula_Alpha" / "01_Formula_Alpha_Label.md"
        set2_md = checked_in_dir / "Formula_Alpha" / "02_Formula_Alpha_Ingredients_Nutrition_1000kcal.md"
        standalone_xlsx = checked_in_dir / "Formula_Alpha" / "Formula_Alpha_Nutrition_1000kcal.xlsx"

        for f in (master_xlsx, set1_md, set2_md, standalone_xlsx):
            self.assertTrue(f.is_file(), f"Checked-in deliverable missing: {f}")

        # Audit checked-in deliverables
        meta, prods = load_canonical_products(EXAMPLE_DIR / "synthetic_brand_data.py")
        res = audit_products(prods, checked_in_dir)
        self.assertTrue(res["passed"], f"Checked-in artifacts failed audit: {res}")
        self.assertTrue(res["outlier_skipped"], "Cohort < 5 must skip outliers")

    def test_04_normalization_idempotence_and_second_run(self):
        """Verify normalize_workspace is idempotent and preserves Set 2 table headers."""
        with tempfile.TemporaryDirectory() as tmp_dir:
            build_synthetic_example.build_all(output_dir=tmp_dir)
            set2_md = Path(tmp_dir) / "Formula_Alpha" / "02_Formula_Alpha_Ingredients_Nutrition_1000kcal.md"
            content_before = set2_md.read_text(encoding="utf-8")

            # Run normalization again
            md_ok, xlsx_ok = normalize_workspace(tmp_dir)
            self.assertEqual((md_ok, xlsx_ok), (1, 2))

            content_after = set2_md.read_text(encoding="utf-8")
            # Headers must not have duplicate normalization columns
            hdr_lines_before = [l for l in content_before.splitlines() if "หน่วยมาตรฐาน (Thai RDI 445)" in l]
            hdr_lines_after = [l for l in content_after.splitlines() if "หน่วยมาตรฐาน (Thai RDI 445)" in l]
            self.assertEqual(len(hdr_lines_before), 1)
            self.assertEqual(len(hdr_lines_after), 1)
            self.assertEqual(hdr_lines_before[0], hdr_lines_after[0])

    def test_05_pipeline_audit_cli_subprocess(self):
        """Verify pipeline_audit.py CLI subprocess exits 0 with PASSED."""
        checked_in_dir = EXAMPLE_DIR / "output" / "NutriSample_Co"
        data_py = EXAMPLE_DIR / "synthetic_brand_data.py"
        audit_script = SCRIPTS_DIR / "pipeline_audit.py"

        cmd = [
            sys.executable,
            str(audit_script),
            "--data",
            str(data_py),
            "--workspace",
            str(checked_in_dir),
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True)
        self.assertEqual(
            proc.returncode,
            0,
            f"pipeline_audit.py failed with returncode {proc.returncode}.\nSTDOUT: {proc.stdout}\nSTDERR: {proc.stderr}",
        )
        self.assertIn("Status: PASSED", proc.stdout)
        self.assertIn("Cohort Outlier Check: SKIPPED", proc.stdout)


if __name__ == "__main__":
    unittest.main()
