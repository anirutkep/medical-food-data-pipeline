# -*- coding: utf-8 -*-
"""
Unit tests for Cross-Platform Hybrid OCR & Dual-Pass Tool (ocr_row_extractor.py).
"""

from __future__ import annotations

import base64
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

# Setup paths to import from bundled skill scripts and assets
PKG_ROOT = Path(__file__).resolve().parent.parent
SKILL_ROOT = PKG_ROOT / "skills" / "medical-food-data-pipeline"
SCRIPTS_DIR = SKILL_ROOT / "scripts"
ASSETS_DIR = SKILL_ROOT / "assets"

if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
if str(ASSETS_DIR) not in sys.path:
    sys.path.insert(0, str(ASSETS_DIR))

from ocr_row_extractor import (
    compare_with_data_py,
    compare_with_pass1,
    extract_rows_from_image,
    group_tokens_by_y,
    main,
)
from apply_unit_normalization import (
    NORM_HDR_1,
    NORM_HDR_2,
    detect_workbook_font,
    detect_workbook_type,
    process_excel_workbook,
    process_set2_markdown,
)
from pipeline_audit import (
    audit_products,
    load_canonical_products,
)
from cleanup_workspace import cleanup_workspace


class TestOcrAndAudit(unittest.TestCase):
    """Test token grouping, cross-platform OCR fallbacks, and dual-pass comparisons."""

    def test_group_tokens_by_y_sorts_top_down_left_right(self):
        """Tokens must be grouped by Y within threshold, sorted top-to-bottom and left-to-right."""
        # Tokens format: (yc, xc, text)
        # Apple Vision coordinate system: yc in [0, 1] (0.0 bottom, 1.0 top)
        raw_tokens = [
            (0.500, 0.60, "18.0 g"),
            (0.502, 0.10, "Protein"),
            (0.800, 0.55, "450 kcal"),
            (0.798, 0.05, "Energy"),
            (0.200, 0.10, "Total Fat 15.0 g, 7.5 g"),
        ]

        rows = group_tokens_by_y(raw_tokens, y_thresh=0.012)
        self.assertEqual(len(rows), 3)

        # Row 0: Topmost (Energy / 450 kcal)
        self.assertAlmostEqual(rows[0]["y"], 0.799, places=3)
        self.assertEqual(rows[0]["columns"], ["Energy", "450 kcal"])
        self.assertEqual(rows[0]["numbers"], [450.0])

        # Row 1: Middle (Protein / 18.0 g)
        self.assertAlmostEqual(rows[1]["y"], 0.501, places=3)
        self.assertEqual(rows[1]["columns"], ["Protein", "18.0 g"])
        self.assertEqual(rows[1]["numbers"], [18.0])

        # Row 2: Bottom (Total Fat 15.0 g, 7.5 g)
        self.assertAlmostEqual(rows[2]["y"], 0.200, places=3)
        self.assertEqual(rows[2]["columns"], ["Total Fat 15.0 g, 7.5 g"])
        self.assertEqual(rows[2]["numbers"], [15.0, 7.5])

    def test_winrt_backend_missing_thai_pack_auto_falls_back_cleanly(self):
        """On Windows, missing Thai WinRT OCR pack should auto-fallback to vlm_fallback (exit 0), but explicit winrt raises / exits 1."""
        with tempfile.TemporaryDirectory() as tmpdir:
            dummy_img = Path(tmpdir) / "label.jpg"
            dummy_img.write_bytes(b"dummy image bytes")

            # Mock platform to win32 and mock powershell reporting MISSING_THAI_PACK
            missing_pack_json = json.dumps({
                "error": "MISSING_THAI_PACK",
                "message": "Windows WinRT Thai (th) OCR language pack not installed.",
            })

            mock_proc = MagicMock()
            mock_proc.returncode = 0
            mock_proc.stdout = missing_pack_json
            mock_proc.stderr = ""

            with patch("sys.platform", "win32"), \
                 patch("shutil.which", return_value="powershell.exe"), \
                 patch("subprocess.run", return_value=mock_proc):

                # 1. Programmatic --backend auto falls through to vlm_fallback without throwing
                # (assuming tesseract is also not available in this test environment)
                with patch("ocr_row_extractor._run_tesseract_ocr", side_effect=RuntimeError("No tesseract")):
                    res = extract_rows_from_image(dummy_img, backend="auto")
                    self.assertEqual(res["status"], "vlm_fallback")
                    self.assertEqual(res["backend"], "vlm_fallback")
                    self.assertIn("instructions", res)

                    # CLI with --backend auto exits 0
                    exit_code = main([str(dummy_img), "--backend", "auto"])
                    self.assertEqual(exit_code, 0)

                # 2. Programmatic --backend winrt raises RuntimeError
                with self.assertRaises(RuntimeError) as ctx:
                    extract_rows_from_image(dummy_img, backend="winrt")
                self.assertIn("Windows WinRT Thai (th) OCR language pack not installed", str(ctx.exception))

                # CLI with --backend winrt exits 1
                exit_code_winrt = main([str(dummy_img), "--backend", "winrt"])
                self.assertEqual(exit_code_winrt, 1)

    def test_winrt_powershell_encoded_command_and_utf8_env_for_thai_path(self):
        """WinRT runner must pass -EncodedCommand base64 and preserve exact Thai filenames in OCR_IMAGE_PATH env."""
        with tempfile.TemporaryDirectory() as tmpdir:
            thai_filename = "อาหารทางการแพทย์ ตัวอย่าง.jpg"
            thai_img = Path(tmpdir) / thai_filename
            thai_img.write_bytes(b"dummy image bytes")

            success_ocr_json = json.dumps({
                "width": 1000,
                "height": 2000,
                "words": [
                    {"x": 100, "y": 200, "w": 50, "h": 20, "text": "450"},
                    {"x": 50, "y": 200, "w": 40, "h": 20, "text": "พลังงาน"},
                ],
            })

            mock_proc = MagicMock()
            mock_proc.returncode = 0
            mock_proc.stdout = success_ocr_json
            mock_proc.stderr = ""

            with patch("sys.platform", "win32"), \
                 patch("shutil.which", return_value="powershell.exe"), \
                 patch("subprocess.run", return_value=mock_proc) as mock_run:

                from ocr_row_extractor import _run_winrt_ocr
                rows = _run_winrt_ocr(thai_img, y_thresh=0.012)
                self.assertEqual(len(rows), 1)
                self.assertEqual(rows[0]["columns"], ["พลังงาน", "450"])
                self.assertEqual(rows[0]["numbers"], [450.0])

                # Verify subprocess.run arguments
                self.assertTrue(mock_run.called)
                cmd_args = mock_run.call_args[0][0]
                self.assertIn("-EncodedCommand", cmd_args)
                encoded_idx = cmd_args.index("-EncodedCommand") + 1
                b64_str = cmd_args[encoded_idx]

                # Assert base64 decodes cleanly from utf-16le
                decoded_script = base64.b64decode(b64_str).decode("utf-16le")
                self.assertIn("OCR_IMAGE_PATH", decoded_script)

                # Assert OCR_IMAGE_PATH environment variable preserved exact Thai path
                env_passed = mock_run.call_args[1].get("env", {})
                self.assertIn("OCR_IMAGE_PATH", env_passed)
                self.assertEqual(env_passed["OCR_IMAGE_PATH"], str(thai_img.resolve()))
                self.assertIn(thai_filename, env_passed["OCR_IMAGE_PATH"])

    def test_compare_json_accepts_dict_and_list_roundtrip(self):
        """compare_with_pass1 must accept flat list, dict with 'rows', and format json map."""
        rows = [{"y": 0.8, "columns": ["Energy", "450"], "numbers": [450.0]}]

        with tempfile.TemporaryDirectory() as tmpdir:
            # 1. Flat list: [{"numbers": [450.0]}]
            f1 = Path(tmpdir) / "pass1_list.json"
            f1.write_text(json.dumps([{"numbers": [450.0]}]), encoding="utf-8")
            self.assertEqual(compare_with_pass1(rows, f1), [])

            # 2. Top-level dict with 'rows': {"status": "ok", "rows": [{"numbers": [450.0]}]}
            f2 = Path(tmpdir) / "pass1_dict_rows.json"
            f2.write_text(json.dumps({"status": "ok", "rows": [{"numbers": [450.0]}]}), encoding="utf-8")
            self.assertEqual(compare_with_pass1(rows, f2), [])

            # 3. Format json mapping: {"img.jpg": {"status": "ok", "rows": [{"numbers": [450.0]}]}}
            f3 = Path(tmpdir) / "pass1_map_rows.json"
            f3.write_text(json.dumps({"img.jpg": {"status": "ok", "rows": [{"numbers": [450.0]}]}}), encoding="utf-8")
            self.assertEqual(compare_with_pass1(rows, f3), [])

            # 4. Format json mapping with direct list: {"img.jpg": [{"numbers": [450.0]}]}
            f4 = Path(tmpdir) / "pass1_map_list.json"
            f4.write_text(json.dumps({"img.jpg": [{"numbers": [450.0]}]}), encoding="utf-8")
            self.assertEqual(compare_with_pass1(rows, f4), [])

            # 5. Mismatch detected
            diff_rows = [{"y": 0.8, "columns": ["Energy", "999"], "numbers": [999.0]}]
            diffs = compare_with_pass1(diff_rows, f1)
            self.assertTrue(len(diffs) > 0)
            self.assertTrue(any("999" in d for d in diffs))

            # 6. Unrecognized structure raises ValueError
            f_invalid = Path(tmpdir) / "invalid.json"
            f_invalid.write_text(json.dumps({"unexpected_key": 12345}), encoding="utf-8")
            with self.assertRaises(ValueError):
                compare_with_pass1(rows, f_invalid)

    def test_compare_with_data_py_valid_and_invalid_schema(self):
        """compare_with_data_py must validate PRODUCTS contract, support filtering, and detect diffs."""
        with tempfile.TemporaryDirectory() as tmpdir:
            dummy_img = Path(tmpdir) / "label.jpg"
            dummy_img.write_bytes(b"dummy")

            # 1. Valid data.py
            valid_py = Path(tmpdir) / "valid_data.py"
            valid_py.write_text(
                'BRAND_META = {"brand_name": "TestBrand", "master_excel": "test.xlsx"}\n'
                'PRODUCTS = [\n'
                '    {\n'
                '        "id": "P_Standard",\n'
                '        "nutrients": [\n'
                '            ("Energy", "kcal", 450.0, 225.0),\n'
                '            ("Protein", "g", 18.0, 9.0),\n'
                '        ],\n'
                '    }\n'
                ']\n',
                encoding="utf-8",
            )

            matching_rows = [
                {"y": 0.8, "columns": ["Energy 450 / 225"], "numbers": [450.0, 225.0]},
                {"y": 0.5, "columns": ["Protein 18.0 / 9.0"], "numbers": [18.0, 9.0]},
            ]
            self.assertEqual(compare_with_data_py(matching_rows, valid_py), [])
            self.assertEqual(compare_with_data_py(matching_rows, valid_py, product_id="P_Standard"), [])

            # Product ID not found raises ValueError
            with self.assertRaises(ValueError):
                compare_with_data_py(matching_rows, valid_py, product_id="NoSuchProduct")

            # Mismatching rows returns diffs
            mismatch_rows = [
                {"y": 0.8, "columns": ["Energy 999 / 225"], "numbers": [999.0, 225.0]},
            ]
            diffs = compare_with_data_py(mismatch_rows, valid_py)
            self.assertTrue(len(diffs) > 0)

            # CLI with matching rows exits 0
            with patch("ocr_row_extractor.extract_rows_from_image", return_value={"status": "ok", "backend": "mock", "rows": matching_rows}):
                code = main([str(dummy_img), "--data-py", str(valid_py)])
                self.assertEqual(code, 0)

            # CLI with mismatching rows exits 1
            with patch("ocr_row_extractor.extract_rows_from_image", return_value={"status": "ok", "backend": "mock", "rows": mismatch_rows}):
                code_diff = main([str(dummy_img), "--data-py", str(valid_py)])
                self.assertEqual(code_diff, 1)

            # 2. Invalid data.py missing PRODUCTS raises ValueError and CLI exits 1
            missing_products_py = Path(tmpdir) / "missing_products.py"
            missing_products_py.write_text('BRAND_META = {"brand_name": "TestBrand"}\n', encoding="utf-8")
            with self.assertRaises(ValueError):
                compare_with_data_py(matching_rows, missing_products_py)

            code_invalid = main([str(dummy_img), "--data-py", str(missing_products_py)])
            self.assertEqual(code_invalid, 1)

            # 3. Empty PRODUCTS list raises ValueError
            empty_products_py = Path(tmpdir) / "empty_products.py"
            empty_products_py.write_text('PRODUCTS = []\n', encoding="utf-8")
            with self.assertRaises(ValueError):
                compare_with_data_py(matching_rows, empty_products_py)

            # 4. Product missing nutrients raises ValueError
            missing_nutrients_py = Path(tmpdir) / "missing_nutrients.py"
            missing_nutrients_py.write_text('PRODUCTS = [{"id": "P1"}]\n', encoding="utf-8")
            with self.assertRaises(ValueError):
                compare_with_data_py(matching_rows, missing_nutrients_py)

    def test_missing_image_file_raises_and_cli_exits_1(self):
        """Missing image file must raise FileNotFoundError and CLI must return 1."""
        non_existent = "non_existent_label_xyz123.jpg"
        with self.assertRaises(FileNotFoundError):
            extract_rows_from_image(non_existent)

        code = main([non_existent])
        self.assertEqual(code, 1)

    def test_invalid_backend_raises(self):
        """Invalid backend name must raise ValueError."""
        with tempfile.TemporaryDirectory() as tmpdir:
            img = Path(tmpdir) / "test.jpg"
            img.write_bytes(b"data")
            with self.assertRaises(ValueError):
                extract_rows_from_image(img, backend="invalid_backend")

    def test_vlm_fallback_backend_returns_instructions(self):
        """Explicit backend='vlm_fallback' returns status='vlm_fallback' with instructions."""
        with tempfile.TemporaryDirectory() as tmpdir:
            img = Path(tmpdir) / "test.jpg"
            img.write_bytes(b"data")
            res = extract_rows_from_image(img, backend="vlm_fallback")
            self.assertEqual(res["status"], "vlm_fallback")
            self.assertEqual(res["backend"], "vlm_fallback")
            self.assertEqual(res["rows"], [])
            self.assertIn("instructions", res)

            # CLI with --backend vlm_fallback returns exit code 0
            code = main([str(img), "--backend", "vlm_fallback"])
            self.assertEqual(code, 0)


    def test_cli_formats_table_markdown_json(self):
        """CLI formats --format json, markdown, and table must render accurately."""
        with tempfile.TemporaryDirectory() as tmpdir:
            img = Path(tmpdir) / "test.jpg"
            img.write_bytes(b"dummy")
            sample_rows = [
                {"y": 0.800, "columns": ["Energy", "450"], "numbers": [450.0]}
            ]
            mock_res = {"status": "ok", "backend": "mock", "image": str(img), "rows": sample_rows}
            with patch("ocr_row_extractor.extract_rows_from_image", return_value=mock_res):
                import io
                from contextlib import redirect_stdout

                # JSON format
                f_json = io.StringIO()
                with redirect_stdout(f_json):
                    code = main([str(img), "--format", "json"])
                self.assertEqual(code, 0)
                parsed = json.loads(f_json.getvalue())
                self.assertIn(str(img), parsed)
                self.assertEqual(parsed[str(img)]["status"], "ok")
                self.assertEqual(parsed[str(img)]["rows"], sample_rows)

                # Markdown format
                f_md = io.StringIO()
                with redirect_stdout(f_md):
                    code = main([str(img), "--format", "markdown"])
                self.assertEqual(code, 0)
                self.assertIn("| Y | Columns |", f_md.getvalue())
                self.assertIn("Energy / 450", f_md.getvalue())

                # Table format
                f_tbl = io.StringIO()
                with redirect_stdout(f_tbl):
                    code = main([str(img), "--format", "table"])
                self.assertEqual(code, 0)
                self.assertIn("[y=0.800] Energy | 450", f_tbl.getvalue())

    def test_apply_normalization_detects_master_vs_standalone_and_is_idempotent(self):
        """detect_workbook_type and process_excel_workbook must handle standalone/master, inherit font, and be idempotent."""
        import openpyxl
        from openpyxl.styles import Font

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)

            # Incomplete MD without Section 4 Nutrition table
            incomplete_md = tmp_path / "02_incomplete.md"
            incomplete_md.write_text("# Product B\n## 1. Description\nNo nutrition table.\n", encoding="utf-8")
            self.assertFalse(process_set2_markdown(incomplete_md))
            self.assertIn("No nutrition table", incomplete_md.read_text(encoding="utf-8"))

            # Complete Set 2 MD with Section 4 Nutrition table
            complete_md = tmp_path / "02_complete.md"
            md_content = (
                "# Product A\n"
                "## 2. ส่วนประกอบ\n"
                "น้ำ 80%, เวย์โปรตีน 10%\n\n"
                "## 4. ข้อมูลโภชนาการ (Nutrition Facts)\n"
                "| ลำดับ | รายการสารอาหาร | หน่วย | ต่อ 100 g | ต่อ 1,000 kcal |\n"
                "| :---: | :--- | :---: | :---: | :---: |\n"
                "| 1 | พลังงาน (Energy) | kcal | 450 | 1000 |\n"
                "| 2 | โปรตีน (Protein) | g | 20 | 44.44 |\n"
                "| 3 | วิตามินเอ (Vitamin A) | IU | 800 | 1777.78 |\n"
            )
            complete_md.write_text(md_content, encoding="utf-8")
            self.assertTrue(process_set2_markdown(complete_md))
            text1 = complete_md.read_text(encoding="utf-8")
            self.assertIn(NORM_HDR_1, text1)
            self.assertIn(NORM_HDR_2, text1)
            self.assertIn("<!-- UNIT_NORMALIZATION_NOTE_445 -->", text1)

            # Run second time -> idempotent (no duplicated columns)
            self.assertTrue(process_set2_markdown(complete_md))
            text2 = complete_md.read_text(encoding="utf-8")
            self.assertEqual(text2.count(NORM_HDR_1), 1)
            self.assertEqual(text2.count(NORM_HDR_2), 1)
            self.assertEqual(text2.count("<!-- UNIT_NORMALIZATION_NOTE_445 -->"), 2)

            # Standalone Excel workbook
            standalone_path = tmp_path / "Sample_Nutrition_1000kcal.xlsx"
            wb_stand = openpyxl.Workbook()
            ws_prod = wb_stand.active
            ws_prod.title = "Sample_Product"
            # Style cells with Cordia New
            c_font = Font(name="Cordia New", size=14)
            ws_prod["A4"] = "ลำดับ"
            ws_prod["B4"] = "รายการสารอาหาร"
            ws_prod["C4"] = "หน่วย"
            ws_prod["D4"] = "ต่อ 100 g"
            ws_prod["E4"] = "ต่อ 1,000 kcal"
            ws_prod["B5"] = "โปรตีน (Protein)"
            ws_prod["C5"] = "g"
            ws_prod["E5"] = "=D5*(1000/450)"
            for row in ws_prod.iter_rows(min_row=4, max_row=5, min_col=1, max_col=5):
                for cell in row:
                    cell.font = c_font
            wb_stand.save(standalone_path)

            self.assertEqual(detect_workbook_type(wb_stand), "standalone")
            self.assertEqual(detect_workbook_font(wb_stand), "Cordia New")
            self.assertTrue(process_excel_workbook(standalone_path))

            wb_loaded = openpyxl.load_workbook(standalone_path, data_only=False)
            self.assertIn("Normalized_1000kcal_445", wb_loaded.sheetnames)
            ws_norm = wb_loaded["Normalized_1000kcal_445"]
            # Verify font inheritance
            self.assertEqual(ws_norm["A1"].font.name, "Cordia New")
            # Verify formulas
            self.assertIn("ROUND(", str(ws_norm["D5"].value))
            self.assertIn("ROUND(", str(ws_norm["G5"].value))

            # Master Excel workbook
            master_path = tmp_path / "Master_Nutrition_1000kcal.xlsx"
            wb_master = openpyxl.Workbook()
            ws_master_prod = wb_master.active
            ws_master_prod.title = "Sample_Product"
            ws_master_prod["B4"] = "สารอาหาร"
            ws_master_prod["C4"] = "หน่วย"
            ws_master_prod["D4"] = "ต่อ 1,000 kcal"
            ws_master_prod["B5"] = "แคลเซียม (Calcium)"
            ws_master_prod["C5"] = "mg"
            ws_master_prod["D5"] = "=100*10"
            for row in ws_master_prod.iter_rows(min_row=4, max_row=5, min_col=1, max_col=4):
                for cell in row:
                    cell.font = c_font
            wb_master.create_sheet("00_ภาพรวมเปรียบเทียบ_1000kcal")
            wb_master.save(master_path)

            self.assertEqual(detect_workbook_type(wb_master), "master")
            self.assertTrue(process_excel_workbook(master_path))

            wb_master_loaded = openpyxl.load_workbook(master_path, data_only=False)
            self.assertIn("02_Normalized_445", wb_master_loaded.sheetnames)
            ws_m_norm = wb_master_loaded["02_Normalized_445"]
            self.assertEqual(ws_m_norm["A1"].font.name, "Cordia New")

            # Running twice on standalone is idempotent
            self.assertTrue(process_excel_workbook(standalone_path))
            wb_reloaded = openpyxl.load_workbook(standalone_path, data_only=False)
            self.assertEqual(wb_reloaded.sheetnames.count("Normalized_1000kcal_445"), 1)

    def test_pipeline_audit_small_cohort_skips_outliers_cleanly(self):
        """Small cohort (< cohort_min) must skip outliers cleanly and pass valid product."""
        import openpyxl

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            prod_dir = tmp_path / "TestBrand" / "Prod1"
            prod_dir.mkdir(parents=True)

            # Create Set 2 Markdown
            md_file = prod_dir / "02_Prod1_Nutrition.md"
            md_file.write_text("# Prod 1\nพลังงานต่อ 1,000 kcal: 1000 kcal\n", encoding="utf-8")

            # Create Standalone Excel
            xlsx_file = prod_dir / "Prod1_Nutrition_1000kcal.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Nutrition"
            ws["A1"] = "=ROUND(100.555, 2)"
            wb.save(xlsx_file)

            # Create data.py
            data_py = tmp_path / "test_data.py"
            data_py.write_text(
                'BRAND_META = {"brand_name": "TestBrand", "master_excel": "TestBrand_Nutrition_1000kcal.xlsx"}\n'
                'PRODUCTS = [\n'
                '    {\n'
                '        "id": "Prod1",\n'
                '        "folder": "TestBrand/Prod1",\n'
                '        "trade_th": "สินค้า 1",\n'
                '        "trade_en": "Product 1",\n'
                '        "base_type": "powder",\n'
                '        "base_value": 100.0,\n'
                '        "base_kcal": 450.0,\n'
                '        "serving_g": 50.0,\n'
                '        "serving_kcal": 225.0,\n'
                '        "nominal_pcf": (18, 32, 50),\n'
                '        "clinical_exemptions": [],\n'
                '        "ingredients": [("Whey", 100.0, "Protein")],\n'
                '        "nutrients": [\n'
                '            ("พลังงาน (Energy)", "kcal", 450.0, 225.0),\n'
                '            ("โปรตีน (Protein)", "g", 20.25, 10.125),\n'
                '            ("ไขมันทั้งหมด (Total Fat)", "g", 16.0, 8.0),\n'
                '            ("คาร์โบไฮเดรตทั้งหมด (Total Carbohydrate)", "g", 56.25, 28.125),\n'
                '            ("แคลเซียม (Calcium)", "mg", 450.0, 225.0),\n'
                '            ("วิตามินซี (Vitamin C)", "mg", 45.0, 22.5),\n'
                '        ],\n'
                '    }\n'
                ']\n',
                encoding="utf-8",
            )

            meta, products = load_canonical_products(data_py)
            self.assertEqual(meta["brand_name"], "TestBrand")
            self.assertEqual(len(products), 1)

            res = audit_products(products, workspace_root=tmp_path, cohort_min=5)
            self.assertEqual(res["product_count"], 1)
            self.assertTrue(res["outlier_skipped"])
            self.assertEqual(res["outlier_errors"], [])
            self.assertEqual(res["ratio_errors"], [])
            self.assertEqual(res["macro_errors"], [])
            self.assertEqual(res["rdi_errors"], [])
            self.assertEqual(res["parity_errors"], [])
            self.assertTrue(res["passed"])

    def test_pipeline_audit_large_cohort_flags_outlier_unless_exempt(self):
        """Cohort of >=5 products flags 12x outlier unless listed in clinical_exemptions."""
        import openpyxl

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)

            products = []
            for i in range(1, 6):
                p_id = f"Prod_{i}"
                p_folder = f"TestBrand/{p_id}"
                p_dir = tmp_path / p_folder
                p_dir.mkdir(parents=True)

                md_file = p_dir / f"02_{p_id}_Nutrition.md"
                md_file.write_text(f"# {p_id}\n1,000 kcal reference\n", encoding="utf-8")

                xlsx_file = p_dir / f"{p_id}_Nutrition_1000kcal.xlsx"
                wb = openpyxl.Workbook()
                ws = wb.active
                ws["A1"] = "=ROUND(12.34, 1)"
                wb.save(xlsx_file)

                # Product 5 has 12x Potassium (4200 mg vs 350 mg)
                potassium_val = 4200.0 if i == 5 else 350.0
                serving_potassium = potassium_val * 0.5

                p_dict = {
                    "id": p_id,
                    "folder": p_folder,
                    "trade_th": f"สินค้า {i}",
                    "trade_en": f"Product {i}",
                    "base_type": "powder",
                    "base_value": 100.0,
                    "base_kcal": 450.0,
                    "serving_g": 50.0,
                    "serving_kcal": 225.0,
                    "nominal_pcf": (18, 32, 50),
                    "clinical_exemptions": [],
                    "nutrients": [
                        ("พลังงาน (Energy)", "kcal", 450.0, 225.0),
                        ("โปรตีน (Protein)", "g", 20.25, 10.125),
                        ("ไขมันทั้งหมด (Total Fat)", "g", 16.0, 8.0),
                        ("คาร์โบไฮเดรตทั้งหมด (Total Carbohydrate)", "g", 56.25, 28.125),
                        ("โพแทสเซียม (Potassium)", "mg", potassium_val, serving_potassium),
                    ],
                }
                products.append(p_dict)

            # Case A: Without exemption -> fails with outlier_errors
            res_a = audit_products(products, workspace_root=tmp_path, cohort_min=5)
            self.assertFalse(res_a["outlier_skipped"])
            self.assertTrue(len(res_a["outlier_errors"]) > 0)
            self.assertTrue(any("potassium" in str(e).lower() or "โพแทสเซียม" in str(e) for e in res_a["outlier_errors"]))
            self.assertFalse(res_a["passed"])

            # Case B: With exemption on Product 5 -> passes cleanly
            products[4]["clinical_exemptions"] = ["potassium"]
            res_b = audit_products(products, workspace_root=tmp_path, cohort_min=5)
            self.assertFalse(res_b["outlier_skipped"])
            self.assertEqual(res_b["outlier_errors"], [])
            self.assertTrue(res_b["passed"])

    def test_cleanup_workspace_removes_cache_and_os_junk(self):
        """cleanup_workspace removes __pycache__, .ocr_cache, .DS_Store, Thumbs.db and preserves files."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            pycache_dir = tmp_path / "__pycache__"
            pycache_dir.mkdir()
            (pycache_dir / "x.pyc").write_bytes(b"pyc")

            ocr_cache_dir = tmp_path / ".ocr_cache"
            ocr_cache_dir.mkdir()
            (ocr_cache_dir / "o.txt").write_text("cache", encoding="utf-8")

            ds_store = tmp_path / ".DS_Store"
            ds_store.write_bytes(b"ds")

            thumbs = tmp_path / "Thumbs.db"
            thumbs.write_bytes(b"thumbs")

            keep_file = tmp_path / "keep.py"
            keep_file.write_text("print('keep')", encoding="utf-8")

            # Dry run: detects all 4 junk paths without deleting
            dry_targets = cleanup_workspace(tmp_path, dry_run=True)
            self.assertEqual(len(dry_targets), 4)
            self.assertTrue(pycache_dir.exists())
            self.assertTrue(ocr_cache_dir.exists())
            self.assertTrue(ds_store.exists())
            self.assertTrue(thumbs.exists())
            self.assertTrue(keep_file.exists())

            # Real run: deletes all 4 junk paths and preserves keep.py
            real_targets = cleanup_workspace(tmp_path, dry_run=False)
            self.assertEqual(len(real_targets), 4)
            self.assertFalse(pycache_dir.exists())
            self.assertFalse(ocr_cache_dir.exists())
            self.assertFalse(ds_store.exists())
            self.assertFalse(thumbs.exists())
            self.assertTrue(keep_file.exists())


if __name__ == "__main__":
    unittest.main()
