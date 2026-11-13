# -*- coding: utf-8 -*-
"""
Unit tests for Unit Normalizer (MOPH Notification No. 445) & Canonical SSOT Template.
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

# Setup paths to import from bundled skill scripts and assets
PKG_ROOT = Path(__file__).resolve().parent.parent
SKILL_ROOT = PKG_ROOT / "skills" / "medical-food-data-pipeline"
SCRIPTS_DIR = SKILL_ROOT / "scripts"
ASSETS_DIR = SKILL_ROOT / "assets"

if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))
if str(ASSETS_DIR) not in sys.path:
    sys.path.insert(0, str(ASSETS_DIR))

from unit_normalizer import (
    THAI_RDI_445_STANDARDS,
    canonicalize_nutrient_name,
    normalize_unit_token,
    get_conversion_factor,
    normalize_nutrient,
)
from brand_data_template import (
    BRAND_META,
    PRODUCTS,
)


class TestUnitNormalizer(unittest.TestCase):
    """Test Thai RDI 33 standards, conversion factors, SI scaling, and canonical template contract."""

    def test_thai_rdi_33_standards_ascii_safe(self):
        """Validate all 33 Thai RDI standards use ASCII-safe units and positive RDI values."""
        self.assertEqual(len(THAI_RDI_445_STANDARDS), 33)
        allowed_units = {"g", "mg", "mcg", "mcg RAE", "mg alpha-TE", "mg NE", "mcg DFE"}
        numbers_seen = set()

        for key, meta in THAI_RDI_445_STANDARDS.items():
            self.assertIn("no", meta, f"Missing 'no' in {key}")
            self.assertIn("th", meta, f"Missing 'th' in {key}")
            self.assertIn("en", meta, f"Missing 'en' in {key}")
            self.assertIn("rdi", meta, f"Missing 'rdi' in {key}")
            self.assertIn("unit", meta, f"Missing 'unit' in {key}")

            self.assertIn(
                meta["unit"],
                allowed_units,
                f"Non-ASCII or invalid unit in {key}: {meta['unit']}",
            )
            self.assertGreater(meta["rdi"], 0, f"RDI must be positive in {key}")
            self.assertIsInstance(meta["no"], int)
            numbers_seen.add(meta["no"])

        self.assertEqual(numbers_seen, set(range(1, 34)))

    def test_section_2_5_vitamin_conversions(self):
        """Test Section 2.5 conversions: Vit A, Vit D, Vit E (d- vs dl-), Niacin, and Folate."""
        # Vitamin A: 1 IU = 0.3 mcg RAE; 1 mcg RE = 1 mcg RAE
        vit_a_iu = normalize_nutrient("วิตามินเอ (Vitamin A)", "หน่วยสากล (IU)", 1000.0)
        self.assertEqual(vit_a_iu["canonical_key"], "vitamin_a")
        self.assertEqual(vit_a_iu["target_unit"], "mcg RAE")
        self.assertAlmostEqual(vit_a_iu["normalized_1000kcal"], 300.0, places=2)

        vit_a_re = normalize_nutrient("1. วิตามินเอ", "ไมโครกรัม RE", 800.0)
        self.assertEqual(vit_a_re["canonical_key"], "vitamin_a")
        self.assertEqual(vit_a_re["target_unit"], "mcg RAE")
        self.assertAlmostEqual(vit_a_re["normalized_1000kcal"], 800.0, places=2)

        # Vitamin D: 40 IU = 1 mcg
        vit_d = normalize_nutrient("วิตามินดี (Vitamin D)", "IU", 600.0)
        self.assertEqual(vit_d["canonical_key"], "vitamin_d")
        self.assertEqual(vit_d["target_unit"], "mcg")
        self.assertAlmostEqual(vit_d["normalized_1000kcal"], 15.0, places=2)

        # Vitamin E: d-alpha (1.5 IU/mg) vs dl-alpha (2.22 IU/mg)
        vit_e_d = normalize_nutrient(
            "วิตามินอี (Vitamin E)",
            "IU",
            15.0,
            ingredients_text="d-alpha-tocopheryl acetate",
        )
        self.assertEqual(vit_e_d["canonical_key"], "vitamin_e")
        self.assertEqual(vit_e_d["target_unit"], "mg alpha-TE")
        self.assertAlmostEqual(vit_e_d["normalized_1000kcal"], 10.0, places=2)

        vit_e_dl = normalize_nutrient(
            "วิตามินอี (Vitamin E)",
            "IU",
            22.2,
            ingredients_text="dl-alpha-tocopheryl acetate",
        )
        self.assertEqual(vit_e_dl["canonical_key"], "vitamin_e")
        self.assertEqual(vit_e_dl["target_unit"], "mg alpha-TE")
        self.assertAlmostEqual(vit_e_dl["normalized_1000kcal"], 10.0, places=2)

        # Niacin: mcg -> mg NE, mg -> mg NE
        niacin_mcg = normalize_nutrient("ไนอะซิน (Niacin)", "ไมโครกรัม", 24978.73)
        self.assertEqual(niacin_mcg["canonical_key"], "niacin")
        self.assertEqual(niacin_mcg["target_unit"], "mg NE")
        self.assertAlmostEqual(niacin_mcg["normalized_1000kcal"], 24.98, places=2)

        niacin_mg = normalize_nutrient("กรดนิโคตินิก / ไนอะซิน (Nicotinic Acid)", "mg", 6.67)
        self.assertEqual(niacin_mg["canonical_key"], "niacin")
        self.assertEqual(niacin_mg["target_unit"], "mg NE")
        self.assertAlmostEqual(niacin_mg["normalized_1000kcal"], 6.67, places=2)

        # Folate: 0.6 mcg folic acid added = 1 mcg DFE
        folate = normalize_nutrient("กรดโฟลิก (Folic acid)", "ไมโครกรัม", 240.0)
        self.assertEqual(folate["canonical_key"], "folate")
        self.assertEqual(folate["target_unit"], "mcg DFE")
        self.assertAlmostEqual(folate["normalized_1000kcal"], 400.0, places=2)

    def test_si_scaling_and_exclusions(self):
        """Test SI scaling (Copper, Manganese) and structural exclusions (CaHMB, HMB, etc.)."""
        # Copper: mg -> mcg (x1000)
        cu = normalize_nutrient("ทองแดง (Copper)", "mg", 0.9)
        self.assertEqual(cu["canonical_key"], "copper")
        self.assertEqual(cu["target_unit"], "mcg")
        self.assertAlmostEqual(cu["normalized_1000kcal"], 900.0, places=2)

        # Manganese: mcg -> mg (/1000)
        mn = normalize_nutrient("แมงกานีส (Manganese)", "ไมโครกรัม", 3000.0)
        self.assertEqual(mn["canonical_key"], "manganese")
        self.assertEqual(mn["target_unit"], "mg")
        self.assertAlmostEqual(mn["normalized_1000kcal"], 3.0, places=2)

        # Exclusions: CaHMB / HMB must NOT map to Calcium
        self.assertIsNone(canonicalize_nutrient_name("แคลเซียม เอช เอ็ม บี (CaHMB)"))
        self.assertIsNone(canonicalize_nutrient_name("CaHMB"))
        self.assertIsNone(canonicalize_nutrient_name("HMB"))
        self.assertIsNone(canonicalize_nutrient_name("เอช เอ็ม บี"))
        self.assertIsNone(canonicalize_nutrient_name("เบตา-แคโรทีน (Beta-carotene)"))
        self.assertIsNone(canonicalize_nutrient_name("พลังงานจากไขมัน (Energy from fat)"))
        self.assertIsNone(canonicalize_nutrient_name("ไขมันทรานส์ (Trans fat)"))

    def test_canonical_template_contract(self):
        """Verify brand_data_template.py satisfies all canonical SSOT schema requirements."""
        # 1. BRAND_META contract
        self.assertIsInstance(BRAND_META, dict)
        self.assertIn("brand_name", BRAND_META)
        self.assertIn("master_excel", BRAND_META)
        self.assertIsInstance(BRAND_META["brand_name"], str)
        self.assertIsInstance(BRAND_META["master_excel"], str)
        self.assertTrue(
            BRAND_META["master_excel"].endswith(".xlsx"),
            "master_excel must be an .xlsx filename",
        )

        # 2. PRODUCTS list contract
        self.assertIsInstance(PRODUCTS, list)
        self.assertGreater(len(PRODUCTS), 0, "PRODUCTS list must not be empty")

        required_keys = {
            "id",
            "folder",
            "trade_th",
            "trade_en",
            "base_type",
            "base_value",
            "base_kcal",
            "serving_g",
            "serving_kcal",
            "nominal_pcf",
            "clinical_exemptions",
            "ingredients",
            "nutrients",
        }

        for p in PRODUCTS:
            # Check presence of all required keys
            missing = required_keys - set(p.keys())
            self.assertEqual(missing, set(), f"Product {p.get('id')} missing keys: {missing}")

            # Basic types and values
            self.assertIsInstance(p["id"], str)
            self.assertTrue(p["id"].strip())
            self.assertIsInstance(p["folder"], str)
            self.assertTrue(p["folder"].strip())
            self.assertIsInstance(p["trade_th"], str)
            self.assertIsInstance(p["trade_en"], str)
            self.assertIn(p["base_type"], ("powder", "rtd_liquid"))

            self.assertIsInstance(p["base_value"], (int, float))
            self.assertGreater(p["base_value"], 0)
            self.assertIsInstance(p["base_kcal"], (int, float))
            self.assertGreater(p["base_kcal"], 0)
            self.assertIsInstance(p["serving_g"], (int, float))
            self.assertGreater(p["serving_g"], 0)
            self.assertIsInstance(p["serving_kcal"], (int, float))
            self.assertGreater(p["serving_kcal"], 0)

            # nominal_pcf: 3 numbers summing to 100
            self.assertIsInstance(p["nominal_pcf"], (tuple, list))
            self.assertEqual(len(p["nominal_pcf"]), 3)
            self.assertAlmostEqual(sum(p["nominal_pcf"]), 100.0, places=1)

            # clinical_exemptions: List[str]
            self.assertIsInstance(p["clinical_exemptions"], list)
            for ex in p["clinical_exemptions"]:
                self.assertIsInstance(ex, str)

            # ingredients: List[Tuple[str, float, str]] -> (name, pct_w_w, clinical_role)
            self.assertIsInstance(p["ingredients"], list)
            self.assertGreater(len(p["ingredients"]), 0)
            ing_pct_sum = 0.0
            for ing in p["ingredients"]:
                self.assertEqual(len(ing), 3, f"Ingredient row {ing} must be a 3-tuple (name, pct, role)")
                self.assertIsInstance(ing[0], str)
                self.assertIsInstance(ing[1], (int, float))
                self.assertIsInstance(ing[2], str)
                ing_pct_sum += float(ing[1])
            self.assertAlmostEqual(
                ing_pct_sum,
                100.0,
                delta=1.0,
                msg=f"Ingredients sum {ing_pct_sum}% is not ~100%",
            )

            # nutrients: List[Tuple[str, str, Optional[float], Optional[float]]]
            self.assertIsInstance(p["nutrients"], list)
            self.assertGreater(len(p["nutrients"]), 4)

            for nut in p["nutrients"]:
                self.assertEqual(
                    len(nut),
                    4,
                    f"Nutrient row {nut} must be a 4-tuple (name, unit, val_base, val_sec)",
                )
                self.assertIsInstance(nut[0], str)
                self.assertIsInstance(nut[1], str)
                if nut[2] is not None:
                    self.assertIsInstance(nut[2], (int, float))
                if nut[3] is not None:
                    self.assertIsInstance(nut[3], (int, float))

            # Macro balance & dual-column ratio check
            energy_row = next((r for r in p["nutrients"] if "energy" in r[0].lower() or "พลังงาน" in r[0].lower()), None)
            protein_row = next((r for r in p["nutrients"] if "protein" in r[0].lower() or "โปรตีน" in r[0].lower()), None)
            fat_row = next((r for r in p["nutrients"] if "fat" in r[0].lower() or "ไขมัน" in r[0].lower()), None)
            carb_row = next((r for r in p["nutrients"] if "carbohydrate" in r[0].lower() or "คาร์โบไฮเดรต" in r[0].lower()), None)

            self.assertIsNotNone(energy_row, "Energy row missing in nutrients")
            self.assertIsNotNone(protein_row, "Protein row missing in nutrients")
            self.assertIsNotNone(fat_row, "Total fat row missing in nutrients")
            self.assertIsNotNone(carb_row, "Carbohydrate row missing in nutrients")

            # 4*P + 9*F + 4*C ~ base_kcal
            p_val = float(protein_row[2])
            f_val = float(fat_row[2])
            c_val = float(carb_row[2])
            calc_kcal = 4.0 * p_val + 9.0 * f_val + 4.0 * c_val
            self.assertAlmostEqual(
                calc_kcal,
                p["base_kcal"],
                delta=5.0,
                msg=f"Macro calories {calc_kcal} does not match stated base_kcal {p['base_kcal']}",
            )

            # Dual-column ratio check (value_base -> value_secondary matches serving_g / base_value)
            expected_ratio = p["serving_g"] / p["base_value"]
            for row in [energy_row, protein_row, fat_row, carb_row]:
                if row[2] is not None and row[3] is not None:
                    actual_ratio = float(row[3]) / float(row[2])
                    self.assertAlmostEqual(
                        actual_ratio,
                        expected_ratio,
                        places=2,
                        msg=f"Dual-column ratio mismatch for {row[0]}: expected {expected_ratio:.4f}, got {actual_ratio:.4f}",
                    )


if __name__ == "__main__":
    unittest.main()
