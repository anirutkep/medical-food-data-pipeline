# -*- coding: utf-8 -*-
"""
Canonical Single Source of Truth (SSOT) Brand Data Template.

This module defines the canonical Python data contract for medical food / enteral
nutrition products. Downstream pipeline tools (apply_unit_normalization, pipeline_audit,
and ocr_row_extractor --data-py) consume this schema.

Schema Structure:
  - BRAND_META: Dict[str, Any]
      * "brand_name": Name of the brand (str)
      * "master_excel": Target filename for the brand master workbook (str, ending in .xlsx)

  - PRODUCTS: List[Dict[str, Any]]
      Each product dictionary MUST contain:
      * "id": Unique alphanumeric product identifier (str, e.g. "Sample_Formula")
      * "folder": Relative directory path for product deliverables (str, e.g. "SampleBrand/SampleFormula")
      * "trade_th": Thai trade name (str)
      * "trade_en": English trade name (str)
      * "base_type": Product formulation type ("powder" or "rtd_liquid")
      * "base_value": Stated base quantity (float, e.g. 100.0 for 100g powder or 100.0 mL)
      * "base_kcal": Energy in kcal per base quantity (float, e.g. 450.0)
      * "serving_g": Serving size in grams or mL (float, e.g. 50.0)
      * "serving_kcal": Energy in kcal per serving (float, e.g. 225.0)
      * "nominal_pcf": Protein/Fat/Carbohydrate calorie distribution percentage tuple (P, F, C summing to 100)
      * "clinical_exemptions": List[str] of canonical MOPH 445 nutrient keys exempted from standard
                              cohort outlier or clinical plausibility checks (e.g. ["potassium", "vitamin_d"]
                              for specialized renal formulas, or [] for standard complete formulas)
      * "ingredients": List[Tuple[str, float, str]] -> (ingredient_name, percentage_w_w, clinical_role)
                       Total percentage should sum to approximately 100.0%
      * "nutrients": List[Tuple[str, str, Optional[float], Optional[float]]] ->
                     (nutrient_name, unit, value_per_base, value_per_serving)
                     Starting with Energy, Protein, Total Fat, and Carbohydrate matching
                     4*Protein + 9*Fat + 4*Carb ~ base_kcal, with value_per_serving / value_per_base
                     matching serving_g / base_value ratio.
"""

from __future__ import annotations
from typing import Any, Dict, List, Optional, Tuple

# Brand-level Metadata
BRAND_META: Dict[str, Any] = {
    "brand_name": "SampleBrand",
    "master_excel": "SampleBrand_Nutrition_1000kcal.xlsx",
}

# Canonical Product Specifications
PRODUCTS: List[Dict[str, Any]] = [
    {
        "id": "Sample_Standard_Formula",
        "folder": "SampleBrand/SampleStandardFormula",
        "trade_th": "แซมเปิล นิวทริชั่น พาวเดอร์",
        "trade_en": "Sample Nutrition Powder",
        "base_type": "powder",  # "powder" | "rtd_liquid"
        "base_value": 100.0,    # 100g base reference
        "base_kcal": 450.0,     # 450 kcal per 100g
        "serving_g": 50.0,      # 50g per serving
        "serving_kcal": 225.0,  # 225 kcal per serving (ratio = 0.500)
        "nominal_pcf": (16, 30, 54),  # % Calorie ratio: 16% Protein, 30% Fat, 54% Carbohydrate (sum = 100)

        # Clinical exemptions: List of canonical MOPH 445 nutrient keys exempted from
        # standard cohort outlier or RDI deficiency checks (e.g. ["potassium", "phosphorus"] for renal formulas)
        "clinical_exemptions": [],

        # Ingredients: List[Tuple[name, pct_w_w, clinical_role]] (sum ~100%)
        "ingredients": [
            ("มอลโทเดกซ์ทริน (Maltodextrin)", 48.00, "แหล่งคาร์โบไฮเดรตเชิงซ้อน / Complex carbohydrate base"),
            ("โซเดียมเคซีเนตและโปรตีนนม (Sodium caseinate & Milk protein)", 18.00, "โปรตีนคุณภาพสูง / High-quality intact protein"),
            ("น้ำมันคาโนลาและน้ำมันดอกทานตะวัน (Canola & Sunflower oils)", 15.00, "กรดไขมันไม่อิ่มตัว MUFA & PUFA / Healthy fats"),
            ("ซูโครส (Sucrose)", 8.00, "คาร์โบไฮเดรตดูดซึมเร็ว / Fast-absorbing carbohydrate"),
            ("ฟรุคโตโอลิโกแซคคาไรด์ (Fructo-oligosaccharides / FOS)", 5.00, "พรีไบโอติกและใยอาหาร / Prebiotic dietary fiber"),
            ("วิตามินและเกลือแร่พรีมิกซ์ (Vitamins & Minerals premix)", 4.00, "วิตามินและแร่ธาตุ 33 ชนิดตาม Thai RDI / Essential micronutrients"),
            ("สารแต่งกลิ่นและอิมัลซิไฟเออร์ (Flavoring & Emulsifiers)", 2.00, "กลิ่นรสและความคงตัว / Flavor & stability"),
        ],

        # Nutrients: List[Tuple[name, unit, value_base, value_serving]]
        # Dual-column ratio = serving_g / base_value = 50.0 / 100.0 = 0.500
        # Macro check: 4 * 18.0 (P) + 9 * 15.0 (F) + 4 * 60.75 (C) = 72 + 135 + 243 = 450.0 kcal
        "nutrients": [
            ("พลังงาน (Energy)", "kcal", 450.0, 225.0),
            ("โปรตีน (Protein)", "g", 18.00, 9.00),
            ("ไขมันทั้งหมด (Total Fat)", "g", 15.00, 7.50),
            ("ไขมันอิ่มตัว (Saturated Fat)", "g", 2.00, 1.00),
            ("คอเลสเตอรอล (Cholesterol)", "mg", 10.0, 5.0),
            ("คาร์โบไฮเดรตทั้งหมด (Total Carbohydrate)", "g", 60.75, 30.38),
            ("ใยอาหาร (Dietary Fiber)", "g", 5.00, 2.50),
            ("โซเดียม (Sodium)", "mg", 350.0, 175.0),
            ("โพแทสเซียม (Potassium)", "mg", 600.0, 300.0),
            ("คลอไรด์ (Chloride)", "mg", 420.0, 210.0),
            ("แคลเซียม (Calcium)", "mg", 400.0, 200.0),
            ("ฟอสฟอรัส (Phosphorus)", "mg", 300.0, 150.0),
            ("แมกนีเซียม (Magnesium)", "mg", 120.0, 60.0),
            ("เหล็ก (Iron)", "mg", 5.0, 2.5),
            ("ไอโอดีน (Iodine)", "mcg", 60.0, 30.0),
            ("สังกะสี (Zinc)", "mg", 4.5, 2.25),
            ("ซีลีเนียม (Selenium)", "mcg", 24.0, 12.0),
            ("ทองแดง (Copper)", "mcg", 360.0, 180.0),
            ("แมงกานีส (Manganese)", "mg", 1.2, 0.6),
            ("โครเมียม (Chromium)", "mcg", 15.0, 7.5),
            ("โมลิบดินัม (Molybdenum)", "mcg", 20.0, 10.0),
            ("วิตามินเอ (Vitamin A)", "mcg RAE", 320.0, 160.0),
            ("วิตามินดี (Vitamin D)", "mcg", 6.0, 3.0),
            ("วิตามินอี (Vitamin E)", "mg alpha-TE", 4.0, 2.0),
            ("วิตามินเค (Vitamin K)", "mcg", 25.0, 12.5),
            ("วิตามินบี 1 (Vitamin B1)", "mg", 0.50, 0.25),
            ("วิตามินบี 2 (Vitamin B2)", "mg", 0.50, 0.25),
            ("ไนอะซิน (Niacin)", "mg NE", 6.0, 3.0),
            ("กรดแพนโทธีนิค (Pantothenic Acid)", "mg", 2.2, 1.1),
            ("วิตามินบี 6 (Vitamin B6)", "mg", 0.6, 0.3),
            ("ไบโอติน (Biotin)", "mcg", 12.0, 6.0),
            ("โฟเลต (Folate)", "mcg DFE", 160.0, 80.0),
            ("วิตามินบี 12 (Vitamin B12)", "mcg", 1.0, 0.5),
            ("วิตามินซี (Vitamin C)", "mg", 40.0, 20.0),
        ],
    }
]
