# -*- coding: utf-8 -*-
"""
Canonical Single Source of Truth (SSOT) Brand Data for Synthetic Worked Example.

Brand: NutriSample Co., Ltd.
Product: Formula Alpha (Standard Complete Enteral Nutrition Powder)
Fictional Placeholder FDA: 10-X-XXXXXX-X-XXXX
"""

from __future__ import annotations
from typing import Any, Dict, List

# Brand-level Metadata
BRAND_META: Dict[str, Any] = {
    "brand_name": "NutriSample_Co",
    "master_excel": "NutriSample_Co_Nutrition_1000kcal.xlsx",
}

# Canonical Product Specifications
PRODUCTS: List[Dict[str, Any]] = [
    {
        "id": "Formula_Alpha",
        "folder": "Formula_Alpha",
        "trade_th": "นิวทริแซมเปิล ฟอร์มูล่า อัลฟ่า",
        "trade_en": "NutriSample Formula Alpha",
        "generic": "อาหารทางการแพทย์สูตรครบถ้วน ชนิดผง สำหรับผู้ที่ต้องการเสริมสารอาหาร (อาหารทางการแพทย์)",
        "pkg_type": "กระป๋องชนิดผง (Powder Can)",
        "fda": "10-X-XXXXXX-X-XXXX",
        "manufacturer": "บริษัท นิวทริแซมเปิล จำกัด (NutriSample Co., Ltd.)",
        "base_type": "powder",
        "base_value": 100.0,
        "base_kcal": 450.0,
        "serving_g": 55.5,
        "serving_kcal": 250.0,
        "nominal_pcf": (16.0, 30.0, 54.0),
        "clinical_exemptions": [],
        "ingredients": [
            ("มอลโทเดกซ์ทริน (Maltodextrin)", 48.00, "แหล่งคาร์โบไฮเดรตเชิงซ้อน / Complex carbohydrate base"),
            ("โซเดียมเคซีเนตและโปรตีนนม (Sodium caseinate & Milk protein)", 18.00, "โปรตีนคุณภาพสูง / High-quality intact protein"),
            ("น้ำมันคาโนลาและน้ำมันดอกทานตะวัน (Canola & Sunflower oils)", 15.00, "กรดไขมันไม่อิ่มตัว MUFA & PUFA / Healthy fats"),
            ("ซูโครส (Sucrose)", 8.00, "คาร์โบไฮเดรตดูดซึมเร็ว / Fast-absorbing carbohydrate"),
            ("ฟรุคโตโอลิโกแซคคาไรด์ (Fructo-oligosaccharides / FOS)", 5.00, "พรีไบโอติกและใยอาหาร / Prebiotic dietary fiber"),
            ("วิตามินและเกลือแร่พรีมิกซ์ (Vitamins & Minerals premix)", 4.00, "วิตามินและแร่ธาตุจำเป็นตาม Thai RDI / Essential micronutrients"),
            ("สารแต่งกลิ่นและอิมัลซิไฟเออร์ (Flavoring & Emulsifiers)", 2.00, "กลิ่นรสและความคงตัว / Flavor & stability"),
        ],
        "nutrients": [
            ("พลังงาน (Energy)", "kcal", 450.0, 250.0),
            ("โปรตีน (Protein)", "g", 18.00, 9.99),
            ("ไขมันทั้งหมด (Total Fat)", "g", 15.00, 8.33),
            ("ไขมันอิ่มตัว (Saturated Fat)", "g", 2.00, 1.11),
            ("คอเลสเตอรอล (Cholesterol)", "mg", 10.0, 5.55),
            ("คาร์โบไฮเดรตทั้งหมด (Total Carbohydrate)", "g", 60.75, 33.72),
            ("ใยอาหาร (Dietary Fiber)", "g", 5.00, 2.78),
            ("โซเดียม (Sodium)", "mg", 350.0, 194.25),
            ("โพแทสเซียม (Potassium)", "mg", 600.0, 333.00),
            ("คลอไรด์ (Chloride)", "mg", 420.0, 233.10),
            ("แคลเซียม (Calcium)", "mg", 400.0, 222.00),
            ("ฟอสฟอรัส (Phosphorus)", "mg", 300.0, 166.50),
            ("แมกนีเซียม (Magnesium)", "mg", 120.0, 66.60),
            ("เหล็ก (Iron)", "mg", 5.00, 2.78),
            ("ไอโอดีน (Iodine)", "mcg", 60.0, 33.30),
            ("สังกะสี (Zinc)", "mg", 4.50, 2.50),
            ("ซีลีเนียม (Selenium)", "mcg", 24.0, 13.32),
            ("ทองแดง (Copper)", "mcg", 360.0, 199.80),
            ("แมงกานีส (Manganese)", "mg", 1.20, 0.67),
            ("โครเมียม (Chromium)", "mcg", 15.0, 8.33),
            ("โมลิบดินัม (Molybdenum)", "mcg", 20.0, 11.10),
            ("วิตามินเอ (Vitamin A)", "mcg RAE", 320.0, 177.60),
            ("วิตามินดี (Vitamin D)", "mcg", 6.00, 3.33),
            ("วิตามินอี (Vitamin E)", "mg alpha-TE", 4.00, 2.22),
            ("วิตามินเค (Vitamin K)", "mcg", 25.0, 13.88),
            ("วิตามินบี 1 (Vitamin B1)", "mg", 0.50, 0.28),
            ("วิตามินบี 2 (Vitamin B2)", "mg", 0.50, 0.28),
            ("ไนอะซิน (Niacin)", "mg NE", 6.00, 3.33),
            ("กรดแพนโทธีนิค (Pantothenic Acid)", "mg", 2.20, 1.22),
            ("วิตามินบี 6 (Vitamin B6)", "mg", 0.60, 0.33),
            ("ไบโอติน (Biotin)", "mcg", 12.0, 6.66),
            ("โฟเลต (Folate)", "mcg DFE", 160.0, 88.80),
            ("วิตามินบี 12 (Vitamin B12)", "mcg", 1.00, 0.56),
            ("วิตามินซี (Vitamin C)", "mg", 40.0, 22.20),
        ],
    }
]
