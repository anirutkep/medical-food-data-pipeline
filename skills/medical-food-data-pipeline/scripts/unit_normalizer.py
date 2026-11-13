# -*- coding: utf-8 -*-
"""
Unit Normalization & MOPH Notification No. 445 (B.E. 2566) Standardizer.

References:
  - Thai RDI (33 Nutrients): Ministry of Public Health Notification No. 445 (B.E. 2566),
    Appendix 3 (Royal Gazette Vol. 141, Special Part 3 Ngor, Page 26, Jan 4, 2024).
  - Unit Conversion Factors: Appendix 1, Section 2.5 (Vitamins & Minerals).
"""

from __future__ import annotations
import re
from typing import Optional, Tuple, Dict

# 33 Official Thai RDI Nutrients (MOPH Notification No. 445, Appendix 3)
# ASCII-safe standard units: g, mg, mcg, mcg RAE, mg alpha-TE, mg NE, mcg DFE
THAI_RDI_445_STANDARDS: Dict[str, Dict[str, object]] = {
    "total_fat": {"no": 1, "th": "ไขมันทั้งหมด", "en": "Total Fat", "rdi": 65.0, "unit": "g"},
    "saturated_fat": {"no": 2, "th": "ไขมันอิ่มตัว", "en": "Saturated Fat", "rdi": 20.0, "unit": "g"},
    "cholesterol": {"no": 3, "th": "คอเลสเตอรอล", "en": "Cholesterol", "rdi": 300.0, "unit": "mg"},
    "protein": {"no": 4, "th": "โปรตีน", "en": "Protein", "rdi": 50.0, "unit": "g"},
    "total_carbohydrate": {"no": 5, "th": "คาร์โบไฮเดรตทั้งหมด", "en": "Total Carbohydrate", "rdi": 300.0, "unit": "g"},
    "dietary_fiber": {"no": 6, "th": "ใยอาหาร", "en": "Dietary Fiber", "rdi": 25.0, "unit": "g"},
    "vitamin_a": {"no": 7, "th": "วิตามินเอ", "en": "Vitamin A", "rdi": 800.0, "unit": "mcg RAE"},
    "vitamin_d": {"no": 8, "th": "วิตามินดี", "en": "Vitamin D", "rdi": 15.0, "unit": "mcg"},
    "vitamin_e": {"no": 9, "th": "วิตามินอี", "en": "Vitamin E", "rdi": 9.0, "unit": "mg alpha-TE"},
    "vitamin_k": {"no": 10, "th": "วิตามินเค", "en": "Vitamin K", "rdi": 60.0, "unit": "mcg"},
    "vitamin_b1": {"no": 11, "th": "วิตามินบี 1/ไธอะมิน", "en": "Vitamin B1/Thiamin", "rdi": 1.2, "unit": "mg"},
    "vitamin_b2": {"no": 12, "th": "วิตามินบี 2/ไรโบฟลาวิน", "en": "Vitamin B2/Riboflavin", "rdi": 1.2, "unit": "mg"},
    "niacin": {"no": 13, "th": "ไนอะซิน", "en": "Niacin", "rdi": 15.0, "unit": "mg NE"},
    "pantothenic_acid": {"no": 14, "th": "กรดแพนโทธีนิค", "en": "Pantothenic Acid", "rdi": 5.0, "unit": "mg"},
    "vitamin_b6": {"no": 15, "th": "วิตามินบี 6", "en": "Vitamin B6", "rdi": 1.3, "unit": "mg"},
    "biotin": {"no": 16, "th": "ไบโอติน", "en": "Biotin", "rdi": 30.0, "unit": "mcg"},
    "folate": {"no": 17, "th": "โฟเลต", "en": "Folate", "rdi": 400.0, "unit": "mcg DFE"},
    "vitamin_b12": {"no": 18, "th": "วิตามินบี 12", "en": "Vitamin B12", "rdi": 2.4, "unit": "mcg"},
    "vitamin_c": {"no": 19, "th": "วิตามินซี", "en": "Vitamin C", "rdi": 100.0, "unit": "mg"},
    "calcium": {"no": 20, "th": "แคลเซียม", "en": "Calcium", "rdi": 1000.0, "unit": "mg"},
    "phosphorus": {"no": 21, "th": "ฟอสฟอรัส", "en": "Phosphorus", "rdi": 700.0, "unit": "mg"},
    "magnesium": {"no": 22, "th": "แมกนีเซียม", "en": "Magnesium", "rdi": 310.0, "unit": "mg"},
    "iron": {"no": 23, "th": "เหล็ก", "en": "Iron", "rdi": 22.0, "unit": "mg"},
    "iodine": {"no": 24, "th": "ไอโอดีน", "en": "Iodine", "rdi": 150.0, "unit": "mcg"},
    "zinc": {"no": 25, "th": "สังกะสี", "en": "Zinc", "rdi": 11.0, "unit": "mg"},
    "selenium": {"no": 26, "th": "ซีลีเนียม", "en": "Selenium", "rdi": 60.0, "unit": "mcg"},
    "copper": {"no": 27, "th": "ทองแดง", "en": "Copper", "rdi": 900.0, "unit": "mcg"},
    "manganese": {"no": 28, "th": "แมงกานีส", "en": "Manganese", "rdi": 3.0, "unit": "mg"},
    "molybdenum": {"no": 29, "th": "โมลิบดินัม", "en": "Molybdenum", "rdi": 45.0, "unit": "mcg"},
    "chromium": {"no": 30, "th": "โครเมียม", "en": "Chromium", "rdi": 35.0, "unit": "mcg"},
    "sodium": {"no": 31, "th": "โซเดียม", "en": "Sodium", "rdi": 2000.0, "unit": "mg"},
    "potassium": {"no": 32, "th": "โพแทสเซียม", "en": "Potassium", "rdi": 3500.0, "unit": "mg"},
    "chloride": {"no": 33, "th": "คลอไรด์", "en": "Chloride", "rdi": 2300.0, "unit": "mg"},
}


def canonicalize_nutrient_name(raw_name: str) -> Optional[str]:
    """Map Thai/English label nutrient strings to one of the 33 canonical keys (or None)."""
    s = raw_name.strip().lower()
    # Exclusions for sub-components or non-33 items that share substrings
    if any(ex in s for ex in ("cahmb", "เอช เอ็ม บี", "hmb", "energy from fat", "พลังงานจากไขมัน", "เบตา-แคโรทีน", "beta-carotene", "mct", "mufa", "pufa", "linoleic", "linolenic", "omega", "โอเมก้า", "trans fat", "ไขมันทรานส์")):
        return None

    # Order matters (e.g. B12 before B1, saturated_fat before total_fat)
    if any(k in s for k in ("ไขมันอิ่มตัว", "saturated fat")):
        return "saturated_fat"
    if any(k in s for k in ("ไขมันทั้งหมด", "total fat")) or re.search(r"(^|\s|\.)ไขมัน(\s|\(|$)", s):
        return "total_fat"
    if any(k in s for k in ("คอเลสเตอรอล", "โคเลสเตอรอล", "cholesterol")):
        return "cholesterol"
    if any(k in s for k in ("โปรตีน", "protein")):
        return "protein"
    if any(k in s for k in ("คาร์โบไฮเดรต", "carbohydrate")):
        return "total_carbohydrate"
    if any(k in s for k in ("ใยอาหาร", "dietary fiber", "fiber")) and not any(x in s for x in ("fos", "inulin", "phgg", "polydextrose", "soluble", "insoluble", "ละลายน้ำ", "ไม่ละลายน้ำ")):
        return "dietary_fiber"

    # Vitamins
    if any(k in s for k in ("วิตามินเอ", "วิตามิน เอ", "vitamin a")):
        return "vitamin_a"
    if any(k in s for k in ("วิตามินดี", "วิตามิน ดี", "vitamin d")):
        return "vitamin_d"
    if any(k in s for k in ("วิตามินอี", "วิตามิน อี", "vitamin e")):
        return "vitamin_e"
    if any(k in s for k in ("วิตามินเค", "วิตามิน เค", "vitamin k")):
        return "vitamin_k"
    if any(k in s for k in ("บี 12", "บี12", "b12", "b 12")):
        return "vitamin_b12"
    if any(k in s for k in ("บี 1", "บี1", "b1", "b 1", "ไทอามีน", "ไธอะมิน", "thiamin")):
        return "vitamin_b1"
    if any(k in s for k in ("บี 2", "บี2", "b2", "b 2", "ไรโบฟลาวิน", "riboflavin")):
        return "vitamin_b2"
    if any(k in s for k in ("ไนอะซ", "ไนอาซ", "niacin", "nicotinic", "บี 3", "บี3", "b3")):
        return "niacin"
    if any(k in s for k in ("แพนโท", "pantothen")):
        return "pantothenic_acid"
    if any(k in s for k in ("บี 6", "บี6", "b6", "b 6")):
        return "vitamin_b6"
    if any(k in s for k in ("ไบโอติน", "biotin")):
        return "biotin"
    if any(k in s for k in ("โฟเลต", "โฟลิก", "โฟลิค", "folate", "folic")):
        return "folate"
    if any(k in s for k in ("วิตามินซี", "วิตามิน ซี", "vitamin c")):
        return "vitamin_c"

    # Minerals
    if any(k in s for k in ("แคลเซียม", "calcium")):
        return "calcium"
    if any(k in s for k in ("ฟอสฟอรัส", "phosphorus")):
        return "phosphorus"
    if any(k in s for k in ("แมกนีเซียม", "magnesium")):
        return "magnesium"
    if any(k in s for k in ("เหล็ก", "iron")):
        return "iron"
    if any(k in s for k in ("ไอโอดีน", "iodine")):
        return "iodine"
    if any(k in s for k in ("สังกะสี", "zinc")):
        return "zinc"
    if any(k in s for k in ("ซีลีเนียม", "selenium")):
        return "selenium"
    if any(k in s for k in ("ทองแดง", "copper")):
        return "copper"
    if any(k in s for k in ("แมงกานีส", "manganese")):
        return "manganese"
    if any(k in s for k in ("โมลิบดินัม", "โมลิบดีนัม", "molybdenum")):
        return "molybdenum"
    if any(k in s for k in ("โครเมียม", "chromium")):
        return "chromium"
    if any(k in s for k in ("โซเดียม", "sodium")):
        return "sodium"
    if any(k in s for k in ("โพแทสเซียม", "potassium")):
        return "potassium"
    if any(k in s for k in ("คลอไรด์", "chloride")):
        return "chloride"

    return None


def normalize_unit_token(raw_unit: str) -> str:
    """Normalize verbose Thai/English unit strings into canonical unit tokens."""
    u = raw_unit.strip()
    ul = u.lower()
    if "rae" in ul:
        return "mcg RAE"
    if "re" in ul and ("mcg" in ul or "µg" in ul or "ไมโครกรัม" in ul or "มคก" in ul):
        return "mcg RE"
    if "dfe" in ul:
        return "mcg DFE"
    if "ne" in ul and ("mg" in ul or "มก" in ul or "มิลลิกรัม" in ul):
        return "mg NE"
    if ("alpha-te" in ul or "α-te" in ul or "อัลฟา-โทโคฟีรอล" in ul or "te" in ul) and ("mg" in ul or "มก" in ul or "มิลลิกรัม" in ul):
        return "mg alpha-TE"
    if "iu" in ul or "หน่วยสากล" in ul:
        return "IU"
    if "kcal" in ul or "กิโลแคลอรี" in ul:
        return "kcal"
    if "kj" in ul or "กิโลจูล" in ul:
        return "kJ"
    if any(k in ul for k in ("mcg", "µg", "ไมโครกรัม", "มคก")):
        return "mcg"
    if any(k in ul for k in ("mg", "มิลลิกรัม", "มก")):
        return "mg"
    if ul in ("g", "กรัม") or "กรัม (g)" in ul:
        return "g"
    return u


def _si_mass_factor(src_unit: str, dst_unit: str) -> Optional[float]:
    """Return multiplier between g, mg, and mcg."""
    scale = {"g": 1.0, "mg": 1e-3, "mcg": 1e-6}
    if src_unit in scale and dst_unit in scale:
        return scale[src_unit] / scale[dst_unit]
    return None


def get_conversion_factor(
    canon_key: Optional[str] = None,
    raw_unit: str = "",
    ingredients_text: str = "",
    canonical_nutrient: Optional[str] = None,
) -> Tuple[float, str, str]:
    """
    Return (multiplier, target_unit, assumption_note) according to MOPH 445 Section 2.5 & Appendix 3.
    Accepts either `canon_key` or `canonical_nutrient` for compatibility.
    If nutrient is not in the 33-item Thai RDI table, falls back to SI/ASCII unit cleanup with factor 1.0.
    """
    key = canon_key if canon_key is not None else canonical_nutrient
    src = normalize_unit_token(raw_unit)
    if not key or key not in THAI_RDI_445_STANDARDS:
        return (1.0, src, "")

    target = str(THAI_RDI_445_STANDARDS[key]["unit"])
    if src == target:
        return (1.0, target, "")

    # Direct SI mass scaling (g <-> mg <-> mcg)
    si = _si_mass_factor(src, target)
    if si is not None:
        return (si, target, "")

    ing_lower = ingredients_text.lower()

    # Section 2.5 specific vitamin rules
    if key == "vitamin_a":
        if src == "IU":
            return (0.3, target, "ข้อ 2.5: 1 IU retinol = 0.3 mcg RAE (สมมติฐาน preformed vitamin A)")
        if src == "mcg RE":
            return (1.0, target, "ข้อ 2.5: 1 mcg RE ≈ 1 mcg RAE (preformed vitamin A)")
        if src == "mcg":
            return (1.0, target, "สมมติฐาน preformed retinol (1 mcg = 1 mcg RAE)")
        if src == "mg":
            return (1000.0, target, "สมมติฐาน preformed retinol (1 mg = 1,000 mcg RAE)")

    if key == "vitamin_d":
        if src == "IU":
            return (1.0 / 40.0, target, "ข้อ 2.5: 1 mcg cholecalciferol = 40 IU vitamin D")

    if key == "vitamin_e":
        if src == "IU":
            # Check ingredients for natural d-alpha vs synthetic dl-alpha / all-rac
            if re.search(r"\b(d-alpha|d-α|rrr-alpha|rrr-α)\b", ing_lower) and "dl-alpha" not in ing_lower:
                return (1.0 / 1.5, target, "ข้อ 2.5: 1 mg alpha-TE = 1.5 IU (d-alpha-tocopherol)")
            return (1.0 / 2.22, target, "ข้อ 2.5: 1 mg alpha-TE = 2.22 IU (all-rac/dl-alpha-tocopherol ค่าเริ่มต้น)")
        if src == "mg":
            return (1.0, target, "สมมติฐาน 1 mg alpha-tocopherol = 1 mg alpha-TE")

    if key == "niacin":
        if src in ("mg", "mg NE"):
            return (1.0, target, "ข้อ 2.5: 1 mg NE = 1 mg niacin (preformed)")
        if src == "mcg":
            return (0.001, target, "ข้อ 2.5: 1,000 mcg = 1 mg niacin = 1 mg NE")

    if key == "folate":
        if src == "mcg":
            return (1.0 / 0.6, target, "ข้อ 2.5: 1 mcg DFE = 0.6 mcg folic acid added (x 1.6667)")
        if src == "mg":
            return (1000.0 / 0.6, target, "ข้อ 2.5: 1 mg folic acid added = 1,666.67 mcg DFE")

    return (1.0, src, f"ไม่สามารถแปลงหน่วย {src} -> {target} อัตโนมัติ")


def normalize_nutrient(
    raw_name: str,
    raw_unit: str,
    value_1000kcal: Optional[float],
    ingredients_text: str = "",
) -> Dict[str, object]:
    """Compute normalized 1,000 kcal value, target unit, multiplier, and assumption note."""
    canon = canonicalize_nutrient_name(raw_name)
    factor, target_unit, note = get_conversion_factor(canon, raw_unit, ingredients_text)
    norm_val = round(float(value_1000kcal) * factor, 2) if isinstance(value_1000kcal, (int, float)) else None
    return {
        "canonical_key": canon,
        "raw_unit": raw_unit,
        "target_unit": target_unit,
        "factor": factor,
        "normalized_1000kcal": norm_val,
        "note": note,
    }
