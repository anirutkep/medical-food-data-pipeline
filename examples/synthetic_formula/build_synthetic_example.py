# -*- coding: utf-8 -*-
"""
Builder for Synthetic Formula Worked Example Deliverables.

Generates:
  1. Set 1 Markdown (01_Formula_Alpha_Label.md)
  2. Set 2 Markdown (02_Formula_Alpha_Ingredients_Nutrition_1000kcal.md)
  3. Standalone Excel (Formula_Alpha_Nutrition_1000kcal.xlsx) with 100% dynamic =ROUND(...) formulas
  4. Brand Master Excel (NutriSample_Co_Nutrition_1000kcal.xlsx) with 100% dynamic =ROUND(...) formulas
  5. Runs apply_unit_normalization.normalize_workspace() to add MOPH 445 normalized columns & sheets
"""

from __future__ import annotations
import argparse
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# Ensure sibling scripts & data module are importable
THIS_DIR = Path(__file__).resolve().parent
SKILL_ROOT = THIS_DIR.parent.parent
SCRIPTS_DIR = SKILL_ROOT / "skills" / "medical-food-data-pipeline" / "scripts"

for p in (str(THIS_DIR), str(SCRIPTS_DIR)):
    if p not in sys.path:
        sys.path.insert(0, p)

import synthetic_brand_data
from apply_unit_normalization import normalize_workspace


def _get_styles(font_name: str = "TH Sarabun New") -> Dict[str, Any]:
    """Generate Corporate Blue styling palette keyed to the specified font family."""
    thin_side = Side(border_style="thin", color="D3D3D3")
    double_side = Side(border_style="double", color="1B365D")
    return {
        "title": Font(name=font_name, size=16, bold=True, color="1B365D"),
        "subtitle": Font(name=font_name, size=13, italic=True, color="555555"),
        "section": Font(name=font_name, size=14, bold=True, color="1B365D"),
        "hdr": Font(name=font_name, size=13, bold=True, color="FFFFFF"),
        "category": Font(name=font_name, size=12, bold=True, color="2F5597"),
        "data": Font(name=font_name, size=12, color="000000"),
        "bold": Font(name=font_name, size=12, bold=True, color="000000"),
        "total": Font(name=font_name, size=13, bold=True, color="000000"),
        "fill_hdr": PatternFill("solid", fgColor="1B365D"),
        "fill_sub": PatternFill("solid", fgColor="D9E1F2"),
        "fill_cat": PatternFill("solid", fgColor="F2F2F2"),
        "fill_total": PatternFill("solid", fgColor="E2EFDA"),
        "border_data": Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side),
        "border_hdr": Border(left=thin_side, right=thin_side, top=thin_side, bottom=Side(border_style="medium", color="1B365D")),
        "border_total": Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side),
    }


def _autofit_columns(ws: Any, min_col: int = 1, max_col: Optional[int] = None, padding: float = 5.0) -> None:
    """Calculate UTF-8 glyph-aware visual width with 1.6x factor for non-ASCII Thai text + 5 padding."""
    if max_col is None:
        max_col = ws.max_column
    for col in range(min_col, max_col + 1):
        max_len = 0.0
        col_letter = get_column_letter(col)
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is not None and not str(val).startswith("="):
                for line in str(val).split("\n"):
                    w = 0.0
                    for ch in line:
                        w += 1.6 if ord(ch) > 127 else 1.0
                    if w > max_len:
                        max_len = w
        ws.column_dimensions[col_letter].width = max(max_len + padding, 12.0)


def generate_set1_markdown(product: Dict[str, Any], output_path: Path) -> None:
    """Generate Set 1 Markdown (01_*_Label.md) full 360 label."""
    p = product
    ing_rows: List[str] = []
    tot_pct = 0.0
    for idx, (name, pct, role) in enumerate(p["ingredients"], 1):
        tot_pct += pct
        ing_rows.append(f"| {idx} | **{name}** | **{pct:.2f}%** | {pct:.2f} | {role} |")
    ing_rows.append(f"| **รวม** | **ส่วนประกอบทั้งหมดของสูตร** | **{tot_pct:.2f}%** | **{tot_pct:.2f}** | **สูตรสมบูรณ์ตรงตามฉลากบรรจุภัณฑ์ 100%** |")
    ing_table = "\n".join(ing_rows)

    nut_rows: List[str] = []
    for name, unit, v_base, v_serv in p["nutrients"]:
        base_str = f"{v_base:.2f}" if v_base is not None else "-"
        serv_str = f"{v_serv:.2f}" if v_serv is not None else "-"
        nut_rows.append(f"| {name} | {unit} | {base_str} | {serv_str} |")
    nut_table = "\n".join(nut_rows)

    content = f"""# ข้อมูลฉลากผลิตภัณฑ์ฉบับสมบูรณ์ 360 องศา (Full Product Label 360°)
## {p['trade_th']} ({p['trade_en']})
### สกัดและสร้างข้อมูลจำลองมาตรฐาน 100% ({p['pkg_type']} ขนาดบรรจุ {p['base_value']:.0f} กรัม)
*{p['manufacturer']}*

---

### 1. ข้อมูลทั่วไปและเอกลักษณ์ผลิตภัณฑ์ (General Information & Regulatory Identity)

* **ชื่อทางการค้า (ภาษาไทย):** {p['trade_th']}
* **ชื่อทางการค้า (ภาษาอังกฤษ):** {p['trade_en']}
* **ชื่อสามัญ/ประเภทอาหาร:** {p['generic']}
* **กลิ่น/รสชาติ:** รสธรรมชาติ (Natural Flavour)
* **ชนิดและขนาดบรรจุภัณฑ์:** {p['pkg_type']} น้ำหนักสุทธิ **400 กรัม** (ให้พลังงาน **{p['base_kcal']:.0f} kcal / 100g**, **{p['serving_kcal']:.0f} kcal ต่อหน่วยบริโภค {p['serving_g']:.1f}g**)
* **เลขสารบบอาหาร (อย.):** `{p['fda']}`
* **การรับรองเครื่องหมายฮาลาล:** ได้รับการรับรองตามมาตรฐานผลิตภัณฑ์อาหารฮาลาล
* **ผู้ผลิต (Manufacturer):** {p['manufacturer']}

---

### 2. คุณสมบัติ ข้อบ่งใช้ และข้อห้ามใช้ทางคลินิก (Key Clinical Features, Indications & Contraindications)

1. **อาหารสูตรครบถ้วนมาตรฐาน (Standard Complete Formula):**
   * ให้พลังงาน **{p['base_kcal']:.0f} กิโลแคลอรีต่อ 100 กรัม** (ความเข้มข้นมาตรฐาน 1.0 kcal/mL เมื่อผสมตามอัตราส่วน)
   * ใช้โปรตีนคุณภาพสูงจากนมและเคซีน พร้อมใยอาหารพรีไบโอติก FOS
2. **การกระจายพลังงานสมดุล:** โปรตีน **{p['nominal_pcf'][0]:.0f}%** : ไขมัน **{p['nominal_pcf'][1]:.0f}%** : คาร์โบไฮเดรต **{p['nominal_pcf'][2]:.0f}%**
3. **คำเตือนและข้อควรระวังเด็ดขาด (Strict Warnings):**
   * **ห้ามให้ทางหลอดเลือดดำ (Not for intravenous use):** สำหรับรับประทานหรือให้อาหารทางสายยางเท่านั้น
   * ใช้ตามคำแนะนำของแพทย์หรือนักกำหนดอาหาร

---

### 3. ส่วนประกอบที่สำคัญเชิงปริมาณ (Quantitative Ingredients & Additives)

| ลำดับ | ส่วนประกอบที่สำคัญ | สัดส่วน (% w/w) | ปริมาณต่อ 100g (%) | บทบาทหน้าที่ทางคลินิกและสูตรตำรับ |
|:---:|:---|:---:|:---:|:---|
{ing_table}

* **ข้อมูลสำหรับผู้แพ้อาหาร (Allergen Information):** มีผลิตภัณฑ์จากนมและผลิตภัณฑ์จากถั่วเหลือง (Contains milk and soy products)

---

### 4. วิธีการเตรียมและการให้ทางสายให้อาหาร (Preparation & Administration Protocol)

#### 4.1 การเตรียมสำหรับดื่มเสริมมื้ออาหาร (Oral Feeding)
* ผสมผง **{p['serving_g']:.1f} กรัม** กับน้ำสะอาด 200 มล. คนให้ละลาย จะได้สารอาหารปริมาตรประมาณ 230 มล. ให้พลังงาน **{p['serving_kcal']:.0f} กิโลแคลอรี**

#### 4.2 โปรโตคอลการบริหารทางสายให้อาหาร (Tube Administration Protocol)
* ระวังความสะอาดขณะเตรียมและให้อาหารทางสายยาง ปรับอัตราการไหลและปริมาตรตามคำแนะนำของแพทย์
* **การล้างสายยาง (Water Flush Invariant):** ต้อง Flush สายให้อาหารด้วยน้ำต้มสุกสะอาด **20-30 มล.** ก่อนและหลังให้อาหารทุกครั้ง หรือทุก 4 ชั่วโมงกรณีให้แบบหยดต่อเนื่อง เพื่อป้องกันสายยางอุดตันและรักษาความสะอาด

#### 4.3 การเก็บรักษา (Storage Guidelines)
* เมื่อเปิดกระป๋องแล้ว ปิดฝาให้สนิท เก็บในที่แห้งและเย็น หลีกเลี่ยงแสงแดด และควรใช้ให้หมดภายใน 1 เดือน

---

### 5. ตารางข้อมูลโภชนาการฉบับเต็มตามฉลากจริง (Full Nutrition Facts)

| รายการสารอาหาร | หน่วยตามฉลาก | ต่อ 100 กรัม | ต่อ 1 หน่วยบริโภค ({p['serving_g']:.1f} กรัม) |
|:---|:---:|:---:|:---:|
{nut_table}
"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")


def generate_set2_markdown(product: Dict[str, Any], output_path: Path) -> None:
    """Generate Set 2 Markdown (02_*_1000kcal.md) clinical reference per 1,000 kcal."""
    p = product
    scale_100 = 1000.0 / float(p["base_kcal"])
    scale_serv = 1000.0 / float(p["serving_kcal"])
    g_1000 = 100.0 * scale_100

    # Ingredients per 1000 kcal
    ing_rows: List[str] = []
    tot_pct = 0.0
    tot_g = 0.0
    for idx, (name, pct, role) in enumerate(p["ingredients"], 1):
        g_val = round(g_1000 * pct / 100.0, 2)
        tot_pct += pct
        tot_g += g_val
        ing_rows.append(f"| {idx} | **{name}** | {pct:.2f}% | **{g_val:.2f} g** | {role} |")
    ing_rows.append(f"| **รวม** | **ปริมาณผงมาตรฐานต่อ 1,000 kcal** | **{tot_pct:.2f}%** | **{tot_g:.2f} g** | *มาตรฐานผงเทียบเคียงพลังงาน 1,000 kcal* |")
    ing_table = "\n".join(ing_rows)

    # Nutrients per 1000 kcal
    p_1000 = c_1000 = f_1000 = 0.0
    nut_rows: List[str] = []
    for idx, (name, unit, v_base, v_serv) in enumerate(p["nutrients"], 1):
        v1000 = round(float(v_base) * scale_100, 2) if v_base is not None else 0.0
        if "โปรตีน" in name:
            p_1000 = v1000
        elif "คาร์โบไฮเดรต" in name:
            c_1000 = v1000
        elif "ไขมันทั้งหมด" in name:
            f_1000 = v1000
        base_str = f"{v_base:.2f}" if v_base is not None else "-"
        serv_str = f"{v_serv:.2f}" if v_serv is not None else "-"
        nut_rows.append(f"| {idx} | {name} | {unit} | {base_str} | {serv_str} | **{v1000:,.2f}** |")
    nut_table = "\n".join(nut_rows)

    p_kcal = p_1000 * 4.0
    c_kcal = c_1000 * 4.0
    f_kcal = f_1000 * 9.0
    tot_kcal = p_kcal + c_kcal + f_kcal
    p_pct = (p_kcal / tot_kcal) * 100.0 if tot_kcal > 0 else 0.0
    c_pct = (c_kcal / tot_kcal) * 100.0 if tot_kcal > 0 else 0.0
    f_pct = (f_kcal / tot_kcal) * 100.0 if tot_kcal > 0 else 0.0

    content = f"""# ข้อมูลส่วนประกอบและสารอาหารฉบับเฉพาะต่อ 1,000 กิโลแคลอรี
## {p['trade_th']} ({p['trade_en']})
### มาตรฐานการเทียบเคียงเชิงคลินิกสำหรับบุคลากรทางการแพทย์ (Clinical Nutrition Reference per 1,000 kcal)

---

### 1. กรอบมาตรฐานการคำนวณ (Calculation Standard Box)

> [!IMPORTANT]
> **ตัวคูณและค่าฐานการคำนวณ (Conversion Standards — {p['trade_en']}):**
> * **พลังงานต่อ 100 กรัม:** **{p['base_kcal']:.2f} kcal**
> * **พลังงานต่อ 1 หน่วยบริโภค ({p['serving_g']:.1f} กรัม):** **{p['serving_kcal']:.2f} kcal**
> * **ปริมาณผงเพื่อให้ได้พลังงาน 1,000 kcal:** $\\frac{{1,000}}{{{p['base_kcal']:.0f}}} \\times 100 = \\mathbf{{{g_1000:.2f}\\text{{ g}}}}$
> * **ตัวคูณจากต่อ 100 กรัม เป็นต่อ 1,000 kcal:** $\\times \\mathbf{{{scale_100:.6f}}}$
> * **ตัวคูณจากต่อ 1 หน่วยบริโภค ({p['serving_g']:.1f} กรัม) เป็นต่อ 1,000 kcal:** $\\times \\mathbf{{{scale_serv:.6f}}}$

---

### 2. ตารางส่วนประกอบสำคัญเชิงปริมาณต่อ 1,000 กิโลแคลอรี (Ingredients per 1,000 kcal)

| ลำดับ | ส่วนประกอบที่สำคัญ | สัดส่วนสูตร (% w/w) | ปริมาณเทียบเคียงต่อ 1,000 kcal ({g_1000:.2f} g) | บทบาทหน้าที่ทางคลินิก |
|:---:|:---|:---:|:---:|:---|
{ing_table}

---

### 3. ระบบสารอาหารเฉพาะทางคลินิก (Special Nutritional Systems)

* **Balanced Macronutrient System ({p['nominal_pcf'][0]:.0f}% P : {p['nominal_pcf'][1]:.0f}% F : {p['nominal_pcf'][2]:.0f}% C):** พลังงานสมดุลตามเกณฑ์มาตรฐานทางการแพทย์
* **High-Quality Protein ({p['nominal_pcf'][0]:.0f}% Energy = {p_1000:.2f} g / 1,000 kcal):** โปรตีนเคซีนและนมคุณภาพสูง เสริมสร้างมวลกล้ามเนื้อ
* **Healthy Fat System ({p['nominal_pcf'][1]:.0f}% Energy = {f_1000:.2f} g / 1,000 kcal):** กรดไขมันไม่อิ่มตัวเชิงเดี่ยวและหลายตำแหน่ง (MUFA/PUFA)
* **Prebiotic FOS Fiber System (11.11 g / 1,000 kcal):** ใยอาหารกระตุ้นการทำงานของจุลินทรีย์ชนิดดีในลำไส้

---

### 4. ตารางสารอาหารฉบับสมบูรณ์ต่อ 1,000 กิโลแคลอรี (Full Nutrition Facts per 1,000 kcal)

| ลำดับ | รายการสารอาหาร | หน่วย | ต่อ 100 กรัม | ต่อ {p['serving_g']:.1f} กรัม | **ต่อ 1,000 kcal (คำนวณ)** |
|:---:|:---|:---:|:---:|:---:|:---:|
{nut_table}

---

### 5. ตารางตรวจสอบสัดส่วนการกระจายพลังงาน (Macronutrient Distribution Verification)

| สารอาหารหลัก | ปริมาณต่อ 1,000 kcal (g) | ตัวคูณพลังงาน (kcal/g) | พลังงานที่ได้ (kcal) | สัดส่วนพลังงานจริง (% Energy) | สัดส่วนอ้างอิงตามฉลาก (Nominal %) |
|:---|:---:|:---:|:---:|:---:|:---:|
| **โปรตีน (Protein)** | {p_1000:.2f} g | 4 kcal/g | {p_kcal:.2f} kcal | **{p_pct:.2f}%** | **{p['nominal_pcf'][0]:.0f}%** |
| **คาร์โบไฮเดรต (Carbohydrate)** | {c_1000:.2f} g | 4 kcal/g | {c_kcal:.2f} kcal | **{c_pct:.2f}%** | **{p['nominal_pcf'][2]:.0f}%** |
| **ไขมัน (Fat)** | {f_1000:.2f} g | 9 kcal/g | {f_kcal:.2f} kcal | **{f_pct:.2f}%** | **{p['nominal_pcf'][1]:.0f}%** |
| **รวมทั้งหมด** | - | - | **{tot_kcal:,.2f} kcal** | **100.00%** | **{p['nominal_pcf'][0]:.0f} : {p['nominal_pcf'][1]:.0f} : {p['nominal_pcf'][2]:.0f}** |
"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")


def _build_product_sheet(ws: Any, product: Dict[str, Any], styles: Dict[str, Any]) -> None:
    """Populate single product worksheet with Corporate Blue styling and 100% dynamic =ROUND(...) formulas."""
    p = product
    sname = p["id"]
    ws.title = sname
    ws.views.sheetView[0].showGridLines = True

    # Title & Subtitle
    ws.cell(1, 1, f"{p['trade_th']} ({p['trade_en']}) - ข้อมูลฉลากโภชนาการและการคำนวณต่อ 1,000 กิโลแคลอรี").font = styles["title"]
    ws.cell(2, 1, f"{p['manufacturer']} | ข้อมูลสังเคราะห์เพื่อการทดสอบระบบ 100% | อย. {p['fda']}").font = styles["subtitle"]

    # 1. Calculation Standards
    ws.cell(4, 1, "1. ข้อมูลมาตรฐานการคำนวณและข้อกำหนดบรรจุภัณฑ์ (Calculation Standards & Packaging)").font = styles["section"]
    spec_rows = [
        ("น้ำหนักสุทธิต่อกระป๋อง (Net Weight per Can)", 400.0, "g", "#,##0.00"),
        ("ปริมาณอ้างอิงฐาน 100 กรัม (Base Quantity 100 g)", float(p["base_value"]), "g", "#,##0.00"),
        ("พลังงานต่อ 100 กรัม (Energy per 100 g)", float(p["base_kcal"]), "kcal", "#,##0.00"),
        ("ปริมาณต่อ 1 หน่วยบริโภค (Serving Size)", float(p["serving_g"]), "g", "#,##0.00"),
        ("พลังงานต่อ 1 หน่วยบริโภค (Energy per Serving)", "=ROUND((B8/B6)*B7, 2)", "kcal", "#,##0.00"),
        ("น้ำหนักผงที่ให้พลังงาน 1,000 kcal (Powder per 1,000 kcal)", "=ROUND((1000/B7)*B6, 2)", "g", "#,##0.00"),
        ("ตัวคูณปรับสัดส่วนจาก 100 กรัม (Scale Factor from 100 g)", "=ROUND(1000/B7, 6)", "เท่า", "0.000000"),
        ("ตัวคูณปรับสัดส่วนจาก 1 หน่วยบริโภค (Scale Factor from Serving)", "=ROUND(1000/B9, 6)", "เท่า", "0.000000"),
    ]
    for idx, (lbl, val, unit, fmt) in enumerate(spec_rows, start=5):
        c1 = ws.cell(idx, 1, lbl)
        c2 = ws.cell(idx, 2, val)
        c3 = ws.cell(idx, 3, unit)
        c1.font = styles["data"]
        c1.border = styles["border_data"]
        c2.font = styles["bold"] if idx in (9, 10) else styles["data"]
        c2.alignment = Alignment(horizontal="right", vertical="center")
        c2.number_format = fmt
        c2.border = styles["border_data"]
        if idx in (9, 10):
            c2.fill = styles["fill_sub"]
        c3.font = styles["data"]
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c3.border = styles["border_data"]

    # 2. Ingredients Breakdown
    ws.cell(14, 1, "2. ส่วนประกอบที่สำคัญเชิงปริมาณและการแปลงค่าต่อ 1,000 กิโลแคลอรี (Ingredients Breakdown)").font = styles["section"]
    for c_idx, h in enumerate(["ลำดับ", "รายการส่วนประกอบสำคัญ", "สัดส่วนตามฉลาก (% w/w)", "ต่อ 100 กรัม (g)", "ต่อ 1,000 kcal (g)", "บทบาทและหน้าที่ทางสรีรวิทยา"], 1):
        cell = ws.cell(15, c_idx, h)
        cell.font = styles["hdr"]
        cell.fill = styles["fill_hdr"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = styles["border_hdr"]

    start_ing = 16
    for idx, (name, pct, role) in enumerate(p["ingredients"], 1):
        r = start_ing + idx - 1
        ws.cell(r, 1, idx).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 2, name).alignment = Alignment(horizontal="left", vertical="center")
        c3 = ws.cell(r, 3, round(pct / 100.0, 4))
        c3.alignment = Alignment(horizontal="right", vertical="center")
        c3.number_format = "0.00%"
        c4 = ws.cell(r, 4, f"=ROUND(C{r}*100, 2)")
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c4.number_format = "#,##0.00"
        c5 = ws.cell(r, 5, f"=ROUND(C{r}*$B$10, 2)")
        c5.alignment = Alignment(horizontal="right", vertical="center")
        c5.number_format = "#,##0.00"
        ws.cell(r, 6, role).alignment = Alignment(horizontal="left", vertical="center")
        for c in range(1, 7):
            ws.cell(r, c).font = styles["data"]
            ws.cell(r, c).border = styles["border_data"]

    end_ing = start_ing + len(p["ingredients"]) - 1
    tot_ing_r = end_ing + 1
    ws.cell(tot_ing_r, 1, "รวม").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(tot_ing_r, 2, "รวมส่วนประกอบสำคัญทั้งหมดของสูตร (Total Formula)").alignment = Alignment(horizontal="left", vertical="center")
    c_t3 = ws.cell(tot_ing_r, 3, f"=ROUND(SUM(C{start_ing}:C{end_ing}), 4)")
    c_t3.alignment = Alignment(horizontal="right", vertical="center")
    c_t3.number_format = "0.00%"
    c_t4 = ws.cell(tot_ing_r, 4, f"=ROUND(SUM(D{start_ing}:D{end_ing}), 2)")
    c_t4.alignment = Alignment(horizontal="right", vertical="center")
    c_t4.number_format = "#,##0.00"
    c_t5 = ws.cell(tot_ing_r, 5, f"=ROUND(SUM(E{start_ing}:E{end_ing}), 2)")
    c_t5.alignment = Alignment(horizontal="right", vertical="center")
    c_t5.number_format = "#,##0.00"
    ws.cell(tot_ing_r, 6, "100% ตรงตามมาตรฐานสูตร").alignment = Alignment(horizontal="left", vertical="center")
    for c in range(1, 7):
        cell = ws.cell(tot_ing_r, c)
        cell.font = styles["total"]
        cell.fill = styles["fill_total"]
        cell.border = styles["border_total"]

    # 3. Nutrition Facts
    nut_sec_r = tot_ing_r + 2
    ws.cell(nut_sec_r, 1, "3. ตารางข้อมูลโภชนาการฉบับเต็มและการคำนวณต่อ 1,000 กิโลแคลอรี (Full Nutrition Facts)").font = styles["section"]
    nut_hdr_r = nut_sec_r + 1
    for c_idx, h in enumerate(["ลำดับ", "รายการสารอาหาร", "หน่วย", "ต่อ 100 กรัม", f"ต่อ 1 หน่วยบริโภค ({p['serving_g']:.1f} กรัม)", "ต่อ 1,000 kcal (คำนวณ)"], 1):
        cell = ws.cell(nut_hdr_r, c_idx, h)
        cell.font = styles["hdr"]
        cell.fill = styles["fill_hdr"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = styles["border_hdr"]

    cur_r = nut_hdr_r + 1
    r_energy = r_p = r_c = r_f = None
    for idx_n, (name, unit, v_base, _) in enumerate(p["nutrients"], 1):
        ws.cell(cur_r, 1, idx_n).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(cur_r, 2, name).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(cur_r, 3, unit).alignment = Alignment(horizontal="center", vertical="center")
        c4 = ws.cell(cur_r, 4, v_base if v_base is not None else 0.0)
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c4.number_format = "#,##0.00"
        c5 = ws.cell(cur_r, 5, f"=ROUND(D{cur_r}*($B$8/$B$6), 2)")
        c5.alignment = Alignment(horizontal="right", vertical="center")
        c5.number_format = "#,##0.00"
        c6 = ws.cell(cur_r, 6, f"=ROUND(D{cur_r}*$B$11, 2)")
        c6.alignment = Alignment(horizontal="right", vertical="center")
        c6.number_format = "#,##0.00"
        for c in range(1, 7):
            ws.cell(cur_r, c).font = styles["bold"] if c == 6 else styles["data"]
            ws.cell(cur_r, c).border = styles["border_data"]

        if "พลังงาน" in name and r_energy is None:
            r_energy = cur_r
        elif "โปรตีน" in name and r_p is None:
            r_p = cur_r
        elif "คาร์โบไฮเดรต" in name and r_c is None:
            r_c = cur_r
        elif "ไขมันทั้งหมด" in name and r_f is None:
            r_f = cur_r
        cur_r += 1

    # 4. Macronutrient Distribution Verification
    mac_sec_r = cur_r + 1
    ws.cell(mac_sec_r, 1, "4. การตรวจสอบสัดส่วนการกระจายพลังงาน (Macronutrient Distribution Verification)").font = styles["section"]
    mac_hdr_r = mac_sec_r + 1
    for c_idx, h in enumerate(["สารอาหารหลัก", "ปริมาณต่อ 1,000 kcal (g)", "ตัวคูณพลังงาน (kcal/g)", "พลังงานที่ได้ (kcal)", "สัดส่วนพลังงานจริง (%)", "สัดส่วนอ้างอิงฉลาก"], 1):
        cell = ws.cell(mac_hdr_r, c_idx, h)
        cell.font = styles["hdr"]
        cell.fill = styles["fill_hdr"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = styles["border_hdr"]

    r_mac_start = mac_hdr_r + 1
    r_mac_tot = r_mac_start + 3
    macro_items = [
        ("โปรตีน (Protein)", f"=ROUND(F{r_p}, 2)", 4, f"{p['nominal_pcf'][0]:.0f}%"),
        ("คาร์โบไฮเดรต (Carbohydrate)", f"=ROUND(F{r_c}, 2)", 4, f"{p['nominal_pcf'][2]:.0f}%"),
        ("ไขมัน (Fat)", f"=ROUND(F{r_f}, 2)", 9, f"{p['nominal_pcf'][1]:.0f}%"),
    ]
    for idx_m, (m_name, m_ref, mult, nom_pct) in enumerate(macro_items):
        rm = r_mac_start + idx_m
        ws.cell(rm, 1, m_name).alignment = Alignment(horizontal="left", vertical="center")
        cm2 = ws.cell(rm, 2, m_ref)
        cm2.alignment = Alignment(horizontal="right", vertical="center")
        cm2.number_format = "#,##0.00"
        ws.cell(rm, 3, mult).alignment = Alignment(horizontal="center", vertical="center")
        cm4 = ws.cell(rm, 4, f"=ROUND(B{rm}*C{rm}, 2)")
        cm4.alignment = Alignment(horizontal="right", vertical="center")
        cm4.number_format = "#,##0.00"
        cm5 = ws.cell(rm, 5, f"=ROUND(D{rm}/$D${r_mac_tot}*100, 2)")
        cm5.alignment = Alignment(horizontal="right", vertical="center")
        cm5.number_format = "#,##0.00"
        ws.cell(rm, 6, nom_pct).alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, 7):
            ws.cell(rm, c).font = styles["data"]
            ws.cell(rm, c).border = styles["border_data"]

    ws.cell(r_mac_tot, 1, "รวมพลังงานทั้งหมด (Total Macronutrient Energy)").alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(r_mac_tot, 2, "-").alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(r_mac_tot, 3, "-").alignment = Alignment(horizontal="center", vertical="center")
    ct4 = ws.cell(r_mac_tot, 4, f"=ROUND(SUM(D{r_mac_start}:D{r_mac_tot-1}), 2)")
    ct4.alignment = Alignment(horizontal="right", vertical="center")
    ct4.number_format = "#,##0.00"
    ct5 = ws.cell(r_mac_tot, 5, f"=ROUND(SUM(E{r_mac_start}:E{r_mac_tot-1}), 2)")
    ct5.alignment = Alignment(horizontal="right", vertical="center")
    ct5.number_format = "#,##0.00"
    ws.cell(r_mac_tot, 6, f"{p['nominal_pcf'][0]:.0f} : {p['nominal_pcf'][1]:.0f} : {p['nominal_pcf'][2]:.0f}").alignment = Alignment(horizontal="center", vertical="center")
    for c in range(1, 7):
        cell = ws.cell(r_mac_tot, c)
        cell.font = styles["total"]
        cell.fill = styles["fill_total"]
        cell.border = styles["border_total"]

    _autofit_columns(ws, min_col=1, max_col=6)


def generate_standalone_excel(product: Dict[str, Any], output_path: Path, font_name: str = "TH Sarabun New") -> None:
    """Generate Single Standalone Excel workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    styles = _get_styles(font_name)
    _build_product_sheet(ws, product, styles)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))


def generate_master_excel(
    brand_meta: Dict[str, Any],
    products: List[Dict[str, Any]],
    output_path: Path,
    font_name: str = "TH Sarabun New",
) -> None:
    """Generate multi-product Brand Master Excel workbook with comparison & energy summary sheets."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet
    styles = _get_styles(font_name)

    # 1. Sheet: 00_ภาพรวมเปรียบเทียบ_1000kcal
    ws_ov = wb.create_sheet(title="00_ภาพรวมเปรียบเทียบ_1000kcal")
    ws_ov.views.sheetView[0].showGridLines = True
    ws_ov.cell(1, 1, f"สรุปภาพรวมเปรียบเทียบผลิตภัณฑ์อาหารทางการแพทย์: {brand_meta['brand_name']}").font = styles["title"]
    ws_ov.cell(2, 1, f"{brand_meta['brand_name']} | ตารางเปรียบเทียบพารามิเตอร์บรรจุภัณฑ์และสารอาหารหลักต่อ 1,000 kcal").font = styles["subtitle"]

    ws_ov.cell(4, 1, "1. พารามิเตอร์มาตรฐานและการคำนวณต่อ 1,000 กิโลแคลอรี (Standard Parameters)").font = styles["section"]
    for c_idx, h in enumerate(["พารามิเตอร์มาตรฐาน (Parameter)", "ค่าอ้างอิง / สูตรคำนวณ", "หน่วย (Unit)", "รายละเอียดอ้างอิงตามฉลากจริง"], 1):
        cell = ws_ov.cell(5, c_idx, h)
        cell.font = styles["hdr"]
        cell.fill = styles["fill_hdr"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = styles["border_hdr"]

    p0 = products[0]
    sname = p0["id"]
    ov_param_rows = [
        ("น้ำหนัก 1 หน่วยบริโภคตามฉลาก (Serving Size)", f"=ROUND('{sname}'!B8, 2)", "g", "ผง 1 หน่วยบริโภคตามฉลากมาตรฐาน"),
        ("ปริมาณอ้างอิงฐาน 100 กรัม (Base Quantity 100 g)", f"=ROUND('{sname}'!B6, 2)", "g", "ฐานการคำนวณ 100 กรัม"),
        ("พลังงานต่อ 100 กรัม (Energy per 100 g)", f"=ROUND('{sname}'!B7, 2)", "kcal", "ความเข้มข้นพลังงานต่อ 100g"),
        ("พลังงานต่อ 1 หน่วยบริโภค (Energy per Serving)", f"=ROUND('{sname}'!B9, 2)", "kcal", "พลังงานต่อหน่วยบริโภค"),
        ("เป้าหมายพลังงานมาตรฐานทางการแพทย์ (Target Energy)", 1000.0, "kcal", "เกณฑ์เปรียบเทียบสารอาหารต่อ 1,000 kcal"),
        ("น้ำหนักผงที่ต้องใช้ต่อ 1,000 kcal (Powder per 1,000 kcal)", f"=ROUND('{sname}'!B10, 2)", "g", "คำนวณจาก (1,000 / พลังงานต่อ 100g) * 100"),
        ("ตัวคูณปรับสัดส่วนสารอาหาร (Scale Factor from 100 g)", f"=ROUND('{sname}'!B11, 6)", "เท่า", "ตัวคูณเทียบเท่าต่อ 1,000 kcal"),
    ]
    for idx, (p_lbl, p_val, p_unit, p_desc) in enumerate(ov_param_rows, start=6):
        c1 = ws_ov.cell(idx, 1, p_lbl)
        c2 = ws_ov.cell(idx, 2, p_val)
        c3 = ws_ov.cell(idx, 3, p_unit)
        c4 = ws_ov.cell(idx, 4, p_desc)
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2.alignment = Alignment(horizontal="right", vertical="center")
        c2.number_format = "0.000000" if idx == 12 else "#,##0.00"
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c4.alignment = Alignment(horizontal="left", vertical="center")
        for c in (c1, c2, c3, c4):
            c.font = styles["bold"] if idx in (10, 11) else styles["data"]
            c.border = styles["border_data"]
            if idx in (10, 11):
                c.fill = styles["fill_sub"]

    ws_ov.cell(14, 1, "2. ตารางสรุปภาพรวมผลิตภัณฑ์และการกระจายพลังงานต่อ 1,000 kcal").font = styles["section"]
    ov_headers = [
        "รหัส", "ชื่อผลิตภัณฑ์", "น้ำหนัก/กระป๋อง (g)", "พลังงาน/100g (kcal)",
        "พลังงาน/หน่วยบริโภค (kcal)", "ผงต่อ 1,000 kcal (g)",
        "โปรตีน/1,000 kcal (g)", "คาร์บ/1,000 kcal (g)", "ไขมัน/1,000 kcal (g)",
        "%P จริง", "%C จริง", "%F จริง", "สัดส่วน P:F:C"
    ]
    for c_idx, h in enumerate(ov_headers, 1):
        cell = ws_ov.cell(15, c_idx, h)
        cell.font = styles["hdr"]
        cell.fill = styles["fill_hdr"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = styles["border_hdr"]

    # In single-product sample, find macro rows dynamically from product sheet
    # Protein at row 28, Fat at 29, Carb at 32 in standard template
    row_prod = 16
    ov_vals = [
        sname,
        p0["trade_th"],
        f"=ROUND('{sname}'!B5, 2)",
        f"=ROUND('{sname}'!B7, 2)",
        f"=ROUND('{sname}'!B9, 2)",
        f"=ROUND('{sname}'!B10, 2)",
        f"=ROUND('{sname}'!F28, 2)",
        f"=ROUND('{sname}'!F32, 2)",
        f"=ROUND('{sname}'!F29, 2)",
        f"=ROUND('{sname}'!E63, 2)",
        f"=ROUND('{sname}'!E64, 2)",
        f"=ROUND('{sname}'!E65, 2)",
        f"{p0['nominal_pcf'][0]:.0f} : {p0['nominal_pcf'][1]:.0f} : {p0['nominal_pcf'][2]:.0f}",
    ]
    for c_idx, val in enumerate(ov_vals, 1):
        cell = ws_ov.cell(row_prod, c_idx, val)
        cell.font = styles["bold"] if c_idx == 1 else styles["data"]
        cell.border = styles["border_data"]
        if c_idx in (1, 13):
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c_idx == 2:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "#,##0.00"
    _autofit_columns(ws_ov, min_col=1, max_col=13)

    # 2. Sheet: 01_สรุปพลังงาน_1000kcal
    ws_en = wb.create_sheet(title="01_สรุปพลังงาน_1000kcal")
    ws_en.views.sheetView[0].showGridLines = True
    ws_en.cell(1, 1, f"สรุปสัดส่วนการกระจายพลังงานสารอาหารหลัก: {brand_meta['brand_name']}").font = styles["title"]
    ws_en.cell(2, 1, "การกระจายพลังงาน โปรตีน : คาร์โบไฮเดรต : ไขมัน เทียบเคียงมาตรฐานต่อ 1,000 kcal").font = styles["subtitle"]

    en_headers = [
        "รหัสผลิตภัณฑ์", "ชื่อทางการค้า", "พลังงานเป้าหมาย (kcal)",
        "โปรตีน (g)", "พลังงานจากโปรตีน (kcal)", "% พลังงานโปรตีน",
        "คาร์บ (g)", "พลังงานจากคาร์บ (kcal)", "% พลังงานคาร์บ",
        "ไขมัน (g)", "พลังงานจากไขมัน (kcal)", "% พลังงานไขมัน",
        "รวมพลังงาน (kcal)", "รวมสัดส่วน (%)"
    ]
    for c_idx, h in enumerate(en_headers, 1):
        cell = ws_en.cell(4, c_idx, h)
        cell.font = styles["hdr"]
        cell.fill = styles["fill_hdr"]
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = styles["border_hdr"]

    r_en = 5
    en_vals = [
        sname,
        p0["trade_th"],
        1000.0,
        f"=ROUND('{sname}'!B63, 2)",
        f"=ROUND('{sname}'!D63, 2)",
        f"=ROUND('{sname}'!E63, 2)",
        f"=ROUND('{sname}'!B64, 2)",
        f"=ROUND('{sname}'!D64, 2)",
        f"=ROUND('{sname}'!E64, 2)",
        f"=ROUND('{sname}'!B65, 2)",
        f"=ROUND('{sname}'!D65, 2)",
        f"=ROUND('{sname}'!E65, 2)",
        f"=ROUND('{sname}'!D66, 2)",
        f"=ROUND('{sname}'!E66, 2)",
    ]
    for c_idx, val in enumerate(en_vals, 1):
        cell = ws_en.cell(r_en, c_idx, val)
        cell.font = styles["bold"] if c_idx in (1, 13, 14) else styles["data"]
        cell.border = styles["border_data"]
        if c_idx == 1:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c_idx == 2:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = "#,##0.00"
    _autofit_columns(ws_en, min_col=1, max_col=14)

    # 3. Product Sheet: Formula_Alpha
    ws_prod = wb.create_sheet(title=sname)
    _build_product_sheet(ws_prod, p0, styles)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(output_path))


def build_all(output_dir: Optional[Path | str] = None) -> Path:
    """
    Build all 4 synthetic deliverables and execute unit normalization.

    Returns the resolved output directory Path.
    """
    if output_dir is None:
        target_dir = THIS_DIR / "output" / "NutriSample_Co"
    else:
        target_dir = Path(output_dir).resolve()

    target_dir.mkdir(parents=True, exist_ok=True)

    meta = synthetic_brand_data.BRAND_META
    products = synthetic_brand_data.PRODUCTS
    p0 = products[0]

    prod_dir = target_dir / p0["folder"]
    prod_dir.mkdir(parents=True, exist_ok=True)

    # 1. Set 1 Markdown
    set1_path = prod_dir / "01_Formula_Alpha_Label.md"
    generate_set1_markdown(p0, set1_path)

    # 2. Set 2 Markdown
    set2_path = prod_dir / "02_Formula_Alpha_Ingredients_Nutrition_1000kcal.md"
    generate_set2_markdown(p0, set2_path)

    # 3. Standalone Excel
    standalone_path = prod_dir / "Formula_Alpha_Nutrition_1000kcal.xlsx"
    generate_standalone_excel(p0, standalone_path)

    # 4. Brand Master Excel
    master_path = target_dir / meta["master_excel"]
    generate_master_excel(meta, products, master_path)

    # 5. Apply Unit Normalization to workspace
    normalize_workspace(target_dir)

    return target_dir


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Builder for Synthetic Formula Worked Example")
    parser.add_argument("--output-dir", default=None, help="Output directory for synthetic deliverables")
    args = parser.parse_args(argv)

    out_p = build_all(args.output_dir)
    print(f"=== Synthetic Formula Deliverables Generated Successfully at: {out_p} ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
