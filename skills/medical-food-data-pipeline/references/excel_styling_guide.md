# Excel Styling & OpenPyXL Implementation Guide

คู่มือมาตรฐานการจัดรูปแบบสมุดงาน Excel และตัวอย่างการเขียนสคริปต์ด้วยไลบรารี `openpyxl` ตามมาตรฐาน Corporate Blue Theme สำหรับกลุ่มผลิตภัณฑ์อาหารทางการแพทย์ (Medical Food / Enteral Nutrition) รองรับการใช้งานข้ามแพลตฟอร์มอย่างสมบูรณ์แบบ (Windows, macOS, Linux / LibreOffice)

---

## 1. จานสีมาตรฐาน Corporate Blue (Color Palette)

| ส่วนประกอบ (Component) | รหัสสี HEX | สีตัวอักษร | วัตถุประสงค์และการใช้งาน |
| :--- | :---: | :---: | :--- |
| **Primary Header Background** | `#1F4E78` / `#1B365D` | `#FFFFFF` (Bold) | น้ำเงินเข้ม สื่อถึงความน่าเชื่อถือระดับสถาบันการแพทย์ |
| **Subheader / Section Bar** | `#D9E1F2` | `#1F4E78` / `#1B365D` | ฟ้าอ่อนพาสเทล แยกกลุ่มสารอาหารหลักหรือหัวข้อย่อย |
| **Category Header** | `#F2F2F2` | `#2F5597` (Bold) | เทาอ่อนพาสเทล สำหรับแบ่งหมวดหมู่ย่อยในตาราง |
| **Interactive / Highlight** | `#FFF2CC` | `#000000` | เหลืองนวล ไฮไลต์เซลล์นำเข้าข้อมูลที่อนุญาตให้แก้ไข |
| **Total / Summary Row** | `#E2EFDA` | `#000000` (Bold) | เขียวอ่อนพาสเทล แถวผลรวมหรือแถวตรวจสอบสมดุล |
| **Zebra Striping (ถ้ามี)** | `#F9FBFD` | `#000000` | ฟ้าขาวนวลตา สลับบรรทัดเพื่อเพิ่มความสบายตา |
| **Gridlines & Borders** | `#D3D3D3` / `#D9D9D9` | - | เส้นขอบบางรอบเซลล์ (Thin border) |
| **Total Double Border** | `#1F4E78` / `#1B365D` | - | เส้นคู่ (Double line) ใต้แถวผลรวมล่างสุด |

---

## 2. การจัดการฟอนต์ข้ามแพลตฟอร์ม (Cross-Platform Typography)

มาตรฐานของสมุดงานรองรับทั้งฟอนต์ **`TH Sarabun New`** (มาตรฐานราชการและกระทรวงสาธารณสุขไทย) และ **`Cordia New`** (มาตรฐานชุดเดิมของระบบเอกสารทางการแพทย์):

1. **ลำดับชั้นขนาดตัวอักษร (Hierarchy):**
   * **ชื่อเรื่องสมุดงาน (Workbook Title):** 16 pt (Bold)
   * **หัวข้อส่วน / Section Header:** 14 pt (Bold)
   * **หัวตาราง / Column Header:** 12–13 pt (Bold)
   * **เนื้อหาข้อมูล / Data Cells:** 11–12 pt (Regular)
   * **หมายเหตุและข้อสมมติฐานทางคลินิก (Notes):** 10–11 pt (Italic)

2. **ระบบตรวจจับฟอนต์อัตโนมัติ (Automatic Font Detection):**
   * เครื่องมือ `apply_unit_normalization.py` ถูกออกแบบให้ตรวจจับฟอนต์ส่วนใหญ่ที่มีอยู่เดิมในสมุดงานโดยอัตโนมัติเมื่อละเว้นแฟล็ก `--font`
   * หากไม่พบข้อมูลฟอนต์เดิม ระบบจะใช้ฟอนต์สำรองเริ่มต้นเป็น **`TH Sarabun New`**
   * ผู้ใช้สามารถระบุฟอนต์ที่ต้องการอย่างเจาะจงได้ผ่านแฟล็กคำสั่ง เช่น `--font "Cordia New"` หรือ `--font "TH Sarabun New"`

3. **รูปแบบตัวเลขมาตรฐาน (Number Formats):**
   * **ทศนิยม 2 ตำแหน่ง:** `#,##0.00` (สารอาหาร, น้ำหนักผงกรัม, พลังงาน kcal)
   * **จำนวนเต็ม:** `#,##0` (พลังงานรวม, ปริมาณบรรจุภัณฑ์กรัมรวม)
   * **เปอร์เซ็นต์:** `0.00%` (สัดส่วน % w/w, สัดส่วนพลังงาน % Energy)

---

## 3. การคำนวณความกว้างคอลัมน์และป้องกันข้อความล้น (Column Autofit & Glyph Width Margin)

ปัญหาสำคัญในการแสดงผลภาษาไทยข้ามแพลตฟอร์มคือ ฟอนต์ภาษาไทยมาตรฐานของ Windows เช่น `Cordia New` หรือ `TH Sarabun New` มักไม่มีติดตั้งมาพร้อมกับ **macOS** หรือระบบปฏิบัติการลินุกซ์ ทำให้ซอฟต์แวร์สเปรดชีตอย่าง Excel บน macOS หรือ **LibreOffice** นำฟอนต์สำรอง (Fallback Fonts) มาเรนเดอร์แทน ซึ่งอาจมีความกว้างของอักขระ สระ และวรรณยุกต์มากกว่าปกติ จนส่งผลให้ข้อความยาวล้นเซลล์หรือถูกตัดทอน (Text Clipping):

* **สูตรคำนวณความกว้างแบบ UTF-8 Glyphs:**
  $$\text{Visual Width} = \sum (\text{ch} > 127 \text{ ? } 1.6 : 1.0) + 5.0$$
* สำหรับอักขระภาษาไทย (Non-ASCII Unicode) ให้คิดน้ำหนักความกว้าง $\ge 1.6\times$ ของตัวอักษรละติน
* **ส่วนเผื่อขอบบังคับ (`+5` column width margin):** เพิ่มระยะปลอดภัยอย่างน้อย **`+5`** หน่วยความกว้างคอลัมน์เสมอ เพื่อป้องกันการตัดทอนข้อความเมื่อเปิดดูบน **macOS** หรือ **LibreOffice**

---

## 4. มาตรฐานสูตรไดนามิกและการเชื่อมโยงข้อมูล (Dynamic Formula Standards)

1. **กฎการครอบคลุมด้วย `=ROUND(...)` 100%:**
   * ทุกเซลล์ที่เป็นการคำนวณทางคณิตศาสตร์ต้องขึ้นต้นด้วย `=ROUND(..., 2)` เพื่อป้องกันปัญหาความคลาดเคลื่อนของทศนิยมสะสม (Floating Point Precision Drift)
   * ห้ามใส่ค่าตัวเลขฮาร์ดโค้ดลงในเซลล์ผลลัพธ์เด็ดขาด

2. **สูตรมาตรฐานประจำสมุดงาน:**
   * **ผงต่อ 1,000 kcal (Powder per 1,000 kcal):**
     ```excel
     =ROUND((1000/D{energy_row})*100, 2)
     ```
   * **ส่วนประกอบต่อ 1,000 kcal จาก % w/w:**
     ```excel
     =ROUND(C{row}*$E${powder_row}, 2)
     ```
   * **สารอาหารต่อ 1,000 kcal:**
     ```excel
     =ROUND((D{row}/$D${energy_row})*1000, 2)
     ```
   * **พลังงานจากสารอาหารหลัก (kcal):**
     ```excel
     =ROUND(B{row}*C{row}, 2)
     ```
   * **สัดส่วนพลังงานจริง (% Energy):**
     ```excel
     =ROUND((D{row}/$D${total_energy_row})*100, 2)
     ```
   * **ผลรวมสารอาหารหรือพลังงาน:**
     ```excel
     =ROUND(SUM(B{start}:B{end}), 2)
     ```
   * **การคำนวณค่า Normalized ตามประกาศ สธ. 445:**
     ```excel
     =ROUND(raw_1000kcal_cell * factor_cell, 2)
     ```

3. **สุขอนามัยการอ้างอิงข้ามแผ่นงาน (Cross-Sheet Formula Hygiene):**
   * หลีกเลี่ยงการใช้ `=IF(...)` ซ้อนหลายชั้นเพื่อเปรียบเทียบข้อความดิบข้ามชีตโดยไม่จำเป็น
   * ห้ามอ้างอิงเซลล์ที่ไม่มีอยู่จริง และห้ามฮาร์ดโค้ดตัวเลขตัวคูณไว้ในสูตร (ให้วางตัวคูณในคอลัมน์เฉพาะแล้วอ้างอิงตำแหน่งเซลล์)
   * ดึงค่าพลังงานและสัดส่วนเข้าสู่แผ่นงานสรุปภาพรวมโดยตรง เช่น:
     ```excel
     ='ProductName'!D{energy_row}
     ='ProductName'!E{macro_row}/100
     ```

4. **การจัดตำแหน่งและการตรึงแนว (Layout, Wrap Text & Freeze Panes):**
   * ข้อความและตัวเลขต้องตั้งค่า `vertical="center"`
   * คอลัมน์รายการส่วนประกอบและชื่อสารอาหารให้เปิด `wrap_text=True`
   * ตรึงแนวแผ่นงาน (Freeze Panes) ใต้แถวหัวตารางเสมอ เช่น `ws.freeze_panes = "B2"` หรือ `ws.freeze_panes = "C4"` เพื่อความสะดวกในการเลื่อนดูข้อมูลทางคลินิก
   * เปิดเส้นตารางเสมอ: `ws.views.sheetView[0].showGridLines = True`

5. **การป้องกันแผ่นงาน (Sheet Protection):**
   * หากมีการล็อกแผ่นงาน (Protect Sheet) ต้องปลดล็อกเซลล์นำเข้าข้อมูล (Input Cells) ให้อยู่ในสถานะ `cell.protection = Protection(locked=False)` และไม่ต้องตั้งรหัสผ่าน (No Password) เพื่อความยืดหยุ่นในการตรวจสอบของผู้เชี่ยวชาญ

---

## 5. โค้ดต้นแบบ OpenPyXL มาตรฐาน (Cross-Platform OpenPyXL Template)

```python
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter

def create_standard_medical_food_workbook(font_name="TH Sarabun New"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "00_ภาพรวมผลิตภัณฑ์"
    ws.views.sheetView[0].showGridLines = True

    # 1. Palette & Styles
    font_title = Font(name=font_name, size=16, bold=True, color="1F4E78")
    font_section = Font(name=font_name, size=14, bold=True, color="1F4E78")
    font_header = Font(name=font_name, size=13, bold=True, color="FFFFFF")
    font_bold = Font(name=font_name, size=12, bold=True)
    font_regular = Font(name=font_name, size=12)
    font_note = Font(name=font_name, size=11, italic=True)

    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_cat = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_total = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin_side = Side(border_style="thin", color="D3D3D3")
    double_side = Side(border_style="double", color="1F4E78")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    total_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    # 2. Freeze Panes
    ws.freeze_panes = "B2"

    return wb

def autofit_thai_columns(ws, min_col=1, max_col=None, margin=5):
    """
    Calculate column width supporting Thai vowels, tone marks, and fallback font rendering
    on macOS and LibreOffice with a safety margin of +5 width units.
    """
    if max_col is None:
        max_col = ws.max_column

    for col in range(min_col, max_col + 1):
        col_letter = get_column_letter(col)
        max_len = 0.0
        for row in range(1, ws.max_row + 1):
            val = ws.cell(row=row, column=col).value
            if val is not None:
                val_str = str(val)
                if val_str.startswith("="):
                    val_str = "123,456.78"
                for line in val_str.split("\n"):
                    # 1.6x factor for non-ASCII Thai characters
                    line_len = sum(1.6 if ord(c) > 127 else 1.0 for c in line)
                    if line_len > max_len:
                        max_len = line_len
        # Apply +5 width margin
        ws.column_dimensions[col_letter].width = min(max(round(max_len + margin, 1), 12.0), 60.0)
```
