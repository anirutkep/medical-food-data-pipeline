# Medical Food Data Pipeline Skill
*(ระบบสกัดและประมวลผลข้อมูลอาหารทางการแพทย์และสารอาหารทางสายให้อาหาร)*

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Version: 0.1.1](https://img.shields.io/badge/Version-0.1.1-informational.svg)](pyproject.toml)
[![Python: >=3.10](https://img.shields.io/badge/Python-3.10%2B-brightgreen.svg)](https://www.python.org/)
[![Thai FDA: MOPH 445](https://img.shields.io/badge/Thai%20FDA-MOPH%20Notification%20445-orange.svg)](https://www.moph.go.th/)

---

## English Overview

**`medical-food-data-pipeline`** is an agent skill and data engineering framework designed for clinical nutritionists, hospital pharmacists, and medical food manufacturers. It automates 360° product label ingestion, OCR text extraction, regulatory standardisation, and 100% dynamic Excel workbook generation compliant with Thai Food and Drug Administration (FDA) **MOPH Notification No. 445 (B.E. 2566 / 2023)**.

### Key Capabilities
- **Regulatory Compliance**: Aligns nutrient classifications, mandatory minerals, and vitamins with Thai FDA **MOPH Notification No. 445**.
- **1,000 kcal Normalization**: Automatically calculates dilution factors, macro distributions (Caloric % Distribution: Protein 4 kcal/g, Fat 9 kcal/g, Carbohydrate 4 kcal/g), and osmolarity.
- **Dynamic Excel Generation**: Emits formula-driven Excel models using `openpyxl` with zero static constants, formatted with professional Thai fonts (`TH Sarabun New` or `Cordia New`).
- **Cross-Platform Agent Support**: Native manifests for **Claude Code**, **Codex**, and universal **Agent Skills** CLI.

---

## ภาพรวมภาษาไทย (Thai Overview)

**`medical-food-data-pipeline`** คือทักษะสำหรับระบบเอเจนต์ปัญญาประดิษฐ์ (AI Agent Skill) และไปป์ไลน์วิศวกรรมข้อมูลทางการแพทย์ พัฒนาขึ้นเพื่อช่วยนักกำหนดอาหาร เภสัชกรโรงพยาบาล และผู้ผลิตอาหารทางการแพทย์ ในการสกัดและตรวจสอบข้อมูลผลิตภัณฑ์อาหารทางการแพทย์ (Medical Food / Enteral Nutrition) จากภาพถ่ายฉลากจริง 360 องศา หรือเอกสารทางการ

### ความสามารถหลัก
1. **มาตรฐานกระทรวงสาธารณสุข**: สอดคล้องตามประกาศกระทรวงสาธารณสุข **ฉบับที่ 445 (พ.ศ. 2566)** เรื่อง อาหารทางการแพทย์
2. **การปรับสัดส่วนมาตรฐาน 1,000 kcal**: คำนวณ Dilution Factor, สัดส่วนพลังงานจากสารอาหารหลัก (% Distribution), และสารอาหารรองอย่างแม่นยำ
3. **การสร้างไฟล์ Excel ด้วยสูตรไดนามิก 100%**: สร้างสมุดงาน Excel ผ่าน `openpyxl` โดยใช้สูตรคำนวณทั้งหมด ไม่ฝังค่าตัวเลขแบบตายตัว (Hardcoded) พร้อมจัดรูปแบบตัวอักษรภาษาไทยมาตรฐาน ได้แก่ **TH Sarabun New** และ **Cordia New**
4. **รองรับเอเจนต์หลากหลายแพลตฟอร์ม**: มีไฟล์ Manifest สมบูรณ์ทั้งสำหรับ **Claude Code**, **Codex**, และ **Agent Skills**

---

## Installation & Setup / วิธีการติดตั้ง

### 1. Agent Skills CLI (Universal / Windows & macOS)
Install the skill into your project using `npx`:

```bash
# macOS / Linux
npx skills add anirutkep/medical-food-data-pipeline

# Windows (Command Prompt / PowerShell)
npx skills add anirutkep/medical-food-data-pipeline
```

### 2. Claude Code Plugin Marketplace
Add this repository to your Claude Code environment:

```text
/plugin marketplace add anirutkep/medical-food-data-pipeline
```

Or install from a cloned local repository root:
```bash
claude-code plugin install .
```

### 3. Codex CLI & Antigravity
For Codex CLI or Antigravity agents (from a cloned local repository root):
```bash
codex skill add ./skills/medical-food-data-pipeline
```

---

## Python Environment & Dependencies

This pipeline requires **Python >= 3.10**. Only one core runtime dependency is required:

- **`openpyxl>=3.1.0`** (No heavy dependencies like `torch`, `easyocr`, or `paddleocr` are forced into the base environment).

### Using `uv` (Recommended)
```bash
# Create virtual environment and install dependencies
uv venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
uv pip install -r requirements.txt
```

### Using standard `python` & `pip`
```bash
# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Windows
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

---

## OCR & Operating System Architecture Notes

Medical food labels in Thailand often present bilingual text (Thai and English). The pipeline provides multi-tier OCR fallback strategies tailored for different operating systems:

| OS Platform | Preferred OCR Engine | Notes & Setup |
| :--- | :--- | :--- |
| **macOS** | **Apple Vision Framework** (`pyobjc-framework-Vision`) | Native hardware-accelerated bilingual OCR supporting Thai script (`th-Thai`) with high accuracy and low resource usage. |
| **Windows** | **Windows WinRT** (`Windows.Media.Ocr`) | Native Windows OCR API. Requires the Windows Thai Language Pack (`Settings > Time & Language > Language > Add Thai`). |
| **Universal Fallback** | **Visual Language Model (VLM)** | Zero local ML footprint: delegates complex, low-contrast, or curved packaging text to multimodal LLM vision tools. |

---

## Typography & Excel Font Standards

All exported Excel workbooks (`.xlsx`) generated by `openpyxl` automatically apply formal Thai typographic standards:
- **`TH Sarabun New`** (Default 16pt / 14pt body): Standard font for Thai official and medical documentation.
- **`Cordia New`** (Fallback 14pt): Widely compatible alternative across enterprise Windows and Microsoft Office environments.

---

## Synthetic Quickstart Command

To build the anonymized synthetic example (`NutriSample_Co` / `Formula_Alpha`), run the automated clinical audit, and execute the full test suite from the repository root (`python` on Windows, `python3` on macOS/Linux):

```bash
# 1. Build synthetic deliverables
python3 examples/synthetic_formula/build_synthetic_example.py

# 2. Run the automated clinical & parity audit
python3 skills/medical-food-data-pipeline/scripts/pipeline_audit.py \
  --data examples/synthetic_formula/synthetic_brand_data.py \
  --workspace examples/synthetic_formula/output/NutriSample_Co

# 3. Run the full test suite
python3 -m unittest discover -s tests -t . -v
```

---

## Clinical & Regulatory Disclaimer (ข้อจำกัดความรับผิดชอบทางคลินิก)

- **English:** This project is a data extraction, normalization, and workbook-formatting tool—**not a certified medical device or standalone clinical decision support system**. All extracted nutrient values, unit conversions, dilutions, and Excel outputs must be independently verified against the physical product packaging, official manufacturer documentation, and applicable Thai FDA / MOPH regulations. Patient formula selection and enteral nutrition prescriptions must remain under the clinical judgment and responsibility of qualified healthcare professionals (registered dietitians, pharmacists, or physicians).
- **ภาษาไทย:** ผลลัพธ์จากระบบนี้เป็นเพียงเครื่องมือช่วยสกัด ปรับมาตรฐาน และจัดรูปแบบข้อมูลเบื้องต้น **ไม่ใช่เครื่องมือวินิจฉัยหรือสั่งการรักษาทางการแพทย์โดยอัตโนมัติ** ค่าสารอาหาร ปริมาณการเจือจาง สูตร Excel และเอกสารทั้งหมดที่ได้ **ต้องตรวจสอบเทียบกับฉลากผลิตภัณฑ์จริง เอกสารทางการ และประกาศกระทรวงสาธารณสุขที่เกี่ยวข้องเสมอ** และการนำไปใช้ดูแลผู้ป่วยต้องอยู่ภายใต้ดุลยพินิจของบุคลากรทางการแพทย์ (นักกำหนดอาหาร เภสัชกร หรือแพทย์) เท่านั้น

---

## License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.
Copyright (c) 2026 `Anirut Kepan, RD`.
