# Monthly Sales Report Specification

## 1. Required 7-Section Structure (Max 3 A4 Pages)

1. **หน้าปก (Cover — Page 1)**:
   - Report title, Month/Year, Report period, Preparation date, Preparer/Author, Approver, Company/Department.
2. **บทสรุปสำหรับผู้บริหาร (Executive Summary — Page 2 Top, ~half page)**:
   - 4 KPI highlights (Total Sales, % vs Target, MoM Growth %, Customer Satisfaction).
   - 3–5 bullet-point highlights with bold key figures.
3. **ผลการดำเนินงาน (Performance — Page 2 Bottom)**:
   - Target vs Actual bar chart (Previous Month, Target, Actual).
   - Comparison table (Sales, MoM, New Customers, Active Customers, Churn, Satisfaction).
4. **สินค้าขายดี Top 4 (Top Products — Page 3 Top-Left)**:
   - Table with Rank, Product, Quantity, Sales (THB), and Total row.
5. **ผลงานทีมขาย (Sales Team — Page 3 Top-Right)**:
   - Table with Rank, Salesperson (with `TOP` badge for #1), Sales (THB), Share %, and Total row.
6. **ปัญหาและแนวทางแก้ไข (Issues & Actions — Page 3 Middle)**:
   - Bullet points pairing each issue with a concrete corrective action.
7. **แผนงานเดือนหน้า (Next Month Plan — Page 3 Bottom)**:
   - Checklist items (`☐`) with owner/responsible team.

## 2. Input JSON Schema for `scripts/generate_report.py`

```json
{
  "company": "บริษัท ตัวอย่าง จำกัด",
  "department": "ฝ่ายขายและการตลาด · Sales & Marketing",
  "doc_tag_cover": "CONFIDENTIAL\nFY2569 / Q1",
  "doc_tag_header": "SALES REPORT\nมี.ค. 2569",
  "title": "รายงานสรุปผลการขาย",
  "month_label": "ประจำเดือนมีนาคม 2569",
  "short_month": "มีนาคม 2569",
  "period": "1 – 31 มีนาคม 2569",
  "prepared_date": "3 เมษายน 2569",
  "prepared_by": "ฝ่ายขายและการตลาด",
  "approved_by": "ผู้จัดการฝ่ายขาย",
  "kpis": {
    "total_sales": 1925000,
    "target_sales": 1800000,
    "prev_sales": 1655000,
    "prev_month_label": "ก.พ. 69",
    "actual_month_label": "ผลจริง มี.ค.",
    "new_customers": 18,
    "active_customers": 127,
    "churn_customers": 3,
    "csat": 4.3
  },
  "highlights": [
    "<b>ยอดขายรวม 1,925,000 บาท</b> สูงกว่าเป้าหมาย (1,800,000 บาท) คิดเป็น <span class=\"pos\">+6.9%</span>",
    "<b>เติบโตจากเดือนกุมภาพันธ์</b> (1,655,000 บาท) ในอัตรา <span class=\"pos\">+16.3% MoM</span>",
    "<b>ขยายฐานลูกค้า</b> ได้ลูกค้าใหม่ 18 ราย คงฐาน Active 127 ราย ลูกค้ายกเลิก (Churn) เพียง 3 ราย",
    "<b>สินค้าเรือธง</b> กลุ่มโน้ตบุ๊กทำยอดสูงสุด 1,000,000 บาท — กว่าครึ่งของยอดขายทั้งเดือน"
  ],
  "top_products": [
    {"name": "โน้ตบุ๊ก", "qty": 40, "sales": 1000000},
    {"name": "จอมอนิเตอร์", "qty": 50, "sales": 425000},
    {"name": "แท็บเล็ต", "qty": 25, "sales": 300000},
    {"name": "คีย์บอร์ด", "qty": 80, "sales": 200000}
  ],
  "sales_team": [
    {"name": "สมหมาย", "sales": 1200000, "top": true},
    {"name": "สุดา", "sales": 425000, "top": false},
    {"name": "วิรัช", "sales": 300000, "top": false}
  ],
  "issues": [
    {"issue": "สต็อกโน้ตบุ๊กรุ่น A1 เหลือน้อย (คงเหลือ 15 เครื่อง)", "action": "เร่งสั่งสต็อกเพิ่ม 50 เครื่อง และกำหนดจุดสั่งซื้อซ้ำ (Reorder Point) เพื่อกันของขาด"},
    {"issue": "ลูกค้า 3 ราย ยกเลิกเนื่องจากเรื่องราคา (Churn)", "action": "ทบทวนกลยุทธ์ราคา จัดแพ็กเกจคุ้มค่า และเปิดตัว Loyalty Program เพื่อรักษาฐานลูกค้า"},
    {"issue": "ระบบ CRM มีปัญหา Downtime 2 ชั่วโมง", "action": "ประสานทีม IT ตรวจสอบเสถียรภาพระบบ และจัดทำแผนสำรองข้อมูล/ระบบสำรอง"}
  ],
  "next_month_title": "แผนงานเดือนเมษายน 2569",
  "next_month_plan": [
    {"task": "สั่งสต็อกโน้ตบุ๊กเพิ่ม 50 เครื่อง", "owner": "จัดซื้อ"},
    {"task": "เปิดตัว Loyalty Program สำหรับลูกค้าเดิม", "owner": "การตลาด"},
    {"task": "อบรมทีมขายหัวข้อ Consultative Selling", "owner": "หัวหน้าทีมขาย"},
    {"task": "ติดตามแก้ไขเสถียรภาพระบบ CRM", "owner": "IT"}
  ]
}
```

## 3. OOXML Thai Font Embedding & Proofing Rules

- **Font**: `Anuphan` (`Anuphan-Regular.ttf`, `Anuphan-Bold.ttf`).
- **DOCX Embedding**:
  - Store `.ttf` files in `word/fonts/font1.ttf` (Regular) and `word/fonts/font2.ttf` (Bold).
  - Register relationships in `word/_rels/fontTable.xml.rels` with type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/font`.
  - Reference `w:embedRegular` and `w:embedBold` inside `w:font[@w:name="Anuphan"]` in `word/fontTable.xml`.
  - Enable `<w:embedTrueTypeFonts/>` in `word/settings.xml`.
- **Thai Language & Spell-check**:
  - Set `<w:lang w:val="en-US" w:bidi="th-TH"/>` in `rPr`.
  - Set `<w:noProof/>` **only** when the run's text contains Thai characters (`\u0E00`–`\u0E7F`). Pure English/numeric runs keep spell-check enabled.
