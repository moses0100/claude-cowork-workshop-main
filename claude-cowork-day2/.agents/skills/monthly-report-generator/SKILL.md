---
name: monthly-report-generator
description: >-
  Generate executive-ready Thai monthly sales reports in both Word (.docx) and PDF
  formats (3 A4 pages, Anuphan font embedded) from raw monthly sales data using
  Formal or Modern corporate templates.
  Use when the user says "ทำรายงานประจำเดือน", "สรุปยอดขายเดือนนี้",
  "สร้าง monthly report", "ทำรายงานยอดขายจากไฟล์นี้",
  "สรุปผลการดำเนินงานรายเดือนเป็น Word/PDF", or attaches raw monthly sales data
  and asks for a formatted report.
  Do NOT use when the user asks for interactive charts/dashboards (e.g. Excel/HTML
  dashboards), deep-dive analysis without a formatted document, or daily/quarterly
  summaries.
---

# Monthly Report Generator

Transform raw monthly sales data into a polished, executive-ready 3-page A4 Thai sales report in both **Word (`.docx`)** and **PDF (`.pdf`)** formats with the **Anuphan** font embedded.

## Bundled Resources

- **Templates**: [Sales Report - Formal.html](./assets/Sales%20Report%20-%20Formal.html) and [Sales Report - Modern.html](./assets/Sales%20Report%20-%20Modern.html)
- **Fonts**: `Anuphan` family in [assets/fonts/](./assets/fonts/) (`Regular`, `Medium`, `SemiBold`, `Bold`, `Light`, `ExtraLight`, `Thin`)
- **Sample Data**: [sample_raw_data.txt](./assets/sample_raw_data.txt)
- **Generator Script**: [scripts/generate_report.py](./scripts/generate_report.py)
- **Validator Script**: [scripts/validate_docx.py](./scripts/validate_docx.py)
- **Detailed Spec**: [references/report_spec.md](./references/report_spec.md)

---

## Workflow

### Step 1: Read Raw Data & Confirm Missing Details

1. Read the user's raw data file (sales, customers, top products, sales team, issues, next month plan).
2. Ask the user before generating:
   - **Template style**: `"Formal"` (solid navy section bars, bordered tables with zebra stripes, framed cover — formal executive style) or `"Modern"` (clean minimal borders, open whitespace — presentation style).
   - **Missing metadata** (only if absent from the raw data): Month/Year, Preparer/Author name, Company/Department name, or Report Title.

### Step 2: Parse, Compute & Verify Numbers

Extract all metrics from the raw data and compute/verify derived values:
- **% vs Target**: `(Actual - Target) / Target * 100`
- **MoM Growth (%)**: `(Current - Previous) / Previous * 100`
- **Product & Salesperson Shares (%)**: `Item Sales / Total Sales * 100`
- **Totals**: Sum quantities and sales across products and salespeople and verify they reconcile with Total Sales.

Prepare a JSON payload matching the schema in [references/report_spec.md](./references/report_spec.md).

### Step 3: Generate `.docx` and `.pdf`

Run [scripts/generate_report.py](./scripts/generate_report.py):

```powershell
uv run --offline --with python-docx --with lxml python <skill_dir>/scripts/generate_report.py --input <data.json> --template <formal|modern> --out-dir <output_dir> --name <base_filename>
```

What this script enforces automatically:
1. **Design Fidelity**: Populates the bundled HTML template (`Formal` or `Modern`, `#1F3864` navy palette) and generates a matching 3-page A4 `.docx`.
2. **Embedded Anuphan Font**:
   - In `.docx`: Embeds `Anuphan` TrueType fonts (`Regular`, `Bold`) inside `word/fonts/` via OOXML `fontTable.xml` + `embedTrueTypeFonts` in `word/settings.xml` so the document is self-contained and renders without tofu boxes (`□`) in Word and previewers.
   - In `.pdf`: Renders the populated HTML via headless Chromium/Edge (`--print-to-pdf`) using local `@font-face` Anuphan `.ttf` files so glyphs are embedded directly in the PDF.
3. **Thai Proofing & Spell-check**:
   - Sets `<w:lang w:val="en-US" w:bidi="th-TH"/>` on runs/styles (`w:bidi` slot for Thai complex script).
   - Adds `<w:noProof/>` **only** on runs containing Thai characters (`\u0E00-\u0E7F`), preventing red squiggly underlines on Thai text while keeping English spell-check active.

### Step 4: Validate & Visual Inspection

Run [scripts/validate_docx.py](./scripts/validate_docx.py) on the generated `.docx` and `.pdf`:

```powershell
uv run --offline --with lxml --with pymupdf python <skill_dir>/scripts/validate_docx.py --docx <output.docx> --pdf <output.pdf> --render-dir <preview_dir>
```

Verify all checks pass:
1. **DOCX XML & Font Validation**: `word/fontTable.xml` embeds `Anuphan`, `word/settings.xml` has `embedTrueTypeFonts`, Thai runs have `w:bidi="th-TH"` and `w:noProof`, English-only runs do **not** have `w:noProof`.
2. **Page Count**: Confirm PDF has **<= 3 pages**.
3. **Visual Inspection**: Read the rendered PNG images of every page (`page_1.png`, `page_2.png`, `page_3.png`) using `read_file` to confirm layout, alignment, no overlapping text, and no Thai tofu boxes.
4. **Number Reconciliation**: Cross-check all figures in the output against the raw data file.
