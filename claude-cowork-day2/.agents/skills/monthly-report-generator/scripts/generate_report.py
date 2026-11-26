import argparse
import base64
import html
import json
import mimetypes
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor
from lxml import etree

SKILL_DIR = Path(__file__).resolve().parent.parent
ASSETS_DIR = SKILL_DIR / "assets"
FONTS_DIR = ASSETS_DIR / "fonts"

NAVY_HEX = "1F3864"
NAVY_RGB = RGBColor(0x1F, 0x38, 0x64)
INK_RGB = RGBColor(0x1B, 0x1F, 0x27)
MUTED_RGB = RGBColor(0x6B, 0x72, 0x80)
POS_RGB = RGBColor(0x1F, 0x7A, 0x4D)
WARN_RGB = RGBColor(0xB4, 0x53, 0x09)
WHITE_RGB = RGBColor(0xFF, 0xFF, 0xFF)

THAI_RE = re.compile(r"[\u0E00-\u0E7F]")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def has_thai(text: str) -> bool:
    return bool(THAI_RE.search(text or ""))


def fmt_int(n: float) -> str:
    return f"{int(round(n)):,}"


def fmt_pct(val: float, signed: bool = True) -> str:
    if signed:
        return f"{val:+.1f}%"
    return f"{val:.1f}%"


def strip_html_tags(text: str) -> str:
    return re.sub(r"<[^>]+>", "", text or "")


def add_run(p, text: str, size_pt: float = 10.0, bold: bool = False, color: RGBColor = INK_RGB, font_name: str = "Anuphan"):
    run = p.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.color.rgb = color

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    rFonts.set(qn("w:cs"), font_name)
    rFonts.set(qn("w:eastAsia"), font_name)

    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rPr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:bidi"), "th-TH")

    if has_thai(text):
        if rPr.find(qn("w:noProof")) is None:
            rPr.append(OxmlElement("w:noProof"))
    return run


def add_mixed_runs(p, html_text: str, size_pt: float = 10.0):
    """Split text into Thai/non-Thai segments and honor <b> and <span class='pos'> tags."""
    tokens = re.split(r"(<b>.*?</b>|<span[^>]*>.*?</span>)", html_text)
    for tok in tokens:
        if not tok:
            continue
        bold = False
        color = INK_RGB
        text = tok
        if tok.startswith("<b>") and tok.endswith("</b>"):
            bold = True
            text = tok[3:-4]
        elif tok.startswith("<span"):
            if "pos" in tok:
                color = POS_RGB
                bold = True
            text = strip_html_tags(tok)
        # Split Thai vs English/number segments so noProof only applies to Thai runs
        parts = re.findall(r"[\u0E00-\u0E7F\s\d.,()%+\-/–—:]+|[^\u0E00-\u0E7F]+", text)
        for part in parts:
            if not part:
                continue
            # Refine split so pure ASCII English words get their own run without noProof
            subparts = re.findall(r"[A-Za-z][A-Za-z0-9\s./+\-]*|[^A-Za-z]+", part)
            for sp in subparts:
                if sp:
                    add_run(p, sp, size_pt=size_pt, bold=bold, color=color)


def set_cell_shading(cell, hex_color: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for side, spec in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{side}")
        if spec is None:
            el.set(qn("w:val"), "none")
        else:
            val, sz, color = spec
            el.set(qn("w:val"), val)
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), color)
        tcBorders.append(el)
    tcPr.append(tcBorders)


def set_table_borders_none(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="none"/><w:left w:val="none"/><w:bottom w:val="none"/>'
        '<w:right w:val="none"/><w:insideH w:val="none"/><w:insideV w:val="none"/>'
        "</w:tblBorders>"
    )
    tblPr.append(borders)


def add_section_heading(doc_or_cell, no: str, title: str, en: str, variant: str):
    if variant == "formal":
        tbl = doc_or_cell.add_table(rows=1, cols=2)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders_none(tbl)
        c0, c1 = tbl.rows[0].cells[0], tbl.rows[0].cells[1]
        c0.width = Cm(12.5)
        c1.width = Cm(5.1)
        set_cell_shading(c0, NAVY_HEX)
        set_cell_shading(c1, NAVY_HEX)
        set_cell_margins(c0, top=90, bottom=90, left=140, right=80)
        set_cell_margins(c1, top=90, bottom=90, left=80, right=140)
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        add_run(p0, f"{no}  ", size_pt=10.5, bold=True, color=WHITE_RGB)
        add_run(p0, title, size_pt=12.0, bold=True, color=WHITE_RGB)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        add_run(p1, en.upper(), size_pt=8.0, bold=False, color=RGBColor(0xD0, 0xD8, 0xE8))
    else:
        p = doc_or_cell.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="4" w:color="1B1F27"/></w:pBdr>')
        p._p.get_or_add_pPr().append(pBdr)
        add_run(p, f"{no}  ", size_pt=10.5, bold=True, color=NAVY_RGB)
        add_run(p, title, size_pt=13.0, bold=True, color=INK_RGB)
        add_run(p, f"   [{en.upper()}]", size_pt=8.0, bold=False, color=MUTED_RGB)


def style_data_table(tbl, headers, rows_data, alignments, variant: str, total_row=None, col_widths=None):
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders_none(tbl)
    border_line = ("single", 4, "D5D9E0")
    navy_line = ("single", 10, NAVY_HEX)

    # Header
    hdr_cells = tbl.rows[0].cells
    for idx, h in enumerate(headers):
        cell = hdr_cells[idx]
        if col_widths:
            cell.width = Cm(col_widths[idx])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        if variant == "formal":
            set_cell_shading(cell, NAVY_HEX)
            set_cell_borders(cell, top=navy_line, bottom=navy_line, left=border_line, right=border_line)
            txt_color = WHITE_RGB
        else:
            set_cell_borders(cell, bottom=navy_line)
            txt_color = NAVY_RGB
        p = cell.paragraphs[0]
        p.alignment = alignments[idx]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        add_run(p, h, size_pt=9.0, bold=True, color=txt_color)

    # Body
    for r_idx, r_vals in enumerate(rows_data):
        row_cells = tbl.add_row().cells
        for c_idx, val in enumerate(r_vals):
            cell = row_cells[c_idx]
            if col_widths:
                cell.width = Cm(col_widths[c_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
            if variant == "formal":
                if r_idx % 2 == 1:
                    set_cell_shading(cell, "F7F9FC")
                set_cell_borders(cell, bottom=border_line, left=border_line, right=border_line)
            else:
                set_cell_borders(cell, bottom=("single", 4, "E7E9EE"))
            p = cell.paragraphs[0]
            p.alignment = alignments[c_idx]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            if isinstance(val, tuple):
                txt, clr, is_b = val
                add_run(p, txt, size_pt=9.5, bold=is_b, color=clr)
            else:
                add_run(p, str(val), size_pt=9.5, bold=False, color=INK_RGB)

    # Total row
    if total_row:
        row_cells = tbl.add_row().cells
        for c_idx, val in enumerate(total_row):
            cell = row_cells[c_idx]
            if col_widths:
                cell.width = Cm(col_widths[c_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            if variant == "formal":
                set_cell_shading(cell, "EEF2F8")
                set_cell_borders(cell, top=navy_line, bottom=navy_line, left=border_line, right=border_line)
            else:
                set_cell_borders(cell, top=navy_line)
            p = cell.paragraphs[0]
            p.alignment = alignments[c_idx]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            add_run(p, str(val), size_pt=9.5, bold=True, color=NAVY_RGB)


def embed_anuphan_fonts_in_docx(docx_path: Path):
    """Embed Anuphan-Regular.ttf and Anuphan-Bold.ttf into .docx OOXML package."""
    reg_ttf = (FONTS_DIR / "Anuphan-Regular.ttf").read_bytes()
    bold_ttf = (FONTS_DIR / "Anuphan-Bold.ttf").read_bytes()

    tmp_zip = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as zout:
        existing_names = set(zin.namelist())
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                root = etree.fromstring(data)
                ns = root.nsmap[None]
                has_ttf = any(
                    el.get("Extension") == "ttf" for el in root.findall(f"{{{ns}}}Default")
                )
                if not has_ttf:
                    el = etree.SubElement(root, f"{{{ns}}}Default")
                    el.set("Extension", "ttf")
                    el.set("ContentType", "application/x-font-ttf")
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif item.filename == "word/settings.xml":
                root = etree.fromstring(data)
                if root.find(f"{{{W_NS}}}embedTrueTypeFonts") is None:
                    el = etree.SubElement(root, f"{{{W_NS}}}embedTrueTypeFonts")
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif item.filename == "word/fontTable.xml":
                root = etree.fromstring(data)
                # Remove existing Anuphan entry if present
                for fnode in list(root.findall(f"{{{W_NS}}}font")):
                    if fnode.get(f"{{{W_NS}}}name") == "Anuphan":
                        root.remove(fnode)
                fnode = etree.SubElement(root, f"{{{W_NS}}}font")
                fnode.set(f"{{{W_NS}}}name", "Anuphan")
                charset = etree.SubElement(fnode, f"{{{W_NS}}}charset")
                charset.set(f"{{{W_NS}}}val", "DE")
                family = etree.SubElement(fnode, f"{{{W_NS}}}family")
                family.set(f"{{{W_NS}}}val", "swiss")
                pitch = etree.SubElement(fnode, f"{{{W_NS}}}pitch")
                pitch.set(f"{{{W_NS}}}val", "variable")
                er = etree.SubElement(fnode, f"{{{W_NS}}}embedRegular")
                er.set(f"{{{R_NS}}}id", "rIdFont1")
                eb = etree.SubElement(fnode, f"{{{W_NS}}}embedBold")
                eb.set(f"{{{R_NS}}}id", "rIdFont2")
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            zout.writestr(item, data)

        # Write embedded font binaries
        zout.writestr("word/fonts/font1.ttf", reg_ttf)
        zout.writestr("word/fonts/font2.ttf", bold_ttf)

        # Write word/_rels/fontTable.xml.rels
        rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rIdFont1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font1.ttf"/>'
            '<Relationship Id="rIdFont2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="fonts/font2.ttf"/>'
            '</Relationships>'
        ).encode("utf-8")
        if "word/_rels/fontTable.xml.rels" not in existing_names:
            zout.writestr("word/_rels/fontTable.xml.rels", rels_xml)

    shutil.move(str(tmp_zip), str(docx_path))


def build_docx(data: dict, variant: str, out_path: Path):
    doc = Document()

    # A4 page setup (21.0 x 29.7 cm), 1.6 cm margins
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.4)
    section.left_margin = Cm(1.6)
    section.right_margin = Cm(1.6)

    # Set Normal style default font + Thai bidi language
    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Anuphan"
    normal_style.font.size = Pt(10.0)
    rPr = normal_style.element.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), "Anuphan")
    rPr.append(rFonts)
    lang = OxmlElement("w:lang")
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:bidi"), "th-TH")
    rPr.append(lang)

    kpis = data["kpis"]
    total_sales = kpis["total_sales"]
    target_sales = kpis["target_sales"]
    prev_sales = kpis["prev_sales"]
    vs_target_pct = (total_sales - target_sales) / target_sales * 100
    mom_pct = (total_sales - prev_sales) / prev_sales * 100

    # ================= PAGE 1: COVER =================
    cover_box = doc.add_table(rows=1, cols=1)
    cover_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders_none(cover_box)
    ccell = cover_box.rows[0].cells[0]
    ccell.width = Cm(17.8)
    if variant == "formal":
        navy_border = ("single", 12, NAVY_HEX)
        set_cell_borders(ccell, top=navy_border, bottom=navy_border, left=navy_border, right=navy_border)
        set_cell_margins(ccell, top=420, bottom=420, left=380, right=380)
    else:
        set_cell_margins(ccell, top=300, bottom=300, left=180, right=180)

    # Masthead inside cover
    logo_path = Path(data["logo"]) if data.get("logo") else None
    has_logo = logo_path is not None and logo_path.exists()
    if has_logo:
        m_tbl = ccell.add_table(rows=1, cols=3)
        set_table_borders_none(m_tbl)
        mc_logo, mc0, mc1 = m_tbl.rows[0].cells[0], m_tbl.rows[0].cells[1], m_tbl.rows[0].cells[2]
        mc_logo.width = Cm(1.8)
        mc0.width = Cm(9.2)
        mc1.width = Cm(5.4)
        mc_logo.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        mc0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        mc1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p_logo = mc_logo.paragraphs[0]
        p_logo.paragraph_format.space_before = Pt(0)
        p_logo.paragraph_format.space_after = Pt(0)
        p_logo.add_run().add_picture(str(logo_path), width=Cm(1.4))
    else:
        m_tbl = ccell.add_table(rows=1, cols=2)
        set_table_borders_none(m_tbl)
        mc0, mc1 = m_tbl.rows[0].cells[0], m_tbl.rows[0].cells[1]
        mc0.width = Cm(11.0)
        mc1.width = Cm(5.4)
    p_co = mc0.paragraphs[0]
    add_run(p_co, data["company"] + "\n", size_pt=11.5, bold=True, color=INK_RGB)
    add_run(p_co, data["department"], size_pt=8.5, bold=False, color=MUTED_RGB)
    p_tag = mc1.paragraphs[0]
    p_tag.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run(p_tag, data.get("doc_tag_cover", "CONFIDENTIAL"), size_pt=8.0, bold=False, color=MUTED_RGB)

    for _ in range(4):
        sp = ccell.add_paragraph()
        sp.paragraph_format.space_before = Pt(10)
        sp.paragraph_format.space_after = Pt(10)

    cover_align = WD_ALIGN_PARAGRAPH.CENTER if variant == "formal" else WD_ALIGN_PARAGRAPH.LEFT
    p_eye = ccell.add_paragraph()
    p_eye.alignment = cover_align
    add_run(p_eye, "รายงานประจำเดือน · MONTHLY SALES REPORT", size_pt=9.0, bold=True, color=NAVY_RGB)

    p_title = ccell.add_paragraph()
    p_title.alignment = cover_align
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(4)
    add_run(p_title, data["title"], size_pt=28.0, bold=True, color=INK_RGB)

    p_mon = ccell.add_paragraph()
    p_mon.alignment = cover_align
    p_mon.paragraph_format.space_after = Pt(18)
    add_run(p_mon, data["month_label"], size_pt=16.0, bold=False, color=NAVY_RGB)

    # Cover metadata 2x2
    meta_tbl = ccell.add_table(rows=2, cols=2)
    meta_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER if variant == "formal" else WD_TABLE_ALIGNMENT.LEFT
    set_table_borders_none(meta_tbl)
    meta_items = [
        ("รอบรายงาน", data["period"]),
        ("วันที่จัดทำ", data["prepared_date"]),
        ("จัดทำโดย", data["prepared_by"]),
        ("ผู้อนุมัติ", data["approved_by"]),
    ]
    for idx, (k, v) in enumerate(meta_items):
        cell = meta_tbl.rows[idx // 2].cells[idx % 2]
        cell.width = Cm(7.0)
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        pm = cell.paragraphs[0]
        pm.alignment = cover_align
        add_run(pm, k + "\n", size_pt=8.0, bold=False, color=MUTED_RGB)
        add_run(pm, v, size_pt=10.5, bold=True, color=INK_RGB)

    for _ in range(3):
        sp = ccell.add_paragraph()
        sp.paragraph_format.space_before = Pt(10)
        sp.paragraph_format.space_after = Pt(10)

    p_f1 = doc.add_paragraph()
    p_f1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_f1.paragraph_format.space_before = Pt(8)
    add_run(p_f1, f"{data['company']} — รายงานภายใน   |   หน้า 1 / 3", size_pt=8.0, color=MUTED_RGB)

    doc.add_page_break()

    # ================= PAGE 2: EXEC SUMMARY + PERFORMANCE =================
    add_section_heading(doc, "01", "บทสรุปสำหรับผู้บริหาร", "Executive Summary", variant)

    kpi_tbl = doc.add_table(rows=1, cols=4)
    kpi_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders_none(kpi_tbl)
    kpi_cards = [
        (fmt_int(total_sales), "ยอดขายรวม (บาท)", INK_RGB),
        (fmt_pct(vs_target_pct), "เทียบกับเป้าหมาย", POS_RGB if vs_target_pct >= 0 else WARN_RGB),
        (fmt_pct(mom_pct), "เติบโตจากเดือนก่อน", POS_RGB if mom_pct >= 0 else WARN_RGB),
        (f"{kpis['csat']:.1f} / 5.0", "ความพึงพอใจลูกค้า", INK_RGB),
    ]
    for idx, (val_txt, lbl_txt, clr) in enumerate(kpi_cards):
        cell = kpi_tbl.rows[0].cells[idx]
        cell.width = Cm(4.4)
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        if variant == "formal":
            set_cell_shading(cell, "F4F6FA")
            set_cell_borders(cell, top=("single", 4, "D5D9E0"), bottom=("single", 4, "D5D9E0"), left=("single", 4, "D5D9E0"), right=("single", 4, "D5D9E0"))
        elif idx > 0:
            set_cell_borders(cell, left=("single", 4, "E7E9EE"))
        pk = cell.paragraphs[0]
        pk.paragraph_format.space_before = Pt(0)
        pk.paragraph_format.space_after = Pt(0)
        add_run(pk, val_txt + "\n", size_pt=15.0, bold=True, color=clr)
        add_run(pk, lbl_txt, size_pt=8.5, bold=False, color=MUTED_RGB)

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(3)
    sp.paragraph_format.space_after = Pt(3)

    for hl in data["highlights"]:
        ph = doc.add_paragraph()
        ph.paragraph_format.left_indent = Cm(0.45)
        ph.paragraph_format.space_before = Pt(2)
        ph.paragraph_format.space_after = Pt(2)
        add_run(ph, "•  ", size_pt=10.0, bold=True, color=NAVY_RGB)
        add_mixed_runs(ph, hl, size_pt=10.0)

    sp2 = doc.add_paragraph()
    sp2.paragraph_format.space_before = Pt(6)
    sp2.paragraph_format.space_after = Pt(4)

    add_section_heading(doc, "02", "ผลการดำเนินงาน", "Performance · Target vs Actual", variant)

    perf_tbl = doc.add_table(rows=1, cols=4)
    perf_rows = [
        ["ยอดขายรวม (บาท)", fmt_int(target_sales), fmt_int(total_sales), (fmt_pct(vs_target_pct), POS_RGB if vs_target_pct >= 0 else WARN_RGB, True)],
        ["เทียบเดือนก่อน (บาท)", fmt_int(prev_sales), fmt_int(total_sales), (fmt_pct(mom_pct), POS_RGB if mom_pct >= 0 else WARN_RGB, True)],
        ["ลูกค้าใหม่ (ราย)", "—", fmt_int(kpis["new_customers"]), ("เพิ่มขึ้น", POS_RGB, True)],
        ["ลูกค้า Active (ราย)", "—", fmt_int(kpis["active_customers"]), ("คงที่", MUTED_RGB, False)],
        ["ลูกค้ายกเลิก / Churn", "—", fmt_int(kpis["churn_customers"]), ("เฝ้าระวัง", WARN_RGB, True)],
        ["ความพึงพอใจ (เต็ม 5)", "—", f"{kpis['csat']:.1f}", ("ดี", POS_RGB, True)],
    ]
    style_data_table(
        perf_tbl,
        ["ตัวชี้วัด", "เป้าหมาย / ฐาน", "ผลจริง", "สถานะ"],
        perf_rows,
        [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER],
        variant,
        col_widths=[6.4, 3.8, 3.8, 3.6],
    )

    p_f2 = doc.add_paragraph()
    p_f2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_f2.paragraph_format.space_before = Pt(14)
    add_run(p_f2, f"{data['company']} — รายงานภายใน   |   หน้า 2 / 3", size_pt=8.0, color=MUTED_RGB)

    doc.add_page_break()

    # ================= PAGE 3: PRODUCTS, TEAM, ISSUES, PLAN =================
    two_col = doc.add_table(rows=1, cols=2)
    two_col.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders_none(two_col)
    lc, rc = two_col.rows[0].cells[0], two_col.rows[0].cells[1]
    lc.width = Cm(8.7)
    rc.width = Cm(8.7)
    set_cell_margins(lc, top=0, bottom=0, left=0, right=100)
    set_cell_margins(rc, top=0, bottom=0, left=100, right=0)

    # Left: Top 4 Products
    add_section_heading(lc, "03", "สินค้าขายดี", "Top 4", variant)
    prod_tbl = lc.add_table(rows=1, cols=4)
    prods = data["top_products"]
    prod_rows = [[str(i + 1), p["name"], fmt_int(p["qty"]), fmt_int(p["sales"])] for i, p in enumerate(prods)]
    tot_qty = sum(p["qty"] for p in prods)
    tot_prod_sales = sum(p["sales"] for p in prods)
    style_data_table(
        prod_tbl,
        ["#", "สินค้า", "จำนวน", "ยอดขาย"],
        prod_rows,
        [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT],
        variant,
        total_row=["", "รวมทั้งสิ้น", fmt_int(tot_qty), fmt_int(tot_prod_sales)],
        col_widths=[0.9, 3.3, 1.8, 2.5],
    )

    # Right: Sales Team
    add_section_heading(rc, "04", "ผลงานทีมขาย", "Sales Team", variant)
    team_tbl = rc.add_table(rows=1, cols=4)
    team = data["sales_team"]
    tot_team_sales = sum(m["sales"] for m in team)
    team_rows = []
    for i, m in enumerate(team):
        name_lbl = f"{m['name']} (TOP)" if m.get("top") else m["name"]
        share = m["sales"] / tot_team_sales * 100 if tot_team_sales else 0
        team_rows.append([str(i + 1), name_lbl, fmt_int(m["sales"]), f"{share:.1f}%"])
    style_data_table(
        team_tbl,
        ["#", "พนักงานขาย", "ยอดขาย", "สัดส่วน"],
        team_rows,
        [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT],
        variant,
        total_row=["", "รวมทั้งทีม", fmt_int(tot_team_sales), "100%"],
        col_widths=[0.9, 3.3, 2.5, 1.8],
    )

    sp3 = doc.add_paragraph()
    sp3.paragraph_format.space_before = Pt(6)
    sp3.paragraph_format.space_after = Pt(2)

    # Section 05: Issues & Actions
    add_section_heading(doc, "05", "ปัญหาและแนวทางแก้ไข", "Issues & Actions", variant)
    for item in data["issues"]:
        pi = doc.add_paragraph()
        pi.paragraph_format.left_indent = Cm(0.4)
        pi.paragraph_format.space_before = Pt(3)
        pi.paragraph_format.space_after = Pt(3)
        add_run(pi, "▪  ", size_pt=10.0, bold=True, color=WARN_RGB)
        add_run(pi, item["issue"] + "\n", size_pt=9.8, bold=True, color=INK_RGB)
        add_run(pi, "แนวทาง: ", size_pt=9.5, bold=True, color=NAVY_RGB)
        add_run(pi, item["action"], size_pt=9.5, bold=False, color=MUTED_RGB)

    sp4 = doc.add_paragraph()
    sp4.paragraph_format.space_before = Pt(4)
    sp4.paragraph_format.space_after = Pt(2)

    # Section 06: Next Month Plan
    add_section_heading(doc, "06", data["next_month_title"], "Next Month Plan", variant)
    for item in data["next_month_plan"]:
        pc = doc.add_paragraph()
        pc.paragraph_format.left_indent = Cm(0.4)
        pc.paragraph_format.space_before = Pt(2)
        pc.paragraph_format.space_after = Pt(2)
        add_run(pc, "☐  ", size_pt=10.5, bold=True, color=NAVY_RGB)
        add_run(pc, item["task"], size_pt=9.8, bold=True, color=INK_RGB)
        add_run(pc, f"  — {item['owner']}", size_pt=9.0, bold=False, color=MUTED_RGB)

    p_f3 = doc.add_paragraph()
    p_f3.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_f3.paragraph_format.space_before = Pt(10)
    add_run(p_f3, f"{data['company']} — รายงานภายใน   |   หน้า 3 / 3", size_pt=8.0, color=MUTED_RGB)

    doc.save(out_path)
    embed_anuphan_fonts_in_docx(out_path)


def build_html(data: dict, variant: str, out_html: Path):
    tpl_name = "Sales Report - Formal.html" if variant == "formal" else "Sales Report - Modern.html"
    tpl_text = (ASSETS_DIR / tpl_name).read_text(encoding="utf-8")

    # Point @font-face urls to absolute file:/// paths so HTML works from any output directory
    fonts_uri = FONTS_DIR.resolve().as_uri()
    tpl_text = tpl_text.replace('url("fonts/', f'url("{fonts_uri}/')

    kpis = data["kpis"]
    total_sales = kpis["total_sales"]
    target_sales = kpis["target_sales"]
    prev_sales = kpis["prev_sales"]
    vs_target_pct = (total_sales - target_sales) / target_sales * 100
    mom_pct = (total_sales - prev_sales) / prev_sales * 100
    mom_diff = total_sales - prev_sales

    max_val = max(total_sales, target_sales, prev_sales, 1)
    h_prev = prev_sales / max_val * 100
    h_target = target_sales / max_val * 100
    h_actual = total_sales / max_val * 100

    hl_html = "\n".join(f"        <li>{h}</li>" for h in data["highlights"])

    prods = data["top_products"]
    tot_qty = sum(p["qty"] for p in prods)
    tot_prod_sales = sum(p["sales"] for p in prods)
    prod_rows_html = "\n".join(
        f'            <tr><td class="ctr">{i+1}</td><td>{html.escape(p["name"])}</td><td class="num">{fmt_int(p["qty"])}</td><td class="num">{fmt_int(p["sales"])}</td></tr>'
        for i, p in enumerate(prods)
    )

    team = data["sales_team"]
    tot_team_sales = sum(m["sales"] for m in team)
    team_rows_html = "\n".join(
        f'            <tr><td class="ctr">{i+1}</td><td>{html.escape(m["name"])}{"<span class=\'star\'>TOP</span>" if m.get("top") else ""}</td><td class="num">{fmt_int(m["sales"])}</td><td class="num">{(m["sales"]/tot_team_sales*100 if tot_team_sales else 0):.1f}%</td></tr>'
        for i, m in enumerate(team)
    )

    issues_html = "\n".join(
        f'        <li><div class="p">{html.escape(it["issue"])}</div><div class="s"><b>แนวทาง:</b> {html.escape(it["action"])}</div></li>'
        for it in data["issues"]
    )
    plan_html = "\n".join(
        f'        <li><span class="box"></span><span class="ctxt"><b>{html.escape(it["task"])}</b> &nbsp;<span class="owner">— {html.escape(it["owner"])}</span></span></li>'
        for it in data["next_month_plan"]
    )

    logo_cover_html = '<div class="logo-slot">LOGO</div>'
    logo_hdr_html = '<div class="logo-slot" style="width:34px;height:34px;font-size:6pt;">LOGO</div>'
    if data.get("logo"):
        lp = Path(data["logo"])
        if lp.exists():
            mime, _ = mimetypes.guess_type(str(lp))
            mime = mime or "image/png"
            b64 = base64.b64encode(lp.read_bytes()).decode("ascii")
            uri = f"data:{mime};base64,{b64}"
            logo_cover_html = f'<img src="{uri}" alt="Logo" style="width:48px;height:48px;object-fit:contain;">'
            logo_hdr_html = f'<img src="{uri}" alt="Logo" style="width:34px;height:34px;object-fit:contain;">'

    body_html = f"""<body data-variant="{variant}">

<div class="deck">

  <!-- ============ PAGE 1 — COVER ============ -->
  <section class="page" data-screen-label="หน้า 1 · ปก">
    <div class="cover">
      <div class="cover-inner">
        <div class="masthead">
          <div class="brand">
            {logo_cover_html}
            <div class="co"><b>{html.escape(data['company'])}</b><span>{html.escape(data['department'])}</span></div>
          </div>
          <div class="doc-tag">{html.escape(data.get('doc_tag_cover', 'CONFIDENTIAL')).replace(chr(10), '<br>')}</div>
        </div>

        <div class="cover-mid">
          <div class="eyebrow">รายงานประจำเดือน · Monthly Sales Report</div>
          <h1>{html.escape(data['title'])}</h1>
          <div class="month">{html.escape(data['month_label'])}</div>
          <div class="rule"></div>
          <div class="cover-meta">
            <div><div class="k">รอบรายงาน</div><div class="v">{html.escape(data['period'])}</div></div>
            <div><div class="k">วันที่จัดทำ</div><div class="v">{html.escape(data['prepared_date'])}</div></div>
            <div><div class="k">จัดทำโดย</div><div class="v">{html.escape(data['prepared_by'])}</div></div>
            <div><div class="k">ผู้อนุมัติ</div><div class="v">{html.escape(data['approved_by'])}</div></div>
          </div>
        </div>
      </div>
    </div>
    <div class="page-footer"><span>{html.escape(data['company'])} — รายงานภายใน</span><span class="mono">หน้า 1 / 3</span></div>
  </section>

  <!-- ============ PAGE 2 ============ -->
  <section class="page" data-screen-label="หน้า 2 · บทสรุป + ผลการดำเนินงาน">
    <div class="masthead" style="margin-bottom:2mm;">
      <div class="brand">{logo_hdr_html}<div class="co"><b style="font-size:10pt;">{html.escape(data['company'])}</b><span style="font-size:7.6pt;">{html.escape(data['title'])} · {html.escape(data['short_month'])}</span></div></div>
      <div class="doc-tag" style="font-size:7pt;">{html.escape(data.get('doc_tag_header', 'SALES REPORT')).replace(chr(10), '<br>')}</div>
    </div>

    <div class="sec">
      <div class="sec-head"><span class="no">01</span><h2>บทสรุปสำหรับผู้บริหาร</h2><span class="en">Executive Summary</span></div>
      <div class="kpis">
        <div class="kpi"><div class="n">{fmt_int(total_sales)}</div><div class="l">ยอดขายรวม (บาท)</div></div>
        <div class="kpi"><div class="n {'pos' if vs_target_pct >= 0 else ''}">{fmt_pct(vs_target_pct)}</div><div class="l">เทียบกับเป้าหมาย</div></div>
        <div class="kpi"><div class="n {'pos' if mom_pct >= 0 else ''}">{fmt_pct(mom_pct)}</div><div class="l">เติบโตจากเดือนก่อน</div></div>
        <div class="kpi"><div class="n">{kpis['csat']:.1f}<span style="font-size:11pt;color:var(--faint);"> / 5.0</span></div><div class="l">ความพึงพอใจลูกค้า</div></div>
      </div>
      <ul class="hl">
{hl_html}
      </ul>
    </div>

    <div class="sec">
      <div class="sec-head"><span class="no">02</span><h2>ผลการดำเนินงาน</h2><span class="en">Performance · Target vs Actual</span></div>
      <div class="row2">
        <div class="chart-card">
          <div class="ct">ยอดขาย เทียบเป้าหมาย</div>
          <div class="cs">หน่วย: ล้านบาท</div>
          <div class="plot">
            <div class="tline" style="bottom:{h_target:.1f}%;"><span class="tl-lab">เป้า {target_sales/1e6:.2f}</span></div>
            <div class="bcol"><div class="bar prev" style="height:{h_prev:.1f}%;"><span class="bv">{prev_sales/1e6:.2f}</span></div><div class="blab">{html.escape(kpis.get('prev_month_label', 'เดือนก่อน'))}</div></div>
            <div class="bcol"><div class="bar target" style="height:{h_target:.1f}%;"><span class="bv">{target_sales/1e6:.2f}</span></div><div class="blab">เป้าหมาย</div></div>
            <div class="bcol"><div class="bar actual" style="height:{h_actual:.1f}%;"><span class="bv">{total_sales/1e6:.2f}</span></div><div class="blab">{html.escape(kpis.get('actual_month_label', 'ผลจริง'))}</div></div>
          </div>
          <div class="chart-foot"><span class="badge">{'เกินเป้า' if vs_target_pct >= 0 else 'ต่ำกว่าเป้า'} {fmt_pct(vs_target_pct)}</span><span class="ng">{mom_diff:+,} บาท จากเดือนก่อน</span></div>
        </div>
        <table class="tbl">
          <thead><tr><th>ตัวชี้วัด</th><th class="num">เป้าหมาย</th><th class="num">ผลจริง</th><th class="ctr">สถานะ</th></tr></thead>
          <tbody>
            <tr><td>ยอดขายรวม (บาท)</td><td class="num">{fmt_int(target_sales)}</td><td class="num">{fmt_int(total_sales)}</td><td class="ctr"><span class="tag {'up' if vs_target_pct >= 0 else 'watch'}">{fmt_pct(vs_target_pct)}</span></td></tr>
            <tr><td>เทียบเดือนก่อน</td><td class="num">{fmt_int(prev_sales)}</td><td class="num">{fmt_int(total_sales)}</td><td class="ctr"><span class="tag {'up' if mom_pct >= 0 else 'watch'}">{fmt_pct(mom_pct)}</span></td></tr>
            <tr><td>ลูกค้าใหม่ (ราย)</td><td class="num">—</td><td class="num">{fmt_int(kpis['new_customers'])}</td><td class="ctr"><span class="tag up">เพิ่มขึ้น</span></td></tr>
            <tr><td>ลูกค้า Active (ราย)</td><td class="num">—</td><td class="num">{fmt_int(kpis['active_customers'])}</td><td class="ctr"><span class="tag flat">คงที่</span></td></tr>
            <tr><td>ลูกค้ายกเลิก / Churn</td><td class="num">—</td><td class="num">{fmt_int(kpis['churn_customers'])}</td><td class="ctr"><span class="tag watch">เฝ้าระวัง</span></td></tr>
            <tr><td>ความพึงพอใจ (เต็ม 5)</td><td class="num">—</td><td class="num">{kpis['csat']:.1f}</td><td class="ctr"><span class="tag up">ดี</span></td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="page-footer"><span>{html.escape(data['company'])} — รายงานภายใน</span><span class="mono">หน้า 2 / 3</span></div>
  </section>

  <!-- ============ PAGE 3 ============ -->
  <section class="page" data-screen-label="หน้า 3 · สินค้า · ทีมขาย · ปัญหา · แผน">
    <div class="masthead" style="margin-bottom:2mm;">
      <div class="brand">{logo_hdr_html}<div class="co"><b style="font-size:10pt;">{html.escape(data['company'])}</b><span style="font-size:7.6pt;">{html.escape(data['title'])} · {html.escape(data['short_month'])}</span></div></div>
      <div class="doc-tag" style="font-size:7pt;">{html.escape(data.get('doc_tag_header', 'SALES REPORT')).replace(chr(10), '<br>')}</div>
    </div>

    <div class="row2" style="grid-template-columns:1fr 1fr;gap:26px;">
      <div class="sec" style="margin-top:7mm;">
        <div class="sec-head"><span class="no">03</span><h2>สินค้าขายดี</h2><span class="en">Top 4</span></div>
        <table class="tbl">
          <thead><tr><th class="ctr">#</th><th>สินค้า</th><th class="num">จำนวน</th><th class="num">ยอดขาย</th></tr></thead>
          <tbody>
{prod_rows_html}
            <tr class="total"><td></td><td>รวมทั้งสิ้น</td><td class="num">{fmt_int(tot_qty)}</td><td class="num">{fmt_int(tot_prod_sales)}</td></tr>
          </tbody>
        </table>
      </div>
      <div class="sec" style="margin-top:7mm;">
        <div class="sec-head"><span class="no">04</span><h2>ผลงานทีมขาย</h2><span class="en">Sales Team</span></div>
        <table class="tbl">
          <thead><tr><th class="ctr">#</th><th>พนักงานขาย</th><th class="num">ยอดขาย</th><th class="num">สัดส่วน</th></tr></thead>
          <tbody>
{team_rows_html}
            <tr class="total"><td></td><td>รวมทั้งทีม</td><td class="num">{fmt_int(tot_team_sales)}</td><td class="num">100%</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <div class="sec">
      <div class="sec-head"><span class="no">05</span><h2>ปัญหาและแนวทางแก้ไข</h2><span class="en">Issues &amp; Actions</span></div>
      <ul class="prob">
{issues_html}
      </ul>
    </div>

    <div class="sec">
      <div class="sec-head"><span class="no">06</span><h2>{html.escape(data['next_month_title'])}</h2><span class="en">Next Month Plan</span></div>
      <ul class="check">
{plan_html}
      </ul>
    </div>

    <div class="page-footer"><span>{html.escape(data['company'])} — รายงานภายใน</span><span class="mono">หน้า 3 / 3</span></div>
  </section>

</div>
</body>"""

    head_part = tpl_text.split("<body")[0]
    title_updated = re.sub(r"<title>.*?</title>", f"<title>{html.escape(data['title'])} — {html.escape(data['short_month'])}</title>", head_part)
    out_html.write_text(title_updated + body_html + "\n</html>\n", encoding="utf-8")


def find_browser_exe() -> str:
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    raise FileNotFoundError("Neither Chrome nor Edge found for headless PDF rendering.")


def render_pdf_from_html(html_path: Path, pdf_path: Path):
    browser = find_browser_exe()
    with tempfile.TemporaryDirectory() as user_data:
        cmd = [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--no-sandbox",
            f"--user-data-dir={user_data}",
            "--no-pdf-header-footer",
            f"--print-to-pdf={pdf_path.resolve()}",
            html_path.resolve().as_uri(),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if res.returncode != 0 or not pdf_path.exists():
            raise RuntimeError(f"Headless PDF render failed ({res.returncode}): {res.stderr}")


def main():
    parser = argparse.ArgumentParser(description="Generate Thai Monthly Sales Report (.docx & .pdf)")
    parser.add_argument("--input", required=True, help="Path to structured JSON input")
    parser.add_argument("--template", choices=["formal", "modern"], required=True, help="Template style")
    parser.add_argument("--out-dir", required=True, help="Output directory")
    parser.add_argument("--name", required=True, help="Base filename without extension")
    args = parser.parse_args()

    data = json.loads(Path(args.input).read_text(encoding="utf-8-sig"))
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    docx_path = out_dir / f"{args.name}.docx"
    html_path = out_dir / f"{args.name}.html"
    pdf_path = out_dir / f"{args.name}.pdf"

    build_docx(data, args.template, docx_path)
    build_html(data, args.template, html_path)
    render_pdf_from_html(html_path, pdf_path)

    print(f"DOCX: {docx_path}")
    print(f"HTML: {html_path}")
    print(f"PDF : {pdf_path}")


if __name__ == "__main__":
    main()
