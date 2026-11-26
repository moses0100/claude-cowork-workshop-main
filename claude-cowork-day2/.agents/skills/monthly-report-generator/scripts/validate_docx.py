import argparse
import re
import zipfile
from pathlib import Path

import pymupdf
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
THAI_RE = re.compile(r"[\u0E00-\u0E7F]")


def validate_docx(docx_path: Path) -> list[str]:
    errors = []
    with zipfile.ZipFile(docx_path, "r") as z:
        names = set(z.namelist())
        if "word/fonts/font1.ttf" not in names or "word/fonts/font2.ttf" not in names:
            errors.append("Missing embedded Anuphan font binaries in word/fonts/")

        settings = etree.fromstring(z.read("word/settings.xml"))
        if settings.find(f"{{{W_NS}}}embedTrueTypeFonts") is None:
            errors.append("Missing <w:embedTrueTypeFonts/> in word/settings.xml")

        font_table = etree.fromstring(z.read("word/fontTable.xml"))
        anuphan = [f for f in font_table.findall(f"{{{W_NS}}}font") if f.get(f"{{{W_NS}}}name") == "Anuphan"]
        if not anuphan:
            errors.append("Anuphan font missing from word/fontTable.xml")
        else:
            if anuphan[0].find(f"{{{W_NS}}}embedRegular") is None:
                errors.append("Anuphan embedRegular missing in word/fontTable.xml")

        doc = etree.fromstring(z.read("word/document.xml"))
        thai_runs = 0
        thai_missing_bidi = 0
        thai_missing_noproof = 0
        eng_with_noproof = 0

        for r in doc.findall(f".//{{{W_NS}}}r"):
            text = "".join(t.text or "" for t in r.findall(f".//{{{W_NS}}}t"))
            if not text.strip():
                continue
            rPr = r.find(f"{{{W_NS}}}rPr")
            has_thai = bool(THAI_RE.search(text))
            has_noproof = rPr is not None and rPr.find(f"{{{W_NS}}}noProof") is not None
            lang = rPr.find(f"{{{W_NS}}}lang") if rPr is not None else None
            bidi_lang = lang.get(f"{{{W_NS}}}bidi") if lang is not None else None

            if has_thai:
                thai_runs += 1
                if bidi_lang != "th-TH":
                    thai_missing_bidi += 1
                if not has_noproof:
                    thai_missing_noproof += 1
            else:
                if has_noproof and re.search(r"[A-Za-z]", text):
                    eng_with_noproof += 1

        if thai_runs == 0:
            errors.append("No Thai runs found in document.xml")
        if thai_missing_bidi:
            errors.append(f"{thai_missing_bidi} Thai runs missing w:bidi='th-TH'")
        if thai_missing_noproof:
            errors.append(f"{thai_missing_noproof} Thai runs missing w:noProof")
        if eng_with_noproof:
            errors.append(f"{eng_with_noproof} English runs unexpectedly have w:noProof")

    return errors


def validate_and_render_pdf(pdf_path: Path, render_dir: Path) -> tuple[int, list[Path], list[str]]:
    render_dir.mkdir(parents=True, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    page_count = len(doc)
    images = []
    fonts_found = set()
    for i, page in enumerate(doc, start=1):
        for f in page.get_fonts():
            fonts_found.add(f[3])
        pix = page.get_pixmap(dpi=150)
        img_path = render_dir / f"page_{i}.png"
        pix.save(img_path)
        images.append(img_path)
    doc.close()
    return page_count, images, sorted(fonts_found)


def main():
    parser = argparse.ArgumentParser(description="Validate generated .docx and .pdf reports")
    parser.add_argument("--docx", required=True)
    parser.add_argument("--pdf", required=True)
    parser.add_argument("--render-dir", required=True)
    args = parser.parse_args()

    docx_errors = validate_docx(Path(args.docx))
    page_count, images, fonts = validate_and_render_pdf(Path(args.pdf), Path(args.render_dir))

    print(f"DOCX validation errors: {len(docx_errors)}")
    for e in docx_errors:
        print(f"  - {e}")
    print(f"PDF pages: {page_count}")
    print(f"PDF embedded fonts: {fonts}")
    print(f"Rendered images: {[str(p) for p in images]}")

    if docx_errors or page_count > 3:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
