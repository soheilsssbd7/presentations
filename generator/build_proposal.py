# -*- coding: utf-8 -*-
"""
سازندهٔ پروپوزال آرون — تولید سند Word فارسی (RTL) با هویت بصری مرجع.

طراحی برای بازاستفاده: تمام متن‌ها از یک فایل JSON می‌آید و ساختار
(تعداد محورها، لایه‌ها و الگوهای جهانی) آزاد است. استایل دقیقاً از
پروپوزال مرجع آرون ⇄ دات‌وان استخراج شده است:

    صفحه: Letter، حاشیه ۰٫۸ اینچ
    متن: B Nazanin 12، فاصله خط ۱٫۲۵، هم‌تراز دوطرفه، RTL
    تیتر سند: B Titr 12، رنگ #1B263B، وسط‌چین
    تیتر بخش: B Titr 14 پررنگ، #1B263B
    تیتر محور: B Nazanin 12، #415A77
    تیتر لایه: B Titr 12 پررنگ، #1B263B
    متن اصلی: #323232
    نوار بالا: پس‌زمینه #F1F5F9 + خط عمودی ۲٫۵pt رنگ #1B263B
    نوار پایین: پس‌زمینه #1B263B + متن سفید

نکته: نشانهٔ **متن** در JSON به متن پررنگ تبدیل می‌شود.
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Twips

# ---------------------------------------------------------------- palette
NAVY = RGBColor(0x1B, 0x26, 0x3B)
SLATE = RGBColor(0x41, 0x5A, 0x77)
INK = RGBColor(0x32, 0x32, 0x32)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BAND_LIGHT = "F1F5F9"
BAND_DARK = "1B263B"

BODY_FONT = "B Nazanin"
TITLE_FONT = "B Titr"
CS_FONT = "B Zar"


# ---------------------------------------------------------------- helpers
def _set_run(run, *, font=BODY_FONT, size=12, color=INK, bold=False, italic=False):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), font)
    rfonts.set(qn("w:hAnsi"), font)
    rfonts.set(qn("w:cs"), CS_FONT)
    rfonts.set(qn("w:hint"), "cs")
    # RTL در سطح run — با حفظ ترتیب معتبر عناصر rPr (szCs پیش از rtl)
    for tag in ("w:rtl", "w:szCs"):
        for el in rpr.findall(qn(tag)):
            rpr.remove(el)
    szcs = OxmlElement("w:szCs")
    szcs.set(qn("w:val"), str(int(size * 2)))
    rpr.append(szcs)
    rtl = OxmlElement("w:rtl")
    rtl.set(qn("w:val"), "1")
    rpr.append(rtl)


def _rtl_para(p):
    ppr = p._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    ppr.append(bidi)


def add_par(
    container,
    text="",
    *,
    font=BODY_FONT,
    size=12,
    color=INK,
    bold=False,
    align=WD_ALIGN_PARAGRAPH.JUSTIFY,
    before=0,
    after=6,
    line=1.25,
    keep_next=False,
    style=None,
):
    p = container.add_paragraph(style=style) if style else container.add_paragraph()
    _rtl_para(p)
    pf = p.paragraph_format
    pf.alignment = align
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing = line
    if keep_next:
        pf.keep_with_next = True
    for chunk, is_bold in _split_bold(text):
        r = p.add_run(chunk)
        _set_run(r, font=font, size=size, color=color, bold=bold or is_bold)
    return p


def _split_bold(text: str):
    """«متن **پررنگ** ساده» → قطعه‌های (متن، پررنگ؟)."""
    parts = re.split(r"(\*\*.+?\*\*)", text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            yield part[2:-2], True
        else:
            yield part, False


def _shade(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def _cell_margins(cell, top=100, start=150, bottom=100, end=150):
    tcpr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for tag, val in (("top", top), ("left", start), ("bottom", bottom), ("right", end)):
        el = OxmlElement(f"w:{tag}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcpr.append(mar)


def _border(cell, edge="right", color=BAND_DARK, size=20):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = tcpr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcpr.append(borders)
    el = OxmlElement(f"w:{edge}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(size))
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), color)
    borders.append(el)


def add_band(doc, lines, *, fill, color=INK, bold_first=False, border=False, size=12):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.rows[0].cells[0]
    cell.width = Inches(6.9)
    _shade(cell, fill)
    _cell_margins(cell)
    if border:
        _border(cell)
    first = cell.paragraphs[0]
    for idx, line in enumerate(lines):
        p = first if idx == 0 else cell.add_paragraph()
        _rtl_para(p)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(line)
        _set_run(r, size=size, color=color, bold=(bold_first and idx == 0))
    return table


def spacer(doc, pts=6):
    p = doc.add_paragraph()
    _rtl_para(p)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1
    r = p.add_run("")
    _set_run(r, size=max(1, pts))
    return p


# ---------------------------------------------------------------- document
def build(data: dict, out_path: str | Path) -> Path:
    meta = data["meta"]
    doc = Document()

    # صفحه و حاشیه
    section = doc.sections[0]
    section.page_width = Twips(12240)
    section.page_height = Twips(15840)
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, attr, Inches(0.8))

    # استایل پایه
    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(12)
    normal.font.color.rgb = INK
    normal.paragraph_format.line_spacing = 1.25
    normal.paragraph_format.space_after = Pt(6)
    rpr = normal.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), BODY_FONT)
    rfonts.set(qn("w:hAnsi"), BODY_FONT)
    rfonts.set(qn("w:cs"), CS_FONT)

    # --- سرصفحهٔ عنوان ---
    add_par(
        doc, meta["doc_title"], font=TITLE_FONT, size=12, color=NAVY,
        align=WD_ALIGN_PARAGRAPH.CENTER, after=2, line=1.25,
    )
    add_par(
        doc, meta["doc_subtitle"], size=12, color=SLATE,
        align=WD_ALIGN_PARAGRAPH.CENTER, after=12, line=1.25,
    )
    add_band(doc, [meta["masthead"]], fill=BAND_LIGHT, color=INK, border=True)
    spacer(doc, 10)

    # --- مروری کلان ---
    ov = data["overview"]
    add_par(
        doc, ov["heading"], font=TITLE_FONT, size=12, color=NAVY,
        before=16, after=8, line=None, keep_next=True,
    )
    for text in ov["paras"]:
        add_par(doc, text)
    if ov.get("bullets"):
        for text in ov["bullets"]:
            p = doc.add_paragraph(text, style="List Bullet")
            _rtl_para(p)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.25
            p.paragraph_format.left_indent = Inches(0.25)
            p.paragraph_format.first_line_indent = Inches(-0.25)
            for r in p.runs:
                _set_run(r, size=12, color=INK)
    if ov.get("closing"):
        add_par(doc, ov["closing"])

    # --- بخش ۱ ---
    s1 = data["section1"]
    add_par(doc, s1["heading"], font=TITLE_FONT, size=14, bold=True, color=NAVY,
            before=16, after=8, line=None, keep_next=True)
    for text in s1["intro"]:
        add_par(doc, text)
    for axis in s1["axes"]:
        add_par(doc, axis["title"], size=12, color=SLATE, before=12, after=4,
                line=None, keep_next=True)
        add_par(doc, axis["body"], after=3)
    for text in s1.get("outro", []):
        add_par(doc, text, before=10, after=6, line=None)

    # --- بخش ۲ ---
    s2 = data["section2"]
    add_par(doc, s2["heading"], font=TITLE_FONT, size=14, bold=True, color=NAVY,
            before=16, after=8, line=None, keep_next=True)
    add_par(doc, s2["intro"], after=6, line=None)
    for layer in s2["layers"]:
        add_par(doc, layer["title"], font=TITLE_FONT, size=12, bold=True, color=NAVY,
                before=10, after=2, line=None, keep_next=True)
        add_par(doc, layer["body"], after=4, line=None)
    add_par(doc, s2["models_heading"], font=TITLE_FONT, size=12, bold=True, color=NAVY,
            before=12, after=4, line=None, keep_next=True)
    add_par(doc, s2["models_intro"], after=4, line=None)
    for model in s2["models"]:
        add_par(doc, model["body"], after=3, line=None)
    add_par(doc, s2["harmony_heading"], font=TITLE_FONT, size=12, bold=True, color=NAVY,
            before=12, after=4, line=None, keep_next=True)
    add_par(doc, s2["harmony"], after=6, line=None)

    # --- بخش ۳ ---
    s3 = data["section3"]
    add_par(doc, s3["heading"], font=TITLE_FONT, size=14, bold=True, color=NAVY,
            before=16, after=8, line=None, keep_next=True)
    add_par(doc, s3["body"])
    add_par(doc, s3["next_heading"], size=12, color=SLATE, before=12, after=4,
            line=None, keep_next=True)
    add_par(doc, s3["next"])

    # --- نوار پایانی ---
    spacer(doc, 8)
    add_band(doc, meta["closing_band"], fill=BAND_DARK, color=WHITE, bold_first=True)

    # فراداده
    props = doc.core_properties
    props.title = meta.get("doc_title", "")
    props.subject = meta.get("doc_subtitle", "")
    props.author = "آرون"
    props.language = "fa-IR"
    props.comments = ""

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return out_path


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="ساخت پروپوزال آرون از دادهٔ JSON")
    ap.add_argument("data", help="مسیر فایل JSON")
    ap.add_argument("-o", "--out", default="out/proposal.docx")
    args = ap.parse_args()
    payload = json.loads(Path(args.data).read_text(encoding="utf-8"))
    path = build(payload, args.out)
    print(f"saved: {path}")
