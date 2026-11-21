# -*- coding: utf-8 -*-
"""
سازندهٔ یادداشت‌های تحلیلی آرون (Word، RTL)
ورودی: JSON با بلوک‌های h2 / p / bullets / table / note
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Twips

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_proposal import (  # noqa: E402
    BAND_LIGHT, BODY_FONT, CS_FONT, INK, NAVY, SLATE, TITLE_FONT, WHITE,
    _rtl_para, _set_run, add_band, add_bullets, add_page_numbers, add_par,
)

BORDER = "C9D2DF"


def _shade(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def _borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), BORDER)
        borders.append(el)
    tblPr.append(borders)


def _rtl_table(table):
    tblPr = table._tbl.tblPr
    bidi = OxmlElement("w:bidiVisual")
    tblPr.append(bidi)


def _cell(cell, text, *, header=False, size=10, width=None):
    cell.text = ""
    if width:
        cell.width = width
    p = cell.paragraphs[0]
    _rtl_para(p)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT if header else WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    for chunk, bold in _split(text):
        r = p.add_run(chunk)
        _set_run(r, size=size, color=NAVY if header else INK, bold=header or bold)


def _split(text):
    import re
    for part in re.split(r"(\*\*.+?\*\*)", text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**") and len(part) > 4:
            yield part[2:-2], True
        else:
            yield part, False


def _mark_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    trPr.append(el)


def build(data: dict, out_path: str | Path) -> Path:
    meta = data["meta"]
    doc = Document()
    section = doc.sections[0]
    section.page_width = Twips(12240)
    section.page_height = Twips(15840)
    for attr in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, attr, Inches(0.8))

    normal = doc.styles["Normal"]
    normal.font.name = BODY_FONT
    normal.font.size = Pt(11)
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

    add_par(doc, meta["title"], font=TITLE_FONT, size=16, bold=True, color=NAVY,
            align=WD_ALIGN_PARAGRAPH.CENTER, after=2, line=1.2)
    add_par(doc, meta["subtitle"], size=11, color=SLATE,
            align=WD_ALIGN_PARAGRAPH.CENTER, after=10, line=1.2)
    add_band(doc, [meta["band"]], fill=BAND_LIGHT, color=INK, border=True, size=10)
    doc.add_paragraph()

    for block in data["blocks"]:
        kind = block["type"]
        if kind == "h2":
            add_par(doc, block["text"], font=TITLE_FONT, size=12.5, bold=True, color=NAVY,
                    before=14, after=5, line=None, keep_next=True)
        elif kind == "p":
            add_par(doc, block["text"], size=11, after=6)
        elif kind == "bullets":
            add_bullets(doc, block["items"], indent=0.25, size=11)
        elif kind == "note":
            add_par(doc, block["text"], size=10, color=SLATE, before=4, after=8, line=1.2)
        elif kind == "table":
            headers = block["headers"]
            rows = block["rows"]
            table = doc.add_table(rows=1, cols=len(headers))
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = True
            _rtl_table(table)
            _borders(table)
            width = Inches(6.9 / len(headers))
            hdr = table.rows[0]
            _mark_header(hdr)
            for i, h in enumerate(headers):
                _cell(hdr.cells[i], h, header=True, size=10, width=width)
                _shade(hdr.cells[i], BAND_LIGHT)
            for row in rows:
                cells = table.add_row().cells
                for i, val in enumerate(row):
                    _cell(cells[i], val, size=9.5, width=width)
            add_par(doc, "", size=4, after=4)
        else:
            raise ValueError(f"unknown block: {kind}")

    if meta.get("page_numbers"):
        add_page_numbers(doc)

    props = doc.core_properties
    props.title = meta["title"]
    props.author = "آرون"
    props.language = "fa-IR"

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_path)
    return out_path


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("data")
    ap.add_argument("-o", "--out", required=True)
    args = ap.parse_args()
    payload = json.loads(Path(args.data).read_text(encoding="utf-8"))
    p = build(payload, args.out)
    print(f"saved: {p}")
