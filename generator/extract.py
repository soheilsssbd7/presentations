# -*- coding: utf-8 -*-
"""
استخراج محتوای پروپوزال مرجع آرون به یک مدل داده‌ای (JSON).

این اسکریپت ساختار سند مرجع را بر اساس لنگرهای متنی می‌شناسد و آن را به
قالبی تبدیل می‌کند که build_proposal.py می‌تواند عیناً بازتولید کند.
هدف: هیچ متنی دستی بازنویسی نشود تا خطای انسانی/تایپی وارد نشود.
"""
import json
import sys
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph


def body_blocks(doc):
    out = []
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}")[1]
        if tag == "p":
            out.append(("p", Paragraph(child, doc)))
        elif tag == "tbl":
            out.append(("t", Table(child, doc)))
    return out


def extract(path):
    doc = Document(path)
    blocks = body_blocks(doc)

    paras = [b[1].text for b in blocks if b[0] == "p"]
    tables = [b[1] for b in blocks if b[0] == "t"]

    # ---- لنگرها ----
    i_s1 = next(i for i, t in enumerate(paras) if t.startswith("۱. بررسی وضعیت"))
    i_s2 = next(i for i, t in enumerate(paras) if t.startswith("۲. راه"))
    i_s3 = next(i for i, t in enumerate(paras) if t.startswith("۳. مدل پیشنهادی"))
    i_models = next(i for i, t in enumerate(paras) if t.strip() == "الگوهای جهانی")
    i_harmony = next(i for i, t in enumerate(paras) if t.strip() == "هارمونی پیشنهادی آرون")
    i_next = next(i for i, t in enumerate(paras) if t.strip() == "گام بعدی")

    # ---- مروری کلان ----
    overview_heading = paras[2]
    # پیوست‌های مقدمه: از پاراگراف ۳ تا قبل از عنوان بخش ۱
    pre_s1 = paras[3:i_s1]
    # سه بولت، سه پاراگراف آخرِ قبل از بولت‌ها
    bullet_idx = [i for i, t in enumerate(pre_s1) if t.startswith("اعتبار گروه:")]
    bullets = pre_s1[bullet_idx[0]:bullet_idx[0] + 3]
    intro_paras = pre_s1[:bullet_idx[0]]
    overview_closing = pre_s1[bullet_idx[0] + 3]

    overview = {
        "heading": overview_heading,
        "paras": intro_paras,
        "bullets": bullets,
        "closing": overview_closing,
    }

    # ---- بخش ۱ ----
    s1 = [t for t in paras[i_s1 + 1:i_s2] if t.strip()]
    # آخرین دو پاراگراف = جمع‌بندی؛ بقیه = مقدمه + جفت‌های محور
    s1_outro = s1[-2:]
    s1_body = s1[:-2]
    # مقدمه = پاراگراف‌هایی که با «الف.» شروع نمی‌شوند و بلند هستند
    axes = []
    intro = []
    k = 0
    while k < len(s1_body):
        if s1_body[k].startswith(("الف.", "ب.", "ج.", "د.", "ه.", "و.")):
            axes.append({"title": s1_body[k], "body": s1_body[k + 1]})
            k += 2
        else:
            intro.append(s1_body[k])
            k += 1

    section1 = {
        "heading": paras[i_s1],
        "intro": intro,
        "axes": axes,
        "outro": s1_outro,
    }

    # ---- بخش ۲ ----
    s2_intro = paras[i_s2 + 1]
    layers = []
    k = i_s2 + 2
    while k < i_models:
        layers.append({"title": paras[k], "body": paras[k + 1]})
        k += 2
    models = paras[i_models + 2:i_harmony]
    section2 = {
        "heading": paras[i_s2],
        "intro": s2_intro,
        "layers": layers,
        "models_heading": paras[i_models],
        "models_intro": paras[i_models + 1],
        "models": [{"title": m.split(":")[0], "body": m} for m in models],
        "harmony_heading": paras[i_harmony],
        "harmony": paras[i_harmony + 1],
    }

    # ---- بخش ۳ ----
    section3 = {
        "heading": paras[i_s3],
        "body": paras[i_s3 + 1],
        "next_heading": paras[i_next],
        "next": paras[i_next + 1],
    }

    masthead = tables[0].rows[0].cells[0].text
    band_lines = tables[1].rows[0].cells[0].text.split("\n")

    return {
        "meta": {
            "doc_title": paras[0],
            "doc_subtitle": paras[1],
            "masthead": masthead,
            "closing_band": band_lines,
            "source_file": Path(path).name,
        },
        "overview": overview,
        "section1": section1,
        "section2": section2,
        "section3": section3,
    }


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "Aron-DotOne-Proposal-final.docx"
    dst = sys.argv[2] if len(sys.argv) > 2 else "generator/data/dotone.json"
    data = extract(src)
    Path(dst).write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"written {dst}")
    print(json.dumps({k: (len(v) if isinstance(v, list) else "…") for k, v in data.items()}, ensure_ascii=False))
