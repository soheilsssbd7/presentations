# -*- coding: utf-8 -*-
"""
کنترل کیفیت پروپوزال آرون — چک‌لیست خودکار پیش از تحویل.

استفاده:
    python3 generator/validate.py templates/Aron-Proposal-Template.docx
    python3 generator/validate.py Aron-DotOne-Proposal-final.docx

خروجی: PASS / WARN / FAIL برای هر قاعده + کد خروج ۱ در صورت وجود FAIL.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

FIXED_TITLES = [
    "پیشنهاد هم‌مسیری و ساخت اعتبار پایدار",
    "مروری کلان",
    "۱. بررسی وضعیت فعلی: دریچه‌های ساخت اعتبار",
    "۲. راه‌حل چیست؟",
    "الگوهای جهانی",
    "هارمونی پیشنهادی آرون",
    "۳. مدل پیشنهادی همکاری و نحوه همراهی",
    "گام بعدی",
]

# محورها و لایه‌ها اختصاصی هر شرکت‌اند: تیتر و تعدادشان از شکاف واقعی
# همان شرکت می‌آید و هرگز از قالب شرکت دیگر کپی نمی‌شود.
AXIS_LABELS = "الفبجدهوزحطیکلمنسعفصقرشتثخ ذضظغ".replace(" ", "")
AXIS_RE = re.compile(r"^(الف|[" + AXIS_LABELS + r"])\. ")
LAYER_RE = re.compile(r"^[۱-۹]\. (?!بررسی|راه|مدل)")

FORBIDDEN = [
    "کمپین", "وایرال", "ترند", "تضمین", "رتبهٔ یک", "رتبه یک",
    "انقلابی", "بی‌نظیر", "جهش", "حذف رقیب", "۲۴/۷",
    "تومان کارکرد", "نفر-ماه", "SLA",
]

TYPOS = {
    "یکپارچه سازی": "یکپارچه‌سازی",
    "تامین ": "تأمین ",
    "برعهده": "بر عهده",
    "دات وان": "دات‌وان",
}

LATIN_DIGITS = re.compile(r"(?<![\w.])[0-9]+(?![\w.])")


def words(t):
    return len(t.split())


def main(path, forbid=()):
    doc = Document(path)
    paras = [p for p in doc.paragraphs]
    texts = [p.text.strip() for p in paras if p.text.strip()]
    joined = "\n".join(texts)
    fails, warns, oks = [], [], []

    def fail(m):
        fails.append(m)

    def warn(m):
        warns.append(m)

    def ok(m):
        oks.append(m)

    # ۱ — تیترهای ثابت
    for title in FIXED_TITLES:
        if any(t.startswith(title) for t in texts):
            ok(f"تیتر ثابت موجود: {title}")
        else:
            fail(f"تیتر ثابت غایب: {title}")

    # ۲ — حجم
    total = sum(words(t) for t in texts)
    if 900 <= total <= 2600:
        ok(f"حجم متن: {total} کلمه (بازهٔ ۹۰۰–۲۶۰۰)")
    else:
        warn(f"حجم متن خارج از بازه: {total} کلمه")

    # ۳ — واژه‌های ممنوع + نشت نام کارفرمایان دیگر
    if forbid:
        hits = [w for w in forbid if w and w in joined]
        if hits:
            fail("نام/واژهٔ متعلق به پروندهٔ دیگر در سند: " + "، ".join(hits))
        else:
            ok(f"هیچ‌یک از {len(forbid)} واژهٔ ممنوع پرونده‌های دیگر در سند نیست")

    for phrase in ("در مقایسه با", "برخلاف", "مشابه ", "همانند "):
        if phrase in joined:
            warn(f"عبارت مقایسه‌ای: «{phrase.strip()}» — سند باید مستقل و بدون ارجاع به پروندهٔ دیگر باشد")

    # ۴ — واژگان ممنوع
    for w in FORBIDDEN:
        if w in joined:
            fail(f"واژهٔ ممنوع: «{w}»")

    # ۵ — خطاهای املا/فاصله
    for bad, good in TYPOS.items():
        if bad in joined:
            warn(f"اصلاح لازم: «{bad.strip()}» → «{good.strip()}»")
    if "  " in joined:
        warn("فاصلهٔ تکراری در متن")
    if re.search(r"[\u0600-\u06FF]\(", joined):
        warn("پرانتز چسبیده به واژهٔ فارسی (فاصلهٔ پیش از ارجاع)")
    if re.search(r"[،؛.:!؟]\s*[،؛.:!؟]", joined):
        warn("نشانه‌گذاری تکراری")

    # ۶ — ساختار (تعداد و تیتر محورها/لایه‌ها آزاد و اختصاصی‌اند)
    axis_titles = [t for t in texts if AXIS_RE.match(t)]
    if len(axis_titles) >= 4:
        ok(f"تعداد محورها: {len(axis_titles)} (اختصاصی شرکت)")
    else:
        fail(f"تعداد محورها کم است: {len(axis_titles)}")
    if len({t[0] for t in axis_titles}) != len(axis_titles):
        warn("برچسب تکراری در محورها (الف/ب/ج/…)")
    layer_titles = [t for t in texts if LAYER_RE.match(t)]
    if len(layer_titles) >= 4:
        ok(f"تعداد لایه‌ها: {len(layer_titles)} (اختصاصی شرکت)")
    else:
        fail(f"تعداد لایه‌ها کم است: {len(layer_titles)}")

    # ۷ — طول محورها و لایه‌ها
    for t in texts:
        if AXIS_RE.match(t):
            i = texts.index(t)
            if i + 1 < len(texts):
                n = words(texts[i + 1])
                if not 35 <= n <= 80:
                    warn(f"طول نامتعارف محور ({n} کلمه): {t[:40]}")
    for t in texts:
        if LAYER_RE.match(t):
            i = texts.index(t)
            if i + 1 < len(texts):
                n = words(texts[i + 1])
                if not 15 <= n <= 40:
                    warn(f"طول نامتعارف لایه ({n} کلمه): {t[:40]}")

    # ۷ب — علامت ممنوعهٔ ٔ (U+0654) هرگز در متن نمی‌آید
    if "\u0654" in joined:
        fail("علامت ممنوعهٔ ٔ (U+0654) در متن هست")
    else:
        ok("علامت ممنوعهٔ ٔ در متن نیست")

    # ۸ — پرونوم و لحن
    ma = len(re.findall(r"(?<![\u0600-\u06FF])ما(?![\u0600-\u06FF])", joined))
    if ma <= 6:
        ok(f"کاربرد «ما»: {ma} بار")
    else:
        warn(f"کاربرد زیاد «ما»: {ma} بار (حداکثر ۶)")

    # ۹ — RTL و فونت
    non_rtl = [
        p.text[:40] for p in paras
        if p.text.strip() and p._p.find(qn("w:pPr")) is not None
        and p._p.pPr.find(qn("w:bidi")) is None
    ]
    if non_rtl:
        warn(f"{len(non_rtl)} پاراگراف بدون راست‌به‌چپ: {non_rtl[:3]}")
    else:
        ok("همهٔ پاراگراف‌ها راست‌به‌چپ")

    sizes = set()
    fonts = set()
    for p in paras:
        for r in p.runs:
            if not r.text.strip():
                continue
            if r.font.size:
                sizes.add(round(r.font.size.pt, 1))
            if r.font.name:
                fonts.add(r.font.name)
    allowed = {12.0, 14.0, 11.0, 10.5, 1.0}
    if sizes - allowed:
        warn(f"اندازه‌های غیرمجاز: {sorted(sizes - allowed)}")
    else:
        ok(f"اندازه‌های فونت مجاز: {sorted(sizes)}")

    # ۱۰ — خط لاتین/اعداد
    latin = []
    for m in LATIN_DIGITS.finditer(joined):
        ctx = joined[max(0, m.start() - 45):m.end() + 1]
        # سالِ داخل ارجاع لاتین مجاز است: (Musk, 2016)
        if re.search(r"[A-Za-z][^()]{0,40}, ?[0-9]{4}\)?$", ctx):
            continue
        latin.append(m.group(0))
    if latin:
        warn(f"عدد لاتین در متن فارسی: {latin[:6]}")
    else:
        ok("همهٔ اعداد فارسی‌اند")

    # ۱۱ — قیمت و تعهد
    if re.search(r"(قیمت|تعرفه|هزینهٔ خدمات|مبلغ قرارداد|نرخ خدمت)", joined):
        fail("اشاره به قیمت/تعرفه — این سند نباید قیمت داشته باشد")
    if re.search(r"(ماه|هفته)\s*(اول|دوم|سوم)\b", joined):
        warn("زمان‌بندی اجرایی در سند دیده می‌شود")

    # گزارش
    for m in oks:
        print(f"  PASS  {m}")
    for m in warns:
        print(f"  WARN  {m}")
    for m in fails:
        print(f"  FAIL  {m}")
    print(f"\nنتیجه: {len(oks)} قبول، {len(warns)} هشدار، {len(fails)} خطا")
    return 1 if fails else 0


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--forbid")]
    forbid = ()
    for a in sys.argv[1:]:
        if a.startswith("--forbid="):
            forbid = tuple(x.strip() for x in a.split("=", 1)[1].split(",") if x.strip())
    target = args[0] if args else "templates/Aron-Proposal-Template.docx"
    sys.exit(main(target, forbid))
