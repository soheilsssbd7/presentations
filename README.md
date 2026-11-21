# پروپوزال‌های آرون

مخزن مرجع برای تولید پروپوزال‌های هم‌خانواده با سند **«آرون ⇄ دات‌وان»**.

| مسیر | چیست |
|---|---|
| `Aron-DotOne-Proposal-final.docx` | سند مرجع (۵ صفحه، ۱۲۸۰ کلمه) |
| `docs/Aron-Proposal-Playbook.md` | دفتر راهنما: معماری، ادبیات، استایل، قواعد و چک‌لیست |
| `templates/Aron-Proposal-Template.docx` | قالب آمادهٔ همان سند با جای‌نگهدارها |
| `generator/build_proposal.py` | مولد داده‌محور سند Word با همان هویت بصری |
| `generator/data/dotone.json` | محتوای سند مرجع به‌صورت داده (بازتولید ۱۰۰٪ عیناً متن) |
| `generator/data/template.json` | دادهٔ قالب با راهنمای هر پاراگراف |
| `generator/validate.py` | کنترل کیفیت خودکار پیش از تحویل |
| `docs/Global-Models-Playbook.md` | پروتکل انتخاب الگوهای جهانی برای هر شرکت + کتابخانهٔ کهن‌الگوها |
| `research/_TEMPLATE.md` | فرم دریافت دادهٔ هر شرکت جدید |
| `F-0xx-<شرکت>/` | پوشهٔ هر پرونده: پروپوزال Word + یادداشت تحقیق مخاطب و ارتباط |

## ساخت سند جدید

```bash
python3 generator/build_proposal.py generator/data/<company>.json \
    -o out/Aron-<Company>-Proposal-<year>.docx
python3 generator/validate.py out/Aron-<Company>-Proposal-<year>.docx
```

## ساخت قالب

```bash
python3 generator/build_proposal.py generator/data/template.json \
    -o templates/Aron-Proposal-Template.docx
```

## بازتولید سند مرجع (تست صحت مولد)

```bash
python3 generator/build_proposal.py generator/data/dotone.json -o out/regen.docx
```

متن خروجی باید با سند مرجع یکی باشد (۵۵ از ۵۵ پاراگراف). تفاوت‌ها فقط در سه اصلاح بصری
مستندشده در دفتر راهنما (فونت تیتر لایه‌ها، رنگ متن نوار پایانی، حذف پاراگراف‌های خالی) است.

## پیش‌نیاز

```bash
pip install --break-system-packages python-docx
```

فونت‌های `B Nazanin`، `B Titr` و `B Zar` باید روی سیستمی که فایل را باز می‌کند نصب باشند.
