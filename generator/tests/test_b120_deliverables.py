import json
import unittest
import zipfile
from pathlib import Path
from lxml import etree
from docx import Document

ROOT = Path(__file__).resolve().parents[2]
DIR = ROOT / 'B-120-بانک-ملت'
PROPOSAL = DIR / 'Aron-B120-بانک-ملت-پروپوزال.docx'
BRIEF = DIR / 'Aron-B120-بانک-ملت-تحقیق-عموم‌ها-و-ارتباط.docx'
NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


class B120Deliverables(unittest.TestCase):
    def test_structural_constraints(self):
        for path in (PROPOSAL, BRIEF):
            with self.subTest(path=path.name), zipfile.ZipFile(path) as z:
                for name in z.namelist():
                    if name.endswith('.xml'):
                        self.assertNotIn('\u0654', z.read(name).decode('utf-8'))
                document = etree.fromstring(z.read('word/document.xml'))
                notes = etree.fromstring(z.read('word/footnotes.xml'))
                self.assertTrue(document.xpath('//w:footnoteReference', namespaces=NS))
                self.assertIn('عموم‌ها', ''.join(notes.xpath('//w:t/text()', namespaces=NS)))
                body = ''.join(document.xpath('//w:body//w:t/text()', namespaces=NS))
                self.assertIn('بانک ملت', body)
                for word in ('پاسارگاد', 'وسبحان', 'EPS', 'سهام', 'نماد'):
                    self.assertNotIn(word, body)
        proposal = '\n'.join(p.text for p in Document(PROPOSAL).paragraphs)
        self.assertIn('آموخته این الگو:', proposal)
        self.assertNotIn('درس این الگو:', proposal)
        self.assertNotIn('گزارش تحقیقات بازار', proposal)
        brief = json.loads((ROOT / 'generator/data/B-120-brief.json').read_text(encoding='utf-8'))
        text = json.dumps(brief, ensure_ascii=False)
        self.assertIn('توقف و عدم تماس', text)
        self.assertIn('شهبازی', text)
        self.assertIn('تصدی فعلی و اختیار خرید تایید نشده', text)

    def test_full_report_revision_and_safety_limits(self):
        proposal = json.loads((ROOT / 'generator/data/B-120.json').read_text(encoding='utf-8'))
        brief = json.loads((ROOT / 'generator/data/B-120-brief.json').read_text(encoding='utf-8'))
        audit = (ROOT / 'research/B-120-PR-audit-internal.md').read_text(encoding='utf-8')
        ptext = json.dumps(proposal, ensure_ascii=False)
        btext = json.dumps(brief, ensure_ascii=False)
        for phrase in ('کارمزد ابتدا کسر', 'توقف برنامه‌ریزی‌شده', 'سیام', 'نقش بانک و شرکا'):
            self.assertIn(phrase, ptext)
        for phrase in ('توقف و عدم تماس', '۳۱ فروردین ۱۳۹۹', 'سمت و اختیار جاری', 'هیچ تماس یا ارسال انجام نشده'):
            self.assertIn(phrase, btext)
        for phrase in ('تمام پاراگراف‌ها', '۱۲۴ کارت', 'وارسی جامعی انجام نشده', 'پ۱۰۹ و پ۱۱۱', 'HOLD / NO OUTREACH'):
            self.assertIn(phrase, audit)
        self.assertNotIn('خوانش خط‌به‌خط تمام پاراگراف‌ها و تک‌تک ۱۲۴ ردیف احراز نشده است', audit)

    def test_solutions_express_communication_strategy(self):
        proposal = json.loads((ROOT / 'generator/data/B-120.json').read_text(encoding='utf-8'))
        layers = proposal['section2']['layers']
        self.assertEqual(len(layers), 7)
        for layer in layers:
            self.assertTrue(layer['title'].strip())
            self.assertTrue(layer['body'].strip())
        joined = ' '.join(layer['title'] + ' ' + layer['body'] for layer in layers)
        for phrase in ('نیاز هر گروه', 'شرایط استفاده', 'سیام و شکایت', 'توقف برنامه‌ریزی‌شده', 'راه‌های رسمی', 'اثر اجتماعی', 'رعایت رضایت و محرمانگی'):
            self.assertIn(phrase, joined)
        for phrase in ('معماری روایت', 'لحظه‌های نیاز', 'صاحب پاسخ', 'سفر مشتری', 'آماده می‌کنیم', 'تهیه می‌کنیم'):
            self.assertNotIn(phrase, joined)


if __name__ == '__main__':
    unittest.main()
