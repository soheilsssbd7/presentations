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


if __name__ == '__main__':
    unittest.main()
