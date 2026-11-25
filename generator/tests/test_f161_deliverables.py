"""Checks for the F-161 package; factual quality requires human review."""
import unittest
from pathlib import Path
from zipfile import ZipFile
from lxml import etree

ROOT = Path(__file__).resolve().parents[2]
FOLDER = ROOT / 'F-161-بانک-پاسارگاد'
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W, 'r': REL}


class F161Tests(unittest.TestCase):
    def test_footnotes_are_real_and_early(self):
        paths = sorted(FOLDER.glob('*.docx'))
        self.assertEqual(len(paths), 2)
        for path in paths:
            with self.subTest(path=path.name), ZipFile(path) as z:
                self.assertIsNone(z.testzip())
                document = etree.fromstring(z.read('word/document.xml'))
                notes = etree.fromstring(z.read('word/footnotes.xml'))
                rels = etree.fromstring(z.read('word/_rels/document.xml.rels'))
                refs = document.xpath('//w:footnoteReference', namespaces=NS)
                self.assertEqual(len(refs), 1)
                self.assertEqual(refs[0].get(f'{{{W}}}id'), '1')
                self.assertIn('عموم‌ها', ''.join(refs[0].getparent().getprevious().itertext()))
                text = ''.join(notes.xpath('//w:footnote[@w:id="1"]//w:t/text()', namespaces=NS))
                for term in ('Publics', 'مشتریان', 'کارکنان', 'جامعه محلی'):
                    self.assertIn(term, text)
                self.assertTrue(rels.xpath('//r:Relationship[contains(@Type,"/footnotes") and @Target="footnotes.xml"]', namespaces=NS))
                paragraphs = document.xpath('//w:body//w:p', namespaces=NS)
                index = next(i for i, p in enumerate(paragraphs) if p.xpath('.//w:footnoteReference', namespaces=NS))
                self.assertLess(index, 7)
                self.assertFalse(document.xpath('//w:br[@w:type="page"] | //w:lastRenderedPageBreak', namespaces=NS))
                for name in z.namelist():
                    if name.startswith('word/') and name.endswith('.xml'):
                        self.assertNotIn('\u0654', z.read(name).decode('utf-8'), name)

    def test_contact_brief_distinguishes_historical_and_current(self):
        path = FOLDER / 'Aron-F161-پاسارگاد-تحقیق-عموم‌ها-و-ارتباط.docx'
        with ZipFile(path) as z:
            root = etree.fromstring(z.read('word/document.xml'))
            text = '\n'.join(root.xpath('//w:t/text()', namespaces=NS))
        for value in ('راحله شهرابی', '۰۲۱-۸۲۸۹۰', 'info@bpi.ir',
                      '۰۲۱-۸۲۸۹۳۰۰۱', 'میرداماد', 'سمت و اختیار کنونی',
                      'توقف و عدم تماس', 'مجوز کتبی', 'متن پیشنهادی برای بررسی داخلی'):
            self.assertIn(value, text)

    def test_audit_and_scope(self):
        audit = (ROOT / 'research/F-161-PR-audit-internal.md').read_text()
        self.assertIn('f7ee975170af812e269f95560e840e0f805d7bd5', audit)
        self.assertIn('۱۱۰ ردیف', audit)
        self.assertIn('توقف و عدم تماس', audit)
        for n in ('۱','۲','۳','۴','۵','۶','۷','۸','۹','۱۰','۱۱','۱۲'):
            self.assertIn(f'| {n}. ', audit)
        for path in FOLDER.glob('*.docx'):
            with ZipFile(path) as z:
                content = z.read('word/document.xml').decode('utf-8')
                for forbidden in ('سهام', 'نماد', 'فرابورس', 'EPS', 'سبحان'):
                    self.assertNotIn(forbidden, content)


if __name__ == '__main__':
    unittest.main()
