"""Regression checks for the Persian deliverables; these do not verify facts."""
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree
import unittest

ROOT = Path(__file__).resolve().parents[2]
FOLDER = ROOT / 'F-052-پارس-آریان'


def word_text(path):
    with ZipFile(path) as archive:
        return '\n'.join(
            ''.join(ElementTree.fromstring(archive.read(name)).itertext())
            for name in archive.namelist()
            if name.startswith('word/') and name.endswith('.xml')
        )


class F052ContentChecks(unittest.TestCase):
    def test_deliverables_and_forbidden_text(self):
        docs = list(FOLDER.glob('*.docx'))
        self.assertEqual(len(docs), 2)
        for file in docs:
            text = word_text(file)
            for phrase in ('\u0654', 'سهام', 'سهم', 'نماد', 'فرابورس',
                           'سود هر', 'تحقیقات بازار آرون', 'پارس‌آریان، ۱۴۰۵'):
                self.assertNotIn(phrase, text, f'{file.name}: {phrase}')

    def test_no_stale_definitive_claims(self):
        text = '\n'.join(word_text(f) for f in FOLDER.glob('*.docx'))
        for phrase in ('دستاوردها بدون سنجه', 'گزارش سالانه غایب',
                       'سخنگوی مشخصی معرفی نشده', 'پاسخ‌گویی متغیر',
                       'بالاترین نرخ تبدیل', 'ارسال نامه یک‌صفحه‌ای'):
            self.assertNotIn(phrase, text)
        self.assertIn('گزارش‌های سالانه', text)
        self.assertIn('بدون تأیید', text)

    def test_historical_context_and_no_outreach(self):
        proposal = word_text(FOLDER / 'Aron-F052-پارسآریان-پروپوزال.docx')
        brief = word_text(FOLDER / 'Aron-F052-پارسآریان-تحقیق-مخاطب-و-ارتباط.docx')
        self.assertIn('۱۳۹۶', proposal)
        self.assertIn('بحران جاری', proposal)
        self.assertIn('جوابیه', proposal)
        self.assertIn('توقف و عدم تماس', brief)
        self.assertIn('مجوز کتبی', brief)
        self.assertNotIn('در هشدار عمومی خود', proposal)

    def test_internal_audit_covers_all_twelve(self):
        content = (ROOT / 'research/F-052-PR-audit-internal.md').read_text()
        for number in '۱۲۳۴۵۶۷۸۹':
            self.assertIn(f'| {number}. ', content)
        for number in ('۱۰', '۱۱', '۱۲'):
            self.assertIn(f'| {number}. ', content)


if __name__ == '__main__':
    unittest.main()
