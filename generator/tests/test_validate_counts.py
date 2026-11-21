"""Counts are descriptive, never a minimum or maximum acceptance gate."""
import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import validate


def fa(number):
    return str(number).translate(str.maketrans('0123456789', '۰۱۲۳۴۵۶۷۸۹'))


class CountsTest(unittest.TestCase):
    def test_no_arbitrary_minimum_or_maximum(self):
        for count in (0, 1, 12, 30):
            with self.subTest(count=count), tempfile.TemporaryDirectory() as directory:
                doc = Document()
                doc.add_paragraph(validate.FIXED_TITLES[0])
                doc.add_paragraph(validate.FIXED_TITLES[1])
                doc.add_paragraph(validate.FIXED_TITLES[2])
                for i in range(count):
                    # Distinct letter labels, including multi-letter labels.
                    label = 'محور' + 'ب' * (i + 1)
                    doc.add_paragraph(f'{label}. عنوان')
                    doc.add_paragraph('متن نمونه ' * 40)
                doc.add_paragraph(validate.FIXED_TITLES[3])
                for i in range(count):
                    doc.add_paragraph(f'{fa(i + 1)}. عنوان لایه')
                    doc.add_paragraph('متن نمونه ' * 20)
                for title in validate.FIXED_TITLES[4:]:
                    doc.add_paragraph(title)
                path = Path(directory) / 'proposal.docx'
                doc.save(path)
                out = io.StringIO()
                with contextlib.redirect_stdout(out):
                    result = validate.main(path)
                self.assertEqual(result, 0, out.getvalue())
                self.assertIn(f'تعداد محورها: {count} ', out.getvalue())
                self.assertIn(f'تعداد لایه‌ها: {count} ', out.getvalue())

    def test_forbidden_mark_in_footer_is_detected(self):
        with tempfile.TemporaryDirectory() as directory:
            doc = Document()
            for title in validate.FIXED_TITLES:
                doc.add_paragraph(title)
            doc.sections[0].footer.paragraphs[0].text = "صفحه\u0654"
            path = Path(directory) / 'footer.docx'
            doc.save(path)
            out = io.StringIO()
            with contextlib.redirect_stdout(out):
                result = validate.main(path)
            self.assertEqual(result, 1)
            self.assertIn("FAIL", out.getvalue())



if __name__ == '__main__':
    unittest.main()
