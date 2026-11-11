"""Structural checks for the HTML-to-DOCX portfolio demonstration.

Run: python -m unittest discover -s tests -v
"""
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]


class ConversionTests(unittest.TestCase):
    def test_structured_word_output(self):
        with tempfile.TemporaryDirectory() as temp:
            output = Path(temp) / "generated.docx"
            subprocess.run(
                [
                    sys.executable,
                    str(ROOT / "converter.py"),
                    str(ROOT / "examples" / "legal-agreement.html"),
                    str(output),
                ],
                check=True,
                capture_output=True,
                text=True,
            )

            document = Document(output)
            self.assertEqual(len(document.tables), 1)
            headings = [
                p for p in document.paragraphs
                if p.style.name in ("Heading 1", "Heading 2")
            ]
            self.assertGreaterEqual(len(headings), 10)
            self.assertTrue(all(
                p._p.get_or_add_pPr().find(qn("w:numPr")) is not None
                for p in headings
            ))
            self.assertIn(
                "Engagement and Scope",
                [p.text for p in headings],
            )

            with ZipFile(output) as archive:
                body = archive.read("word/document.xml")
                numbering = archive.read("word/numbering.xml")
                self.assertIn(b'TOC \\o "1-2"', body)
                self.assertIn(b"lowerLetter", numbering)
                self.assertIn(b"lowerRoman", numbering)
                self.assertIn(b"w:numPr", body)

            # A TOC field is created; page numbers still need a Word field refresh.


if __name__ == "__main__":
    unittest.main()
