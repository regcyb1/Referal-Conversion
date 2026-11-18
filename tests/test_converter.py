from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import pymupdf

from convert import InputPDF, convert_pdfs, sheet_names
from pdf_extractor import ConversionError, extract_pdf, safe_sheet_name


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "NPTMF2503MBA/NPTMF2503MBA-BM038-3-M-ENT-ANP-TECHS.pdf"
needs_sample = unittest.skipUnless(SAMPLE.exists(), "sample PDFs are not stored in git")


class ParsingTests(unittest.TestCase):
    def test_names_follow_excel_limits(self):
        used = set()
        names = [safe_sheet_name(value, used) for value in (
            "'Bad:/Name[1]'", "A" * 60, "a" * 60, "History", "///", "''",
            "x" * 30 + "'tail", "X" * 30 + "'tail",
        )]
        self.assertEqual(names[0], "Bad__Name_1_")
        self.assertTrue(all(0 < len(name) <= 31 for name in names))
        self.assertEqual(len(names), len({name.casefold() for name in names}))
        self.assertTrue(all(not name.startswith("'") and not name.endswith("'") for name in names))
        self.assertNotIn("history", {name.casefold() for name in names})

    def test_index_sheet_is_optional(self):
        self.assertEqual(sheet_names([InputPDF(Path("a.pdf"))]), (None, ("a",)))

    def test_main_name_reserved_before_pdf_tabs(self):
        summary, names = sheet_names([InputPDF(Path("summary.pdf")), InputPDF(Path("other.pdf"), "Summary")], "Summary")
        self.assertEqual((summary, names), ("Summary", ("summary (2)", "Summary (3)")))

    @needs_sample
    def test_wrapped_headers_and_source_annotations(self):
        result = extract_pdf(SAMPLE)
        self.assertEqual(len(result.rows), 27)
        self.assertEqual(result.headers[4:7], ("Presentation (10%)", "Project (50%)", "Test (40%)"))
        self.assertEqual(result.rows[-1].flags, ("N", "N", "N"))
        self.assertEqual(result.rows[-1].values[4:9], (0, 0, 0, 0, 0))

    def test_unreadable_or_unrecognized_pdf_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            invalid = Path(directory) / "broken.pdf"
            invalid.write_bytes(b"This is not a PDF")
            with self.assertRaises(ConversionError):
                extract_pdf(invalid)
            blank = Path(directory) / "scan.pdf"
            with pymupdf.open() as document:
                document.new_page()
                document.save(blank)
            with self.assertRaisesRegex(ConversionError, "scanned PDF"):
                extract_pdf(blank)

    @needs_sample
    def test_missing_table_cannot_silently_drop_students(self):
        with patch.object(pymupdf.Page, "find_tables") as mocked:
            mocked.return_value.tables = []
            with self.assertRaisesRegex(ConversionError, "student count does not reconcile"):
                extract_pdf(SAMPLE)

    @needs_sample
    def test_invalid_batch_leaves_existing_output_unchanged(self):
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "keep.xlsx"
            output.write_bytes(b"original workbook")
            invalid = Path(directory) / "broken.pdf"
            invalid.write_bytes(b"not a PDF")
            with patch("convert._write_workbook") as writer:
                with self.assertRaises(ConversionError):
                    convert_pdfs([InputPDF(SAMPLE), InputPDF(invalid)], output, overwrite=True)
                writer.assert_not_called()
            self.assertEqual(output.read_bytes(), b"original workbook")

    @needs_sample
    def test_duplicate_inputs_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(ConversionError, "more than once"):
                convert_pdfs([InputPDF(SAMPLE), InputPDF(SAMPLE)], Path(directory) / "results.xlsx")


if __name__ == "__main__":
    unittest.main()
