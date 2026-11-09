from __future__ import annotations

import io
import json
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

from openpyxl import load_workbook
from openpyxl.worksheet._writer import ALL_TEMP_FILES

import pdfs_review as app


class ExcelTextTests(unittest.TestCase):
    def convert(self, text):
        report = io.StringIO()
        exporter = app.ExcelText(report, "cell-details.jsonl")
        value = exporter.cell(
            text, sheet="pdf_files", row=7, column="parser_warnings", record={"pdf_id": 3}
        )
        details = [json.loads(line) for line in report.getvalue().splitlines()]
        return value, details, exporter

    def test_text_length_boundaries_and_unicode(self):
        for text in ("A" * 32767, "文" * 32767, "😀" * 16383 + "A"):
            with self.subTest(units=app.utf16_units(text)):
                value, details, exporter = self.convert(text)
                self.assertEqual(value, text)
                self.assertEqual(details, [])
        for text in ("A" * 32768, "文" * 40000, "😀" * 20000, "a" * 100001):
            value, details, exporter = self.convert(text)
            self.assertLessEqual(app.utf16_units(value), 32767)
            self.assertTrue(
                value.startswith("[Excel-safe preview; full value: cell-details.jsonl #1")
            )
            self.assertEqual(details[0]["value"], text)
            self.assertEqual(details[0]["pdf_id"], 3)
            self.assertEqual(details[0]["column"], "parser_warnings")
            self.assertEqual(exporter.cells, 1)
            self.assertIsNone(app.INVALID_XML_TEXT.search(value))

    def test_line_feed_limit(self):
        self.assertEqual(self.convert("warning\n" * 253)[1], [])
        original = "warning\n" * 254
        value, details, _ = self.convert(original)
        self.assertLessEqual(value.count("\n"), 253)
        self.assertEqual(details[0]["value"], original)
        self.assertEqual(details[0]["reasons"], ["line_feeds"])

    def test_xml_controls_and_surrogates_preserve_original(self):
        text = "Обычный текст 東京 😀\x00\x01\x0b\x0c\ud800X\udfff\ufffe\uffff"
        value, details, _ = self.convert(text)
        self.assertIsNone(app.INVALID_XML_TEXT.search(value))
        self.assertEqual(details[0]["value"], text)
        self.assertEqual(details[0]["reasons"], ["xml_characters"])
        self.assertIn("東京 😀", value)
        self.assertIn("\\u0000", value)

    def test_normal_literals_are_unchanged(self):
        for text in ("=HYPERLINK(1)", "#REF!", "Обычный текст 東京 😀", "a\tb\nc\rd"):
            self.assertEqual(self.convert(text)[0], text)
            self.assertEqual(self.convert(text)[1], [])


class ExportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="pdfs-export-test-")
        self.root = Path(self.temp.name)
        self.pages = self.root / "pages.jsonl"
        self.files = self.root / "pdfs.jsonl"
        self.outdir = self.root / "outputs"
        self.output = self.outdir / "review.xlsx"
        self.baseline_temps = set(ALL_TEMP_FILES)

    def tearDown(self):
        remaining = set(ALL_TEMP_FILES) - self.baseline_temps
        self.temp.cleanup()
        self.assertEqual(remaining, set(), "openpyxl leaked worksheet temp files")

    def write_spools(self, **changes):
        record = {
            "levels": ["=literal.pdf"],
            "pdf_id": 1,
            "page_number": 1,
            "file_name": "=literal.pdf",
            "class": "digital",
            "text_chars": 60,
            "readiness": "READY_WITH_TEXT",
            **changes,
        }
        for path in (self.pages, self.files):
            path.write_text(json.dumps(record, ensure_ascii=True) + "\n", encoding="utf-8")

    def build(self):
        return app.build_xlsx(self.pages, self.output, 1, 1, self.files, 1)

    def test_warning_overflow_and_all_string_columns_round_trip(self):
        text = "repeated parser warning\n" * 5000
        self.write_spools(
            parser_warnings=text,
            full_path="文" * 40000,
            page_error="bad\x00error",
            file_error="bad\ud800error",
            levels=["😀" * 20000],
        )
        result = self.build()
        self.assertGreater(result["excel_text_cells_adjusted"], 0)
        details = [
            json.loads(line)
            for line in (self.outdir / result["cell_details_file"]).read_text().splitlines()
        ]
        self.assertEqual(
            next(d["value"] for d in details if d["column"] == "parser_warnings"), text
        )
        self.assertTrue(any(d["column"] == "level_1" for d in details))
        wb = load_workbook(self.output)
        try:
            self.assertEqual(wb.sheetnames, ["pdf_pages", "pdf_files"])
            for sheet in wb:
                self.assertEqual(sheet.max_row, 7)
                self.assertEqual(sheet["B7"].value, "=ROW()-ROW($B$6)")
                headers = {cell.value: cell.column for cell in sheet[6]}
                self.assertEqual(sheet.cell(7, headers["file_name"]).value, "=literal.pdf")
                for cell in sheet[7][2:]:
                    self.assertNotEqual(cell.data_type, "f")
                    if isinstance(cell.value, str):
                        self.assertLessEqual(app.utf16_units(cell.value), 32767)
                        self.assertIsNone(app.INVALID_XML_TEXT.search(cell.value))
        finally:
            wb.close()

    def test_unadjusted_export_has_no_details_file(self):
        self.write_spools()
        result = self.build()
        self.assertEqual(result["excel_text_cells_adjusted"], 0)
        self.assertIsNone(result["cell_details_file"])
        self.assertEqual(list(self.outdir.glob("*.jsonl")), [])

    def test_mid_sheet_failure_closes_all_streams_and_preserves_old_output(self):
        self.write_spools(parser_warnings="x" * 40000)
        self.outdir.mkdir()
        self.output.write_bytes(b"previous output")
        original = app.add_flat_sheet

        def fail(*args, **kwargs):
            original(*args, **kwargs)
            if args[5] == "pdf_files":
                raise RuntimeError("synthetic sheet failure")

        with mock.patch.object(app, "add_flat_sheet", side_effect=fail):
            with self.assertRaisesRegex(RuntimeError, "synthetic sheet failure"):
                self.build()
        self.assertEqual(self.output.read_bytes(), b"previous output")
        self.assertEqual(list(self.outdir.iterdir()), [self.output])

    def test_save_failure_closes_zip_and_worksheet_tempfiles(self):
        self.write_spools()
        with mock.patch.object(app.ExcelWriter, "save", side_effect=OSError("synthetic disk full")):
            with self.assertRaisesRegex(OSError, "synthetic disk full"):
                self.build()
        self.assertEqual(list(self.outdir.iterdir()), [])

    def test_replace_failure_keeps_previous_workbook_and_details(self):
        self.write_spools(parser_warnings="x" * 40000)
        self.outdir.mkdir()
        self.output.write_bytes(b"previous output")
        with mock.patch.object(app.os, "replace", side_effect=PermissionError("locked")):
            with self.assertRaises(PermissionError):
                self.build()
        self.assertEqual(self.output.read_bytes(), b"previous output")
        self.assertEqual(list(self.outdir.iterdir()), [self.output])

    def test_report_write_failure_is_not_silently_ignored(self):
        self.write_spools(parser_warnings="x" * 40000)
        with mock.patch.object(app.ExcelText, "cell", side_effect=OSError("details unavailable")):
            with self.assertRaisesRegex(OSError, "details unavailable"):
                self.build()
        self.assertEqual(list(self.outdir.iterdir()), [])

    def test_end_to_end_oversized_diagnostics_remain_nonfatal(self):
        import fitz

        with fitz.open() as doc:
            doc.new_page().insert_text((30, 40), "Synthetic test text")
            data = doc.tobytes()
        inputs = self.root / "inputs"
        inputs.mkdir()
        (inputs / "sample.pdf").write_bytes(data)
        text = "synthetic parser warning\n" * 5000
        with (
            mock.patch.object(app, "PROJECT_DIR", self.root),
            mock.patch.object(fitz.TOOLS, "mupdf_warnings", return_value=text),
            redirect_stdout(io.StringIO()) as stdout,
        ):
            self.assertEqual(app.run(["--root", str(inputs), "--output", str(self.output)]), 0)
        summary = json.loads(self.output.with_name("review-summary.json").read_text())
        self.assertEqual(summary["errors"], 0)
        self.assertEqual(summary["parser_warnings"], 5000)
        self.assertEqual(summary["excel_text_fields"], {"pdf_files.parser_warnings": 1})
        self.assertEqual(summary["settings"]["text_threshold"], 40)
        self.assertEqual(summary["settings"]["zip_encoding"], "cp437")
        self.assertFalse(summary["settings"]["check_zip_crc"])
        self.assertIn("WARNING: 1 cells", stdout.getvalue())

    def test_locked_output_is_rejected_before_scan(self):
        self.outdir.mkdir()
        self.output.write_bytes(b"existing workbook")
        inputs = self.root / "inputs"
        inputs.mkdir()
        real_open = open

        def locked(path, mode="r", *args, **kwargs):
            if str(path) == app.fs_path(self.output) and mode == "rb+":
                raise PermissionError("synthetic Excel lock")
            return real_open(path, mode, *args, **kwargs)

        with (
            mock.patch.object(app, "PROJECT_DIR", self.root),
            mock.patch("builtins.open", side_effect=locked),
            mock.patch.object(app.Review, "scan") as scan,
        ):
            with self.assertRaisesRegex(PermissionError, "Close Excel"):
                app.run(["--root", str(inputs), "--output", str(self.output)])
        scan.assert_not_called()
        self.assertEqual(self.output.read_bytes(), b"existing workbook")


if __name__ == "__main__":
    unittest.main()
