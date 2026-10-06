from __future__ import annotations

import io
import json
import os
import struct
import tempfile
import unittest
import warnings
import zipfile
import zlib
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

import fitz
import py7zr
from openpyxl import load_workbook

import pdfs_review as app


def sample_pdf(pages=1):
    with fitz.open() as doc:
        for _ in range(pages):
            page = doc.new_page()
            page.insert_text(
                (30, 40), "This page contains more than forty extracted text characters."
            )
        return doc.tobytes()


def zip_bytes(members):
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in members:
            archive.writestr(name, payload)
    return stream.getvalue()


def rar_bytes(members):
    """Create a neutral RAR4 STORE archive; no external RAR creation utility."""

    def header(body):
        return struct.pack("<H", zlib.crc32(body) & 0xFFFF) + body

    result = b"Rar!\x1a\x07\x00" + header(struct.pack("<BHHHI", 0x73, 0, 13, 0, 0))
    for name, data in members:
        raw = name.encode("ascii")
        body = (
            struct.pack(
                "<BHHIIBIIBBHI",
                0x74,
                0x8000,
                32 + len(raw),
                len(data),
                len(data),
                3,
                zlib.crc32(data),
                0,
                20,
                0x30,
                len(raw),
                0o100644,
            )
            + raw
        )
        result += header(body) + data
    return result + header(struct.pack("<BHH", 0x7B, 0, 7))


def legacy_zip(name, data, encoding, unicode_name=None):
    raw = name.encode(encoding)
    extra = b""
    if unicode_name:
        field = b"\x01" + struct.pack("<I", zlib.crc32(raw)) + unicode_name.encode("utf-8")
        extra = struct.pack("<HH", 0x7075, len(field)) + field
    crc = zlib.crc32(data)
    local = (
        struct.pack(
            "<IHHHHHIIIHH",
            0x04034B50,
            20,
            0,
            0,
            0,
            0,
            crc,
            len(data),
            len(data),
            len(raw),
            len(extra),
        )
        + raw
        + extra
        + data
    )
    central = (
        struct.pack(
            "<IHHHHHHIIIHHHHHII",
            0x02014B50,
            20,
            20,
            0,
            0,
            0,
            0,
            crc,
            len(data),
            len(data),
            len(raw),
            len(extra),
            0,
            0,
            0,
            0,
            0,
        )
        + raw
        + extra
    )
    return (
        local
        + central
        + struct.pack("<IHHHHIIH", 0x06054B50, 0, 0, 1, 1, len(central), len(local), 0)
    )


class ReviewTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="pdfs-review-test-")
        self.root = Path(self.temp.name)
        self.inputs = self.root / "inputs"
        self.inputs.mkdir()

    def tearDown(self):
        self.temp.cleanup()

    def seven_zip(self, members):
        stream = io.BytesIO()
        with py7zr.SevenZipFile(stream, "w") as archive:
            for name, payload in members:
                archive.writestr(payload, name)
        return stream.getvalue()

    def scan(self, **overrides):
        args = app.parse_args(["--root", str(self.inputs)])
        for key, value in overrides.items():
            setattr(args, key, value)
        args.output = self.root / "result.xlsx"
        args.runtime = self.root / "runtime"
        args.runtime.mkdir(exist_ok=True)
        rows, errors = io.StringIO(), io.StringIO()
        app.configure_rar_backend(None)
        review = app.Review(args, rows, errors, args.runtime)
        review.scan(self.inputs)
        records = [json.loads(line) for line in rows.getvalue().splitlines()]
        self.assertEqual(list(args.runtime.iterdir()), [])
        return records, review

    def test_loose_case_insensitive_and_skips(self):
        (self.inputs / "document.PdF").write_bytes(sample_pdf(2))
        (self.inputs / "note.txt").write_text("ordinary text")
        (self.inputs / "empty.zip").write_bytes(zip_bytes([]))
        (self.inputs / "others.zip").write_bytes(zip_bytes([("note.doc", b"other file")]))
        rows, review = self.scan()
        self.assertEqual(len(rows), 2)
        self.assertEqual([r["page_number"] for r in rows], [1, 2])
        self.assertEqual(review.stats["errors"], 0)
        self.assertEqual(review.stats["skipped_files"], 2)

    def test_mixed_nested_archives(self):
        inner = zip_bytes([("Bottom/leaf.PDF", sample_pdf(2)), ("ignored.txt", b"skip")])
        rar = rar_bytes([("Middle/inner.zip", inner)])
        seven = self.seven_zip([("Layer/middle.rar", rar)])
        (self.inputs / "outer.zip").write_bytes(zip_bytes([("Top/middle.7z", seven)]))
        rows, review = self.scan()
        self.assertEqual(len(rows), 2)
        self.assertEqual(review.stats["errors"], 0, review.errors.getvalue())
        self.assertEqual(rows[0]["archive_depth"], 4)
        self.assertEqual(
            rows[0]["levels"],
            [
                "outer.zip",
                "Top",
                "middle.7z",
                "Layer",
                "middle.rar",
                "Middle",
                "inner.zip",
                "Bottom",
                "leaf.PDF",
            ],
        )
        self.assertTrue(
            rows[0]["full_path"].endswith(
                "outer.zip::Top/middle.7z::Layer/middle.rar::Middle/inner.zip::Bottom/leaf.PDF"
            )
        )

    def test_reverse_order_and_solid_7z(self):
        inner = self.seven_zip(
            [("one.pdf", sample_pdf()), ("two.pdf", sample_pdf(2)), ("other.txt", b"skip")]
        )
        middle = zip_bytes([("inner.7z", inner)])
        (self.inputs / "outer.7z").write_bytes(
            self.seven_zip([("nested.zip", middle), ("direct.pdf", sample_pdf())])
        )
        rows, review = self.scan()
        self.assertEqual(len(rows), 4)
        self.assertEqual(review.stats["errors"], 0, review.errors.getvalue())
        self.assertEqual(review.stats["pdfs_found"], 3)

    def test_long_unicode_and_literal_names(self):
        name = "Каталог/" + "Документ東京" * 60 + ".PDF"
        (self.inputs / "unicode.zip").write_bytes(
            zip_bytes([(name, sample_pdf()), ("=SUM(1,2).pdf", sample_pdf())])
        )
        rows, review = self.scan()
        self.assertEqual(review.stats["errors"], 0)
        self.assertEqual({r["file_name"] for r in rows}, {name.split("/")[-1], "=SUM(1,2).pdf"})
        spool = self.root / "rows.jsonl"
        spool.write_text(
            "".join(json.dumps(row, ensure_ascii=False) + "\n" for row in rows), encoding="utf-8"
        )
        output = self.root / "review.xlsx"
        app.build_xlsx(spool, output, len(rows), review.stats["max_level"])
        workbook = load_workbook(output)
        try:
            sheet = workbook["pdf_pages"]
            self.assertEqual(sheet["B6"].value, "№пп")
            self.assertEqual(sheet["B7"].value, "=ROW()-ROW($B$6)")
            self.assertEqual(sheet["C6"].fill.fgColor.rgb, "FF00B0F0")
            self.assertEqual(sheet.sheet_view.zoomScale, 85)
            self.assertIs(sheet.sheet_view.showGridLines, False)
            self.assertIsNone(sheet.freeze_panes)
            self.assertTrue(sheet.auto_filter.ref.startswith("B6:"))
            self.assertEqual(sheet.cell(6, sheet.max_column).value, "last")
            values = {
                c.value for row in sheet.iter_rows(min_row=7) for c in row if c.data_type == "s"
            }
            self.assertIn(name.split("/")[-1], values)
            self.assertIn("=SUM(1,2).pdf", values)
            self.assertFalse(
                any(c.data_type == "f" for row in sheet.iter_rows(min_row=7) for c in row[2:])
            )
        finally:
            workbook.close()

    def test_legacy_zip_encodings_and_unicode_metadata(self):
        for encoding in ("cp866", "cp1251"):
            with zipfile.ZipFile(
                io.BytesIO(legacy_zip("Отчёт.pdf", sample_pdf(), encoding))
            ) as archive:
                self.assertEqual(app.zip_name(archive.infolist()[0], encoding), "Отчёт.pdf")
        data = legacy_zip("plain.pdf", sample_pdf(), "cp437", unicode_name="Документ東京.pdf")
        (self.inputs / "extra.zip").write_bytes(data)
        rows, review = self.scan()
        self.assertEqual(review.stats["errors"], 0)
        self.assertEqual(rows[0]["file_name"], "Документ東京.pdf")

    def test_duplicate_zip_names_remain_separate_pdfs(self):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            data = zip_bytes([("same.pdf", sample_pdf()), ("same.pdf", sample_pdf(2))])
        (self.inputs / "duplicates.zip").write_bytes(data)
        rows, review = self.scan()
        self.assertEqual(review.stats["errors"], 0)
        self.assertEqual(len(rows), 3)
        self.assertEqual(len({r["pdf_id"] for r in rows}), 2)

    def test_errors_are_visible_and_password_pdf(self):
        (self.inputs / "broken.zip").write_bytes(b"broken")
        (self.inputs / "broken.pdf").write_bytes(b"not PDF")
        with fitz.open(stream=sample_pdf(), filetype="pdf") as doc:
            locked = doc.tobytes(
                encryption=fitz.PDF_ENCRYPT_AES_256, user_pw="sample", owner_pw="owner"
            )
        (self.inputs / "locked.pdf").write_bytes(locked)
        rows, review = self.scan()
        self.assertEqual(review.stats["errors"], 3)
        self.assertTrue(all(r["class"] == "error" and r["page_number"] is None for r in rows))
        self.assertEqual({r["object_type"] for r in rows}, {"pdf", "archive"})

    def test_page_failure_does_not_skip_later_pages(self):
        (self.inputs / "document.pdf").write_bytes(sample_pdf(2))
        original = app.classify_page

        def classify(page, **kwargs):
            if page.number == 0:
                raise ValueError("synthetic page failure")
            return original(page, **kwargs)

        with mock.patch.object(app, "classify_page", side_effect=classify):
            rows, review = self.scan()
        self.assertEqual([r["class"] for r in rows], ["error", "digital"])
        self.assertEqual([r["page_number"] for r in rows], [1, 2])

    def test_depth_and_decompression_limits(self):
        inner = zip_bytes([("leaf.pdf", sample_pdf())])
        (self.inputs / "outer.zip").write_bytes(zip_bytes([("inner.zip", inner)]))
        rows, review = self.scan(max_depth=1)
        self.assertEqual(review.stats["errors"], 1)
        self.assertEqual(rows[0]["object_type"], "archive")
        rows, review = self.scan(max_extracted_gib=1 / 1024**3)
        self.assertEqual(review.stats["errors"], 1)
        self.assertEqual(rows[0]["class"], "error")

    def test_member_names_never_direct_extraction(self):
        (self.inputs / "paths.zip").write_bytes(
            zip_bytes([("../outside.pdf", sample_pdf()), ("/absolute.pdf", sample_pdf())])
        )
        rows, review = self.scan()
        self.assertEqual(len(rows), 2)
        self.assertEqual(review.stats["errors"], 0)
        self.assertFalse((self.root / "outside.pdf").exists())

    def test_windows_path_conversion(self):
        with (
            mock.patch.object(app.os, "name", "nt"),
            mock.patch.object(app.os.path, "abspath", side_effect=lambda p: str(p)),
        ):
            self.assertEqual(app.fs_path(r"C:\Data\file.pdf"), r"\\?\C:\Data\file.pdf")
            self.assertEqual(
                app.fs_path(r"\\server\share\file.pdf"), r"\\?\UNC\server\share\file.pdf"
            )
            self.assertEqual(app.fs_path(r"\\?\C:\Data\file.pdf"), r"\\?\C:\Data\file.pdf")

    def test_run_writes_summary_and_returns_error_status(self):
        (self.inputs / "document.pdf").write_bytes(sample_pdf(2))
        runtime_project = self.root / "project"
        runtime_project.mkdir()
        output = self.root / "out" / "review.xlsx"
        with mock.patch.object(app, "PROJECT_DIR", runtime_project), redirect_stdout(io.StringIO()):
            self.assertEqual(app.run(["--root", str(self.inputs), "--output", str(output)]), 0)
            (self.inputs / "bad.pdf").write_bytes(b"bad")
            self.assertEqual(app.run(["--root", str(self.inputs), "--output", str(output)]), 2)
        summary = json.loads(output.with_name("review-summary.json").read_text())
        self.assertEqual(summary["pages"], 2)
        self.assertEqual(summary["rows"], 3)
        self.assertEqual(summary["errors"], 1)

    def test_physical_paths_over_260_characters(self):
        directory = self.inputs
        for i in range(4):
            directory = directory / (f"level-{i}-" + "a" * 70)
            os.makedirs(app.fs_path(directory), exist_ok=True)
        with open(app.fs_path(directory / "document.PDF"), "wb") as handle:
            handle.write(sample_pdf())
        rows, review = self.scan()
        self.assertEqual(len(rows), 1)
        self.assertEqual(review.stats["errors"], 0)

    def test_ten_archive_levels(self):
        data = zip_bytes([("bottom.pdf", sample_pdf())])
        for index in range(9):
            data = zip_bytes([(f"folder-{index}/inner.zip", data)])
        (self.inputs / "outer.zip").write_bytes(data)
        rows, review = self.scan()
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["archive_depth"], 10)
        self.assertEqual(review.stats["errors"], 0)

    def test_empty_input_is_successful(self):
        output = self.root / "empty.xlsx"
        with mock.patch.object(app, "PROJECT_DIR", self.root), redirect_stdout(io.StringIO()):
            self.assertEqual(app.run(["--root", str(self.inputs), "--output", str(output)]), 0)
        workbook = load_workbook(output, read_only=True)
        try:
            self.assertEqual(sum(1 for _ in workbook["pdf_pages"].iter_rows()), 6)
        finally:
            workbook.close()


class ClassifierTests(unittest.TestCase):
    def classify(self, page, **kwargs):
        settings = {
            "text_threshold": 40,
            "img_cover_threshold": 0.9,
            "min_img_frac": 0.005,
            "vector_threshold": 100,
        }
        settings.update(kwargs)
        return app.classify_page(page, **settings)

    def image(self):
        pixmap = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 4, 4), False)
        pixmap.clear_with(128)
        return pixmap.tobytes("png")

    def test_all_classes_and_vector_boundary(self):
        with fitz.open() as doc:
            page = doc.new_page(width=300, height=600)
            self.assertEqual(self.classify(page)[0], "other")
            for i in range(101):
                page.draw_rect(fitz.Rect(i, 10, i + 1, 11))
            self.assertEqual(self.classify(page)[0], "vector")
            self.assertEqual(self.classify(page, vector_threshold=101)[0], "other")
            page.insert_text(
                (20, 50), "This page contains more than forty extracted text characters."
            )
            self.assertEqual(self.classify(page)[0], "digital")
            page.insert_image(page.rect, stream=self.image(), keep_proportion=False)
            self.assertEqual(self.classify(page)[0], "ocr")
            page = doc.new_page(width=300, height=600)
            page.insert_image(page.rect, stream=self.image(), keep_proportion=False)
            self.assertEqual(self.classify(page)[0], "scan")

    def test_rotated_and_cropped_full_page_scan(self):
        with fitz.open() as doc:
            page = doc.new_page(width=300, height=600)
            page.insert_image(page.rect, stream=self.image(), keep_proportion=False)
            for rotation in (0, 90, 180, 270):
                page.set_rotation(rotation)
                self.assertEqual(self.classify(page)[:3], ("scan", 0, 1.0))
            page.set_rotation(0)
            page.set_cropbox(fitz.Rect(20, 30, 280, 550))
            self.assertEqual(self.classify(page)[:3], ("scan", 0, 1.0))

    def test_overlaps_are_counted_once(self):
        self.assertEqual(app._union_area([(0, 0, 10, 10), (5, 0, 15, 10)]), 150)
        with fitz.open() as doc:
            page = doc.new_page(width=300, height=300)
            for _ in range(2):
                page.insert_image(
                    fitz.Rect(0, 0, 150, 300), stream=self.image(), keep_proportion=False
                )
            self.assertEqual(self.classify(page)[2], 0.5)


if __name__ == "__main__":
    unittest.main()
