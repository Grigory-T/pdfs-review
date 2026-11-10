from __future__ import annotations

import io
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

import fitz

import pdfs_review as app
from pdfs_settings import settings_args


class SettingsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="pdfs-settings-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.path = self.root / "pdfs-review.txt"

    def test_template_defaults_preserve_pdf_cli(self):
        template = (app.PROJECT_DIR / "pdfs-review.txt").read_text()
        self.path.write_text(template.replace("folder =\n", "folder = input\n"))
        args = app.parse_args(settings_args(self.path))
        old = app.parse_args([])
        for name in (
            "check_zip_crc",
            "zip_encoding",
            "max_depth",
            "max_extracted_gib",
            "max_pdf_gib",
            "max_pages",
            "text_threshold",
            "img_cover_threshold",
            "min_img_frac",
            "vector_threshold",
        ):
            self.assertEqual(getattr(args, name), getattr(old, name), name)
        self.assertTrue(args.timestamp_output)
        self.assertEqual(args.output, self.root / "pdfs-review.xlsx")

    def test_literal_paths_and_notepad_encodings(self):
        folder = self.root / "пример 100% ! # & (data)"
        for encoding in ("utf-8", "utf-8-sig", "utf-16", "utf-16-be"):
            text = f'[review]\nfolder = "{folder}"\nreport = overview.xlsx\ncheck_zip_crc = yes\n'
            data = text.encode(encoding)
            if encoding == "utf-16-be":
                data = b"\xfe\xff" + data
            self.path.write_bytes(data)
            args = app.parse_args(settings_args(self.path))
            self.assertEqual(args.root, str(folder))
            self.assertTrue(args.check_zip_crc)
            self.assertEqual(args.output, self.root / "overview.xlsx")

    def test_invalid_settings_are_rejected(self):
        examples = [
            "[review]\nfolder=",
            "[wrong]\nfolder=input",
            "[review]\nfolder=x\nfolder=y",
            "[DEFAULT]\nother=x\n[review]\nfolder=input",
            "[review]\nfolder=x\nworkers=8",
            "[review]\nfolder=x\ncheck_zip_crc=perhaps",
            "[review]\nfolder=x\ndiagnostics=bad",
        ]
        for key, value in (
            ("max_depth", "0"),
            ("max_pages", "1.2"),
            ("max_extracted_gib", "nan"),
            ("max_pdf_gib", "inf"),
            ("text_threshold", "-1"),
            ("vector_threshold", "-1"),
            ("img_cover_threshold", "1.1"),
            ("min_img_frac", "-0.1"),
            ("zip_encoding", "unknown-encoding"),
            ("zip_encoding", "base64_codec"),
        ):
            examples.append(f"[review]\nfolder=input\n{key}={value}")
        for report in (
            "",
            "../escape.xlsx",
            "a/b.xlsx",
            r"C:\outside.xlsx",
            "wrong.csv",
            "NUL.xlsx",
            "NUL .xlsx",
            "CONIN$.xlsx",
            "test?.xlsx",
            "test.xlsx.",
            "😀" * 100 + ".xlsx",
        ):
            examples.append(f"[review]\nfolder=input\nreport={report}")
        for text in examples:
            with self.subTest(text=text):
                self.path.write_text(text, encoding="utf-8")
                with self.assertRaises(ValueError):
                    settings_args(self.path)

    def test_relative_folder_and_threshold_overrides(self):
        self.path.write_text(
            "[review]\nfolder=input\ntext_threshold=1\nvector_threshold=0\n"
            "img_cover_threshold=0.75\nmin_img_frac=0\nzip_encoding=cp866\n"
        )
        args = app.parse_args(settings_args(self.path))
        self.assertEqual(args.root, str(self.root / "input"))
        self.assertEqual(args.text_threshold, 1)
        self.assertEqual(args.vector_threshold, 0)
        self.assertEqual(args.img_cover_threshold, 0.75)
        self.assertEqual(args.min_img_frac, 0)
        self.assertEqual(args.zip_encoding, "cp866")

    def test_support_excluded_and_input_inside_support_rejected(self):
        support = self.root / "support"
        support.mkdir()
        with fitz.open() as doc:
            doc.new_page()
            payload = doc.tobytes()
        (support / "must-not-be-scanned.pdf").write_bytes(payload)
        (self.root / "user.pdf").write_bytes(payload)
        output = self.root / "result.xlsx"
        runtime = self.root / "work"
        with (
            mock.patch.object(app, "PROJECT_DIR", self.root),
            mock.patch.object(app, "SUPPORT_DIR", support),
            mock.patch.object(app, "RUNTIME_DIR", runtime),
            redirect_stdout(io.StringIO()),
        ):
            self.assertEqual(app.run(["--root", str(self.root), "--output", str(output)]), 0)
            with self.assertRaisesRegex(ValueError, "support folder"):
                app.run(["--root", str(support), "--output", str(output)])
        import json

        self.assertEqual(
            json.loads(output.with_name("result-summary.json").read_text())["pdfs_found"], 1
        )


if __name__ == "__main__":
    unittest.main()
