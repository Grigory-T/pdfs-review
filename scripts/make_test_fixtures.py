"""Generate synthetic-only Windows acceptance inputs with the developer venv."""

from __future__ import annotations

import argparse
import io
import zipfile
from pathlib import Path

import fitz
import py7zr


def pdf(pages):
    with fitz.open() as doc:
        for _ in range(pages):
            doc.new_page().insert_text(
                (30, 40), "Synthetic acceptance document with more than forty text characters."
            )
        return doc.tobytes()


def create(output):
    # Reuse the neutral RAR STORE constructor already covered by scanner tests.
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "tests"))
    from test_pdfs_review import rar_bytes

    stream = io.BytesIO()
    with py7zr.SevenZipFile(stream, "w") as seven:
        seven.writestr(rar_bytes([("Bottom/deep.pdf", pdf(2))]), "Layer/inner.rar")
    nested = io.BytesIO()
    with zipfile.ZipFile(nested, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("Top/middle.7z", stream.getvalue())
        archive.writestr("Top/" + "Лист東京" * 80 + ".PDF", pdf(1))
        archive.writestr("ignored.txt", b"Synthetic skipped file")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("loose.PDF", pdf(1))
        archive.writestr("outer.zip", nested.getvalue())
        archive.writestr("ignored.txt", b"Synthetic skipped file")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    create(parser.parse_args().output)
