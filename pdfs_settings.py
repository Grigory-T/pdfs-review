"""Literal Notepad settings for the portable entry point; no shell evaluation."""

from __future__ import annotations

import codecs
import configparser
import math
import re
from pathlib import Path

POSITIVES = {"max_depth", "max_pages", "max_extracted_gib", "max_pdf_gib"}
INTEGERS = {"max_depth", "max_pages", "text_threshold", "vector_threshold"}
FRACTIONS = {"img_cover_threshold", "min_img_frac"}
NUMBERS = POSITIVES | INTEGERS | FRACTIONS
ALLOWED = NUMBERS | {"folder", "report", "check_zip_crc", "zip_encoding", "diagnostics"}


def pasted_path(value: str) -> str:
    value = value.strip()
    if value.startswith(('"', "'")):
        if len(value) < 2 or value[-1] != value[0]:
            raise ValueError("Folder path has an unmatched quote")
        value = value[1:-1]
    if not value or any(c in value for c in ("\x00", "\r", "\n")):
        raise ValueError("Set folder in pdfs-review.txt to the input folder")
    return value


def settings_args(path: Path) -> list[str]:
    raw = path.read_bytes()
    encoding = "utf-16" if raw.startswith((b"\xff\xfe", b"\xfe\xff")) else "utf-8-sig"
    parser = configparser.ConfigParser(interpolation=None, strict=True)
    try:
        parser.read_string(raw.decode(encoding))
    except (UnicodeError, configparser.Error) as exc:
        raise ValueError("Use valid UTF-8 or BOM UTF-16 settings in pdfs-review.txt") from exc
    if parser.defaults() or parser.sections() != ["review"]:
        raise ValueError("Settings must have one [review] section and no DEFAULT values")
    settings = parser["review"]
    unknown = set(settings) - ALLOWED
    if unknown:
        raise ValueError(f"Unknown settings: {', '.join(sorted(unknown))}")
    folder = pasted_path(settings.get("folder", ""))
    # No environment expansion or interpolation: %, $, # and backslashes are literal.
    if not Path(folder).is_absolute():
        folder = str(path.parent / folder)
    report = settings.get("report", "pdfs-review.xlsx").strip()
    if (
        not report
        or report[-1] in " ."
        or len(report.encode("utf-16-le", errors="surrogatepass")) // 2 > 180
        or re.search(r'[<>:"/\\|?*\x00-\x1f]', report)
        or Path(report).suffix.casefold() != ".xlsx"
        or re.match(
            r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9]|CONIN\$|CONOUT\$)$",
            report.split(".", 1)[0].rstrip(" ."),
            re.I,
        )
    ):
        raise ValueError("report must be a normal .xlsx base filename, not a path")
    args = ["--root", folder, "--output", str(path.parent / report), "--timestamp-output"]
    for name in sorted(NUMBERS & set(settings)):
        try:
            value = int(settings[name]) if name in INTEGERS else float(settings[name])
        except ValueError as exc:
            raise ValueError(f"Invalid numeric setting: {name}") from exc
        if (
            not math.isfinite(value)
            or value < 0
            or (name in POSITIVES and value == 0)
            or (name in FRACTIONS and value > 1)
        ):
            raise ValueError(f"Numeric setting outside permitted range: {name}")
        args.extend(["--" + name.replace("_", "-"), str(value)])
    for name in ("check_zip_crc", "diagnostics"):
        if name in settings:
            try:
                enabled = settings.getboolean(name)
            except ValueError as exc:
                raise ValueError(f"{name} must be yes or no") from exc
            if name == "check_zip_crc" and enabled:
                args.append("--check-zip-crc")
    if "zip_encoding" in settings:
        try:
            codecs.lookup(settings["zip_encoding"])
            b"name".decode(settings["zip_encoding"])
        except (LookupError, ValueError) as exc:
            raise ValueError("zip_encoding must be a supported text encoding") from exc
        args.extend(["--zip-encoding", settings["zip_encoding"]])
    return args
