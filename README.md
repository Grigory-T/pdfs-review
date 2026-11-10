# PDFs Review

Create one Excel workbook of PDF page statistics and file readiness from an input folder. Scan ordinary
folders and PDFs inside ZIP, 7z and RAR archives, including archives nested in any
order. Other file extensions are skipped normally. An archive may contain zero,
one or many PDFs.

## Windows: download, edit TXT, run BAT

[Download the Windows ZIP](../../releases/latest/download/pdfs-review-windows-x64.zip)
· [Releases and checksums](../../releases/latest)

1. Extract the whole ZIP into a private, writable **local** folder.
2. Open `pdfs-review.txt` in Windows Notepad, set `folder`, and save.
3. Double-click `pdfs-review.bat`. Use this same BAT file for every review.

No installed Python, uv, 7-Zip or administrator rights are needed. First start
downloads the fixed, checksum-verified Python/uv/packages/archive backend. Later
unchanged starts work offline, without invoking uv or checking for updates.
Windows x86-64 and 64-bit Windows PowerShell 5.1 are required. Network/proxy access
to the official component sources is needed for first setup only.

This follows [Folder Review](https://github.com/Grigory-T/folder-review)'s current
portable startup. Download the **Windows release ZIP**, not GitHub's source ZIP.

```text
pdfs-review.bat             the normal entry point
pdfs-review.txt             editable input and analysis settings
RUN.bat                    alias for the same entry point
support/                   verified source and local runtime/state
pdfs-review-<timestamp>.xlsx page statistics and per-PDF readiness
```

Reports, summaries, errors reports and any lossless JSONL companions stay beside
the BAT file. Timestamped filenames preserve earlier results automatically.
Runtime, cache, setup downloads and temporary work stay under `support/state/`,
not AppData/system temp. Only one operation may use the prepared folder at a time.
Normal completion and controlled failures remove temporary scan work; a hard
crash may leave private `work-*` directories. Close all runs before cleaning them.

Keep the whole prepared folder together when moving it: the local venv is rebuilt
from the retained cache offline. An incomplete moved cache fails clearly. For an
upgrade or repair, extract a fresh release rather than mixing support files.
Input can be local, mapped-drive or UNC/SMB. Code, runtime and results stay local.
The support folder is excluded from self-scans and cannot itself be selected as input.

### TXT settings

Paste paths literally, with no escaping. Quoted paths, percent signs, Unicode,
UTF-8 and BOM UTF-16 are supported. Relative input paths are relative to the TXT.
Unknown, duplicate and invalid settings fail before scanning. Edit values only,
keeping the `[review]` section. Optional analysis thresholds and safety limits
match the existing defaults documented below.

```ini
[review]
folder = C:\Data\inputs
report = pdfs-review.xlsx
```

`report` is a base filename, not a path; a timestamp is always appended. The
supplied TXT lists all supported options and permitted values. `diagnostics=no`
keeps the optional setup journal off; PDF summary/error reports are still produced.
`pdfs-review.bat --no-pause` is available for scripted runs. The portable BAT does
not accept the old `--root` scan arguments: edit the TXT or use the source CLI.

### Source CLI (Windows or Linux development)

An installed [uv](https://docs.astral.sh/uv/) and Python 3.11–3.14 are required
only for source development. RAR support additionally needs installed 7-Zip,
unrar or unar. The existing CLI and `ROOT_FOLDER` override remain available:

```sh
uv sync --frozen --no-dev
uv run python pdfs_review.py --root /data/inputs
```

Source CLI results beside the script by default:

- `pdfs-review.xlsx`: `pdf_pages` (page detail) and `pdf_files` (one row per PDF occurrence).
- `pdfs-review-summary.json`: counts, class/readiness totals and elapsed time.
- `pdfs-review-errors.txt`: paths and reasons for actual failures.
- `pdfs-cell-details-<unique>.jsonl`: created only when cells need Excel-safe
  previews; retains complete original values and their sheet/row/field locations.

Use `--output` to change the result location, or `--timestamp-output` to preserve
earlier workbooks without replacing a file open in Excel. Exit codes: `0` for a completed run
without errors, `2` for a saved workbook with failures, `1` for a fatal failure,
and `130` for interruption. Portable setup/environment failures return `10`, BAT
launcher failures `30`. Non-PDF files and archives with no PDFs are not errors.
Excel-safe previews are reported as export warnings, not PDF failures.

## Table

One row represents one PDF page. The table uses the flat format of
[Folder Review](https://github.com/Grigory-T/folder-review): a title, numbered
columns, blue field headers, black `№пп`/`last` boundaries, borders and autofilter.

| Fields | Meaning |
| --- | --- |
| `level_1 ... level_N`, `leaf_level` | Complete hierarchy relative to the input folder, including archive names and member folders |
| `full_path`, `file_name` | Original location and filename; `::` separates archive boundaries |
| `pdf_id` | Run-local PDF occurrence ID, shared by its page rows; duplicate archive names remain separate occurrences |
| `object_type` | `pdf` normally; archive/folder/file on a discovery failure |
| `file_size_bytes`, `in_archive`, `archive_depth` | File size and archive context |
| `pdf_page_count`, `page_number` | Total PDF pages and 1-based page number |
| `class`, `text_chars`, `img_cover`, `vector_count` | Page classification and its measurements |
| `page_read_ok`, `text_extract_ok`, `has_extractable_text` | Separate page-read and text checks; blank means not checked |
| `error_stage`, `page_error` | Failure stage and reason, otherwise blank |

Unreadable or password-protected PDFs remain error rows with a blank page number.
Failed pages keep their page number. Archive or folder failures also remain visible
as error rows with a blank page number; summary `pages` excludes those extra rows.
No file hashes or extracted text are included.

## File readiness

`pdf_files` repeats the same complete hierarchy and `pdf_id`, with readable/text-ready
flags, counts of pages with and without text, separate read/text/classification
error counts, encryption/password flags, repair status, parser diagnostics, and
informational header/EOF markers.

| Readiness | Meaning |
| --- | --- |
| `READY_WITH_TEXT` | All pages readable, every text extraction succeeded, at least one page has non-whitespace text |
| `READY_NO_TEXT_LAYER` | All pages readable and text extraction succeeded, but no page has non-whitespace text |
| `READY_TEXT_PARTIAL` | All pages readable, but at least one text extraction failed |
| `PDF_NOT_READY` | Open, member-read, page-read, password or configured-limit failure |

Readiness is an operational check using PyMuPDF, not strict PDF conformance, PDF/A
certification or proof of successful rendering. A no-text PDF is valid for this
workflow and is not a processing error. It may contain scans, blank pages or
vector artwork. Parser warnings and automatic repairs are recorded without
turning an otherwise readable PDF into a failure. Their counts are engine-specific.

Text presence uses **any non-whitespace text**, independently of the classifier's
40-character default. A classification failure remains a run error, but does not
discard successfully extracted text metrics or change an independently successful
readiness check. Empty-password encrypted PDFs are supported; other passwords
are not supplied or guessed. See [check definitions](docs/readiness.md).

## Page classification

Text characters come from PyMuPDF text blocks. Image coverage is the union of image
rectangles clipped to the page, without counting overlaps twice. Very small image
fragments are ignored. `vector_count` counts drawing objects, not individual strokes.

| Class | Default rule, evaluated in this order |
| --- | --- |
| `scan` | Image coverage ≥0.90 and fewer than 40 text characters |
| `ocr` | Image coverage ≥0.90 and at least 40 text characters |
| `digital` | Lower image coverage and at least 40 text characters |
| `vector` | Lower image coverage, fewer than 40 text characters and more than 100 drawing objects |
| `other` | Remaining pages |

These are structural heuristics. The tool does not perform OCR or prove that an
existing text layer came from OCR. CLI options expose the text, image and vector
thresholds for task-specific adaptation.

## Paths, encoding and archive handling

- PDF and archive matching is case-insensitive and based on `.pdf`, `.zip`, `.7z`
  and `.rar` extensions. PDF attachments embedded inside a PDF are not traversed.
- Windows access uses extended-length local and UNC paths. Keep the checkout and
  `.venv` in a short local path. Individual filesystem component limits still apply.
- Archive names are logical labels only. Selected members are streamed into short
  generated temporary filenames, so their original length or `../` components
  cannot direct extraction outside temporary storage. Directory links are skipped.
- ZIP UTF-8 flags and valid Unicode Path metadata are honored. Unmarked legacy
  names use the ZIP standard CP437. Specify `--zip-encoding cp866` or `cp1251` when
  the source archive uses a different unmarked encoding; it cannot be inferred reliably.
- Excel stores names as literal Unicode strings, including names starting with `=`.
  Ordinary names are not normalized or shortened. Oversized strings, excessive
  line feeds and XML-forbidden characters get explicitly marked Excel-safe
  previews. Complete original values are saved losslessly in a uniquely named
  JSONL sidecar beside the workbook, linked by an ID in each affected cell.
  Limits are 32,767 UTF-16 units and 253 line feeds per cell. A sidecar write
  failure remains fatal: the tool never silently discards the original value.
- Processing is sequential. PDF rows are spooled to local disk and Excel is written
  incrementally. A solid 7z archive is decoded once; selected members remain in
  temporary storage until that archive is processed. Solid RAR member reads can be slower.
- Selected ZIP members are always read to EOF and CRC-checked. `--check-zip-crc`
  additionally reads every skipped regular ZIP member without retaining it.
  This costs additional decompression time. ZIP counters describe **direct** PDF
  members, not PDFs found later in nested archives. Empty/non-PDF ZIPs are normal.
- Defaults limit archive depth to 20, archive entries to 100,000, individual members
  to 50 GiB, cumulative decompression to 100 GiB, individual PDFs to 2 GiB and
  PDFs to 10,000 pages. `--max-depth`, `--max-extracted-gib`, `--max-pdf-gib` and
  `--max-pages` override these run limits. Limit failures are reported. PDF size
  is checked before parsing, after an archive member has been captured to disk.
- Temporary files are removed after processing. The workbook is replaced atomically
  after a successful write. Worksheet XML streams are closed before cleanup even
  if export fails. A locked/read-only existing workbook is checked before scanning;
  close it in Excel or use `--timestamp-output`. A late write failure leaves the
  previous workbook intact. Each run starts fresh.

The summary also records effective parsing settings, the PyMuPDF version,
`excel_text_cells_adjusted`, affected fields/reasons and `cell_details_file`.
Keep any JSONL sidecar together with its workbook; it can contain sensitive paths
and parser messages. It is not an input dataset or an additional PDF analysis.

## Adaptation and development

Keep the repository as the general base. For a particular task, make a separate
local checkout and change the input path, CLI defaults or a small configuration
section. Keep source documents and generated reports out of Git.

This repository is the canonical development base. Folder Review supplies the
portable startup and hierarchy reference, not its worker count, comparison/duplicate
features or newer report layout. The PDF parser, page/file row grain and established
flat-table format remain unchanged. See [development and migration decisions](docs/development.md). Local context
belongs in ignored `.local/`, never in public commits. Do not run simultaneous
reviews from the same checkout because they share the local spool files.

```sh
uv run python -m unittest discover -s tests -v
python scripts/build_portable.py --output dist
```

The deterministic source-only release ZIP uses `portable/pins.json`, hashed
requirements exported from `uv.lock`, and a source integrity manifest. Runtime
downloads are fixed and verified before use. The launcher is adapted from the
MIT-licensed reference, with full attribution in `portable/UPSTREAM_LICENSE.txt`.
Build a clean release using `python scripts/build_portable.py --output dist --release`.
User settings never select different runtime/package versions. These checks are
not an OS sandbox; use an isolated environment for hostile documents/archives.

## License

AGPL-3.0-or-later. PyMuPDF is available under AGPL or a commercial license; see its
[licensing documentation](https://pymupdf.readthedocs.io/en/latest/about.html#license-and-copyright).
