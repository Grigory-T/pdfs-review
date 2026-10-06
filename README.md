# PDFs Review

Create one Excel table of PDF page statistics from an input folder. Scan ordinary
folders and PDFs inside ZIP, 7z and RAR archives, including archives nested in any
order. Other file extensions are skipped normally. An archive may contain zero,
one or many PDFs.

## Run

Requirements: Windows or Linux, installed [uv](https://docs.astral.sh/uv/), and
Python 3.11–3.14. RAR decompression also needs 7-Zip, unrar or unar. On Windows,
the usual `C:\Program Files\7-Zip\7z.exe` installation is discovered automatically.
ZIP and 7z processing do not need an external archive application.

```powershell
git clone https://github.com/Grigory-T/pdfs-review.git
Set-Location .\pdfs-review
.\RUN.bat --root "C:\Data\inputs"
```

For routine use, set `ROOT_FOLDER` near the top of `pdfs_review.py`, then run
`RUN.bat`. The launcher creates a local `.venv` and installs the locked
dependencies. It uses an installed Python without downloading another version.
First-time dependency installation needs package access.

On Linux:

```sh
uv sync --frozen --no-dev
uv run python pdfs_review.py --root /data/inputs
```

Results beside the script:

- `pdfs-review.xlsx`: one sheet, `pdf_pages`.
- `pdfs-review-summary.json`: counts, class totals and elapsed time.
- `pdfs-review-errors.txt`: paths and reasons for actual failures.

Use `--output` to change the result location. Exit codes: `0` for a completed run
without errors, `2` for a saved workbook with failures, `1` for a fatal failure,
and `130` for interruption. Non-PDF files and archives with no PDFs are not errors.

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
| `page_error` | Failure reason, otherwise blank |

Unreadable or password-protected PDFs remain error rows with a blank page number.
Failed pages keep their page number. Archive or folder failures also remain visible
as error rows with a blank page number; summary `pages` excludes those extra rows.
No file hashes or extracted text are included.

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
  Names are not normalized or shortened. A value exceeding Excel's cell limit fails
  explicitly rather than silently truncating it.
- Processing is sequential. PDF rows are spooled to local disk and Excel is written
  incrementally. A solid 7z archive is decoded once; selected members remain in
  temporary storage until that archive is processed. Solid RAR member reads can be slower.
- Defaults limit archive depth to 20, archive entries to 100,000, individual members
  to 50 GiB, and cumulative decompression to 100 GiB. `--max-depth` and
  `--max-extracted-gib` override the run limits. Limit failures are reported.
- Temporary files are removed after processing. The workbook is replaced atomically
  after a successful write. Close it in Excel before rerunning. Each run starts fresh.

## Adaptation and development

Keep the repository as the general base. For a particular task, make a separate
local checkout and change the input path, CLI defaults or a small configuration
section. Keep source documents and generated reports out of Git.

```sh
uv run python -m unittest discover -s tests -v
```

## License

AGPL-3.0-or-later. PyMuPDF is available under AGPL or a commercial license; see its
[licensing documentation](https://pymupdf.readthedocs.io/en/latest/about.html#license-and-copyright).
