# Check definitions

## Two row grains

`pdf_pages` contains one row per attempted PDF page, plus explicit discovery/open
failure rows with no page number. `pdf_files` contains one row per encountered PDF
occurrence, including an unreadable selected archive member. Both use the same
run-local `pdf_id`; identical member names remain separate occurrences.

Full hierarchy columns start with the first object below the input folder, include
archive names and member folders, and end with the PDF filename. Ancestors are
repeated on every row, so filtering and sorting do not discard location context.
The absolute `full_path` uses `::` at archive boundaries. This is a logical path:
it is not a physical extracted filename or an automatic hyperlink into an archive.

## Independent checks

1. Read the PDF header/EOF marker regions as informational flags.
2. Open with PyMuPDF. Try an empty password if authentication is needed.
3. Reject zero-page documents and configured size/page limits explicitly.
4. Load every page and check that its page rectangle is finite and nonempty.
5. Extract text blocks once per page. Count all span characters for classification;
   separately check whether any span has non-whitespace text for readiness.
6. Calculate image coverage and drawing counts for classification. A failure here
   does not erase successful page/text checks.
7. Capture parser diagnostics for this PDF only, including any automatic repair.

Each page failure is isolated, and later pages are still attempted. Unknown checks
are blank, not false or zero. `pages_without_text` counts successful extractions
with no meaningful text, not pages whose extraction failed.

For a fully examined PDF:

```text
pdf_page_count = pages_read + page_read_errors
pages_read = pages_with_text + pages_without_text + text_extraction_errors
```

`readable` requires every page to pass the page-read check and no file-level
failure. `text_ready` additionally requires every extraction to succeed and at
least one page with text. These flags do not promise text on every page, semantic
quality, accessibility, strict structural conformance or rendering correctness.

Missing markers, repair and parser warnings alone do not fail readiness. MuPDF
diagnostic line counts cannot be compared directly with another parser's warning
count. No extracted text or document metadata content is written to the workbook.
Paths and parser/error messages can contain sensitive source information: keep
generated workbooks and logs local to the machine processing the input.

## ZIP checks and run outcome

Reading a selected ZIP member through EOF verifies its CRC. With `--check-zip-crc`,
unselected regular members are also consumed once and discarded. No-PDF and empty
ZIPs can pass. `zips_full_crc_passed` counts archives whose complete requested
member reads succeeded; `zips_full_crc_failed` counts opened archives with failed
member reads. Unopenable ZIPs remain archive errors and make `zips_found` exceed
`zips_opened`. A limit/password failure can prevent complete checking and must
not be interpreted as proof of a bad CRC.

ZIP presence counters concern direct PDF members only. An outer ZIP containing
another archive can have no direct PDFs but still lead to PDF results. RAR/7z
decoder checks are not reported as full ZIP CRC checks.

A run succeeds with exit 0 when traversal, page/text/classification processing and
workbook verification complete without errors. Skips, no-text PDFs and parser
warnings are allowed. Exit 2 means reports were saved but actual failures need
review. Exit 1 is fatal, and exit 130 is interruption.

The writer verifies both saved sheet row counts and reconciles PDF discovery with
file overview rows. Summary totals derive from processing counters, not fixed
population expectations. Dataset-specific expected counts belong in a separate
adaptation, not the universal tool.
