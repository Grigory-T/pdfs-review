# Development base

Maintain the general tool here. Keep one implementation, an installed-Python `uv`
environment, a small launcher, locked dependencies, focused tests and concise docs.
Task-specific copies may change paths/defaults but should not become competing
implementations of traversal or classification.

## Migrated decisions

| Reference | Reusable decision |
| --- | --- |
| Folder Review | Full materialized hierarchy per leaf; flat-table presentation; ordinary skips are not failures |
| Earlier page analysis | Text-block character counts; union image coverage; drawing count; page-level results |
| Readiness diagnostic | Independent page/text checks; no-text is valid; isolated parser warnings; optional complete ZIP reads; PDF bounds; timestamped outputs |

Use PyMuPDF for both classification and readiness, rather than adding a second
parser and Excel writer solely to reproduce one task. Readiness is engine-specific;
keep exact definitions in [readiness.md](readiness.md).

Do not migrate fixed ZIP/PDF populations, private locations, old result workbooks,
work-machine configuration, or dataset gates into public source. Source snapshots
and user-selected historical context may be retained in ignored `.local/` as
reference only. Old source locations need not be deleted.

Do not copy Folder Review's thread count into PDF processing. Its workload and
libraries differ. The stable baseline is sequential, with one text extraction per
page and disk-spooled results. Any parallel implementation must first prove
equivalent results, bounded memory/disk use, independent parser diagnostics,
deterministic occurrence IDs and continued processing after worker failures.

## Change verification

```sh
uv sync --frozen --no-dev
uv run --frozen python -m unittest discover -s tests -v
git diff --check
```

Tests use synthetic documents and archives only. Cover loose PDFs, arbitrary
ZIP/7z/RAR order, solid 7z, duplicate members, ten archive levels, Unicode/legacy
ZIP encodings, long physical/logical paths, path-safe extraction, literal Excel
strings, each page class, readiness states, errors and configured limits.

For a launcher change, run the test suite and a launcher smoke test on Windows
using synthetic inputs. Verify both sheets' values, formulas and flat-table style.
Do not claim Windows acceptance when that machine is unavailable. Recheck public
staged files for private paths, hostnames, logs, credentials and datasets before
committing/pushing. Never force-add ignored local context.

## Human copy/run workflow

Distribute source files, not `.venv`, runtime caches, inputs or generated reports.
Copy into a short local project path before running; do not execute from the
distribution share. Generic PowerShell example:

```powershell
Set-Location "C:\Projects"
& robocopy.exe "D:\ToolDelivery\pdfs-review" ".\pdfs-review" /E /R:2 /W:1 /XJ /XD .git .venv .runtime .local inputs outputs
if ($LASTEXITCODE -ge 8) { throw "robocopy failed with exit code $LASTEXITCODE" }
Set-Location ".\pdfs-review"
.\RUN.bat --root "C:\Data\inputs"
```

Replace only the source, local project and input paths. `/E` preserves destination
extras; it does not delete an existing input folder or workbook. Robocopy codes
below 8 are nonfatal. `RUN.bat` resolves installed `uv`, creates local `.venv`,
keeps caches/temp beside the code and forwards the Python exit code.
