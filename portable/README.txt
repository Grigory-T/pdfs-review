PDFS REVIEW - WINDOWS X86-64

1. Extract the whole pdfs-review-windows-x64.zip to a private, writable LOCAL folder.
2. Open pdfs-review.txt in Notepad. Set folder to your input path, then save.
3. Double-click pdfs-review.bat. Use that same BAT file for every review.

No system Python, uv, 7-Zip or administrator installation is needed. First start
downloads fixed, checksum-verified components. Later unchanged starts are offline.
Input can be local or SMB; code/runtime/results must stay on a local fixed drive.
Windows x86-64 with 64-bit Windows PowerShell 5.1 is required.

Reports receive timestamps and stay beside the BAT file. They contain PDF page
statistics, full hierarchy and per-file readiness. Existing reports are retained.
Keep the uniquely named JSONL companion with its matching workbook whenever
Excel-safe previews were needed. PDF summary and errors reports are also retained.

User settings use literal paths: no escaping or environment-variable expansion.
Quoted paths and UTF-8 / BOM UTF-16 are accepted. Unknown/duplicate/invalid settings
fail before scanning. Advanced options retain the established parser defaults.
The portable BAT accepts only --no-pause for scripted runs. Use Python directly
with the source checkout if you need the developer CLI.

Keep support intact; personal files belong outside it. All runtime, downloads,
caches, workspace lock and temporary work stay under support/state/. Temporary
work is removed on normal completion and controlled failures. A hard crash can
leave private work-* folders; close all runs before manually cleaning them.
When scanning the tool's parent folder, support is excluded. Input inside support
is rejected. Source documents are never rewritten, repaired or OCR-processed.

Only one operation may use this prepared tool folder at a time. Keep the WHOLE
folder when moving it: relocation rebuilds only its venv from cached wheels offline.
An incomplete cache fails explicitly. Use a fresh ZIP with online first setup
for an upgrade or repair, rather than mixing support files from different releases.
Setup diagnostics are off by default; diagnostics=yes enables a bounded neutral
setup journal. PDF results and console output can contain sensitive input paths.

Exit codes: 0 complete; 2 workbook saved with processing errors; 1 review/settings
failure; 10 setup/environment failure; 30 BAT launcher failure; 130 interruption.
Runtime checks are not an OS sandbox or strict PDF conformance validation.
See THIRD_PARTY.txt, UPSTREAM_LICENSE.txt and LICENSE.
