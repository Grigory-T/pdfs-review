# Maintenance instructions

- This checkout is the canonical general-purpose PDF review tool.
- Keep one implementation, neutral public docs, synthetic tests and locked `uv`
  dependencies. Avoid one-off dataset logic in the base.
- Preserve complete hierarchy, case-insensitive extensions, Unicode and Windows
  long-path support, short path-safe archive captures, and the flat Excel format.
- Expected skips and no-text PDFs are normal. Keep actual failures visible and
  parser diagnostics separate. Use the definitions in `docs/readiness.md`.
- Use `.local/` only for private context/source snapshots. Never commit it, real
  documents, outputs, operational paths, machine details or credentials. Do not
  automatically access or transfer private input data.
- Run synthetic tests and `git diff --check` before publishing. If changing the
  launcher, verify on real Windows when available and report any missing checks.
- Keep processing sequential unless a tested, bounded process-based design is
  explicitly requested. Folder Review is a hierarchy/style reference, not a
  mandate to use its concurrency model.
