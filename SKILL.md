---
name: citationresolver
description: Managed end-to-end DOCX citation repair for Zotero. Use when citations are broken, plain-text, or stale and must be rebuilt into live Zotero fields with verification.
---

# CitationResolver Skill

## Managed Workflow (Required)

1. Run `docx_zotero_integrator.py` on the input `.docx`.
2. Detect reference section adaptively (References/Bibliography/Referenzen headings or strongest numbered block fallback).
3. Execute attempt 1 with preferred pattern (`auto-safe` by default).
4. Evaluate quality gates (`fields_ok`, `bibliography_ok`, `doc_prefs_ok`, `unmatched_count`).
5. Retry fallback patterns automatically if needed.
6. Select best attempt and return one final output.
7. Optionally run Word field update.
8. Return output path + JSON report summary.

## Primary Command

```bash
python docx_zotero_integrator.py "/path/to/file.docx"
```
