---
name: citation-resolver
description: Manual-first agent workflow for messy DOCX citation repair. Use when citations are dirty/inconsistent and require preflight inspection + cleanup checklist before script rebuild.
---

# Citation Resolver Skill

## Operating Mode

Default is `manual-first`, not script-first.

The agent must build and execute a cleanup checklist before running the integrator.

## Read-First Contract

Before any script execution, the agent must:

1. Read document internals (`document.xml`, plus comments/notes XML when present).
2. Summarize what was found (citation patterns, likely issues, missing/ambiguous parts).
3. Present a working plan/todo list for this specific file.
4. Only then run automation.

## Phase 1: Manual Preflight Checklist

1. Inspect `word/document.xml` text flow for citation styles and anomalies.
2. Inspect `word/comments.xml` (if present) for citation-related notes.
3. Inspect `word/footnotes.xml` and `word/endnotes.xml` (if present) for citation markers.
4. Identify mixed citation forms:
   - numeric in `()`, `[]`, superscript
   - malformed ranges/lists
   - orphan numbers with no reference entry
5. Identify identifier fragments in text/notes/comments:
   - DOI variants (`doi:`, `https://doi.org/...`, quoted/trailing punctuation)
   - PMID forms (`PMID: 12345678`)
6. Inspect references section detection readiness:
   - heading candidates (`References`, `Bibliography`, etc.)
   - numbered block continuity quality
7. Produce a todo list with explicit actions and expected impact.
8. Only proceed when preflight todo status is acceptable.

## Phase 2: Script Rebuild (After Checklist)

1. Run `docx_zotero_integrator.py` on input `.docx`.
2. Use preferred pattern for first attempt (`auto-safe` default), allow managed retries.
3. Validate quality gates (`fields_ok`, `bibliography_ok`, `doc_prefs_ok`, `unmatched_count`).
4. Optionally run Word field update.
5. Return final output path + JSON report + concise change log.

## Primary Command

```bash
python docx_zotero_integrator.py "/path/to/file.docx"
```

## Agent Todo Template

Use this checklist structure in responses before script execution:

```text
[ ] Detect citation pattern inventory
[ ] Review comments/footnotes/endnotes for citation hints
[ ] Extract DOI/PMID fragments from manuscript text
[ ] Validate reference-section anchor and numbered continuity
[ ] Flag ambiguous citations for manual decision
[ ] Confirm script run settings (pattern preference, word update yes/no)
```
