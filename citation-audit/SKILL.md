---
name: citation-audit
description: Claim-by-claim citation auditing for manuscript DOCX files with online source validation, minimal text correction, and structured XLSX audit logs.
---

# Citation Audit

## Use This Skill When
- You receive a manuscript `.docx` and must verify every cited claim against the actual source.
- The user asks for claim-level (not paragraph-level) citation checking.
- Output must include a corrected manuscript and a structured audit table.

## Required Outputs
- Corrected manuscript `.docx`.
- `Manuscript_Citation_Audit.xlsx` with 2 sheets (`Claims`, `References`).

## Required Workflow
1. Load manuscript and isolate main text.
: Exclude the reference list from claim extraction.

2. Extract atomic cited claims.
: For each sentence containing citation markers, split into atomic claims where needed.
: Assign IDs `C1, C2, ...`.
: Record exact quoted claim segment and cited reference number(s).

3. Validate each claim online against the cited source.
: Prefer DOI landing page, PubMed, journal full text, and official guideline pages.
: Check alignment for metric type, population, analysis level, terminology, and key numeric values.
: Classify each claim as `Fully supported`, `Partially supported`, or `Not supported`.

4. Correct only what is necessary.
: Keep wording, tone, and structure stable.
: If partially/not supported, minimally edit the claim or replace citation only when required.
: Keep reference numbering stable unless unavoidable.

5. Preserve Word compatibility.
: Prefer document-edit methods that preserve OOXML structure (styles, runs, fields, relationships).
: If citation fields are stale/broken, run:
`python <codex_home>/skills/citation-resolver/docx_zotero_integrator.py "<path-to-docx>"`
: Prefer the tool's `auto-safe` flow and quality gates (`fields_ok`, `bibliography_ok`, `doc_prefs_ok`, `unmatched_count`).
: Avoid raw OOXML rewrites that replace whole paragraph/run trees.

6. Build audit workbook.
: `Claims` sheet columns:
`Claim_ID`, `Manuscript_text_segment`, `Original_cited_refs`, `Issue_with_original`, `Corrected_cited_refs`, `Rationale_for_change`
: `References` sheet columns:
`Ref#`, `Short citation (FirstAuthor et al. Journal Year)`, `DOI`, `Study type (Guideline / Meta-analysis / Cohort / Review / Technical paper)`, `What it actually demonstrates`, `Which Claim_IDs it supports`

7. Final consistency checks.
: Every edited claim must map to at least one rationale row.
: Every referenced DOI in the workbook must map to at least one claim.
: Confirm the corrected `.docx` opens cleanly and keeps structure.

## Multi-Agent Pattern (Recommended)
- Split claim IDs into batches and validate in parallel agents.
- Reconcile disagreements using primary-source precedence (journal/DOI/PubMed over secondary summaries).
- Keep one canonical master table for final merge.

## References
- Prompt template for this workflow: `references/prompt-template.md`
- Word-safe handling notes: `references/word-safe-rules.md`



