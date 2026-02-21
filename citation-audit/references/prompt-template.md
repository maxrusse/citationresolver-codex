# Generic Prompt Template: Claim-Level Citation Audit (DOCX -> DOCX + XLSX)

You will receive a full manuscript (`.docx`). Perform a complete claim-by-claim citation audit.

## Scope
- Work per claim, not per paragraph.
- Exclude the reference list from claim extraction.
- Validate cited claims against the real source online (DOI, PubMed, journal, official guideline pages).

## Phase 1 - Claim Extraction
1. Extract every sentence segment that contains a citation.
2. Split into atomic claims where needed.
3. Assign `Claim_ID` values (`C1, C2, ...`).
4. Record for each claim:
- `Claim_ID`
- `Exact manuscript text segment`
- `Cited reference number(s)`

## Phase 2 - Online Validation (Per Claim)
For each claim:
1. Retrieve the cited source online.
2. Verify support for:
- reported metric type
- population/context
- analysis level (per-patient / per-lesion / etc.)
- numerical values
- terminology precision
3. Label outcome:
- `Fully supported`
- `Partially supported`
- `Not supported`
4. If unsupported, either:
- minimally adjust claim wording to match cited evidence, or
- replace citation only when strictly necessary.

## Phase 3 - Manuscript Correction Rules
- Keep edits minimal and precise.
- Preserve original style and structure.
- Keep reference numbering stable unless unavoidable.
- Remove overstatements and unsupported superlatives.

## Deliverables (required)
1. Corrected manuscript `.docx`.
2. `Manuscript_Citation_Audit.xlsx` with:
- Sheet `Claims` columns:
`Claim_ID`, `Manuscript_text_segment`, `Original_cited_refs`, `Issue_with_original`, `Corrected_cited_refs`, `Rationale_for_change`
- Sheet `References` columns:
`Ref#`, `Short citation (FirstAuthor et al. Journal Year)`, `DOI`, `Study type (Guideline / Meta-analysis / Cohort / Review / Technical paper)`, `What it actually demonstrates`, `Which Claim_IDs it supports`


