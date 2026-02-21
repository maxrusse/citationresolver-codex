# codex_citation

Codex skill project with two separate skills:
- `citation-resolver` for DOCX/Zotero citation repair.
- `citation-audit` for claim-level manuscript citation auditing.

The resolver workflow is focused on one job:
- take a Word `.docx` with broken plain numeric citations,
- rebuild live Zotero citation/bibliography fields,
- validate the result with a managed multi-step workflow.

## Managed Workflow (Always On)

The script does not run as blind one-shot. It executes step-by-step attempts with quality gates:

- detect reference section headings (e.g. References/Bibliography/Referenzen) and fallback to the strongest numbered reference block
- match/add references in local Zotero
- rebuild in-text citation fields
- insert/update bibliography field
- inject/update Zotero document prefs
- evaluate quality gates (`fields_ok`, `bibliography_ok`, `doc_prefs_ok`, `unmatched_count`)
- retry with fallback citation patterns when needed
- select best attempt and produce one final output file

## Requirements

- Python 3.10+
- Zotero running locally (for connector add operations)
- Local Zotero DB (default: `~/Zotero/zotero.sqlite`)
- Optional: Word + `pywin32` for `--word-update`

## Usage

```bash
python docx_zotero_integrator.py "/path/to/file.docx"
```

Optional:

```bash
# set first-attempt pattern preference (workflow may retry others)
python docx_zotero_integrator.py "/path/to/file.docx" --citation-pattern auto-safe
python docx_zotero_integrator.py "/path/to/file.docx" --citation-pattern auto
python docx_zotero_integrator.py "/path/to/file.docx" --citation-pattern paren

# run Word field refresh after generation
python docx_zotero_integrator.py "/path/to/file.docx" --word-update

# write report file
python docx_zotero_integrator.py "/path/to/file.docx" --report-json ./report.json
```

## Output

The JSON report includes:

- `workflow_mode` (`managed`)
- `managed_workflow` attempts and selected attempt
- `reference_section` detection details (`heading_idx`, `first_ref_idx`, `detected_by`)
- pattern detection stats
- matched/unmatched references
- conversion counts and Zotero match details
- optional Word update result

## Codex Skill

`SKILL.md` defines the operational workflow for `citation-resolver` so Codex can run this tool as a managed repair pipeline.

Install/update the local Codex skill in one command:

```bash
python install_skill.py
```

Optional:

```bash
python install_skill.py --codex-home "C:/path/to/.codex"
python install_skill.py --skill-name citation-resolver --dry-run
```

## Citation Audit Skill

The `citation-audit/` folder contains the second skill:
- `citation-audit/SKILL.md`
- `citation-audit/agents/openai.yaml`
- `citation-audit/references/*`

Install by copying that folder to:
- `<codex_home>/skills/citation-audit`
