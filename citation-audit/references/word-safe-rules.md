# Word-Safe Rules for DOCX Editing

1. Prefer tooling that preserves full OOXML structure.
- Keep existing styles, numbering, relationships, and field codes intact.
- Avoid flattening rich-text runs into single-run paragraphs.

2. Use citation field repair for bibliographic integrity when needed.
- Run `docx_zotero_integrator.py` from the `citationresolver` skill on the corrected file.
- Keep the best attempt only after quality gates pass.

3. Avoid fragile fallback patterns.
- Do not bulk rewrite `word/document.xml` unless no safer option exists.
- If forced to touch OOXML directly, make minimal node-level edits and preserve untouched nodes byte-for-byte where possible.

4. Validate output before delivery.
- The file opens in Word.
- Citations/references remain in expected order.
- No missing styles or broken section layout.
