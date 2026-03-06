Reconstruction Contract v0.4
(DOCX ↔ RAW JSON rules aligned with parser v43 + reconstructor v2.15 + schema v2.16)

0. Versioning
- RAW JSON MUST declare:
  - meta.schema_version
- This contract corresponds to:
  - rules_version = "0.4"
  - schema_version = "2.16"
  - parser_version = "v43"
  - reconstructor_version = "v2.15"

1. Current Operational Pipeline
- Current deterministic operational path:
  1) Word SaveAs materialization
  2) add_custom_attrs (my:id on root paragraphs/tables and table rows)
  3) parser (DOCX -> RAW JSON)
  4) reconstructor (RAW JSON + donor materialized DOCX -> reconstructed DOCX)
- tools/materialize_effective.py is deprecated and not part of the official path.

2. General Principles
2.1 Determinism
- Parser/reconstructor MUST preserve token order and significant whitespace.
- Reconstructor MUST NOT auto-normalize formatting beyond values provided by RAW.

2.2 No synthetic formatting defaults
- Parser/reconstructor MUST avoid guessed formatting values.
- Explicit zeros must be preserved when present in OOXML/RAW.

2.3 ID contract
- Root body paragraphs/tables and table rows are addressed by my:id.
- Reconstructor patch-operations rely on id/anchor/position/derive_from.

3. Supported Content (actual code behavior)
3.1 Root content types
- paragraph
- table

3.2 Run types
- Parsed: text, tab, break, sym, cr, softHyphen, noBreakHyphen, picture
- Reconstructed: text, tab, break, sym, cr, picture
- Not reconstructed by v2.15: softHyphen, noBreakHyphen

3.3 Complex objects
- Tables: parsed and reconstructed (surgical patching of tblPr/trPr/tcPr + row/cell paragraph ops)
- Pictures: parsed and reconstructed using existing relation_id
- Shapes: schema-allowed but out of reconstruction scope in v2.15

4. Parser rules (current)
4.1 Body traversal
- Parser reads root body children and accepts only w:p and w:tbl into content.

4.2 Paragraph extraction
- paragraph fields: type, id, p_style_id, runs (+ optional p_format)
- p_format stores alignment, indents, spacing, tabs, list_info, and booleans when present.

4.3 Run extraction
- Each emitted run has id and parent_id.
- xml:space="preserve" is captured as run.meta.preserve=true.
- First visual tab can be marked with run.meta.leading=true.

4.4 Tables
- table fields: type, id, rows (+ optional tblPr, tbl_grid)
- row fields: id, cells (+ optional trPr)
- cell fields: id, content (+ optional tcPr)

5. Reconstructor rules (current)
5.1 Donor-based reconstruction
- Reconstructor opens donor materialized DOCX package and patches document.xml.
- Root operations:
  - deleted=true removes root element
  - anchor+position reorder/insert
  - derive_from clones donor elements for new ids

5.2 Paragraph patching
- p_format patch supports:
  - alignment, text_alignment
  - indent_*_twip
  - space_*_twip, line_spacing_twip, line_rule, *autospacing, *_lines
  - keep_next, keep_lines, page_break_before, widow_control, contextual_spacing, snap_to_grid
  - tabs[]
  - list_info{numId, ilvl}

5.3 Runs emission
- text -> w:t (+ xml:space preserve heuristic)
- tab -> w:tab
- break -> w:br (+ optional type)
- cr -> w:cr
- sym -> w:sym (font/char fields expected by current implementation)
- picture -> w:drawing with existing relation_id

6. Known compatibility limitations
- RAW containing run.type softHyphen/noBreakHyphen is schema-valid but not reconstructable by v2.15.
- sym run reconstruction expects font/char fields, while parser writes encoded text for sym.
- Some r_format fields in schema are parsed but not emitted back by reconstructor (partial r_format round-trip).

7. Schema alignment notes for v2.16
- tabStop posTwip/val are optional in schema to allow default tab-stop behavior.
- Producers SHOULD still provide posTwip and val when explicit tab stops are required.

8. Testing expectations
- RAW should validate against schema/raw.schema.json.
- No run merge during parsing.
- Empty paragraphs preserved.
- Deterministic xml:space behavior.
