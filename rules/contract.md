Reconstruction Contract v0.2
(DOCX ↔ RAW JSON Core Rules, synced with UltimateReconstructorV10 + RAW JSON Schema v2.8.2)

0. Versioning
- RAW JSON MUST declare:
  meta.schema_version
  meta.rules_version
- This contract corresponds to:
  rules_version = "0.2"
  schema_version = "2.8.2"
  parser_version = "v41"
  reconstructor_version = "v10"

1. General Principles
1.1 Goal
- The goal of RAW is deterministic 1:1 visual reconstruction of a constrained DOCX subset (“forms”).
- RAW JSON is a structural IR. No visual optimization is allowed during parsing.
- Reconstruction must be deterministic and schema-driven.

1.2 No synthesis
- Parser and reconstructor MUST NOT invent formatting that is not present in OOXML or RAW.
- If a value is absent in OOXML, it MUST NOT appear in RAW as a synthesized default.
  (Future normalization/reconstruction policies may be introduced as opt-in modes.)

1.3 Scope (v0.2)
Included:
- Paragraphs (<w:p>) with runs (<w:r>) and tokens <w:t>, <w:tab/>, <w:br/>, <w:cr/>, <w:sym/>
- Basic paragraph formatting: alignment, indents, spacing, tabs, numbering refs
- Basic run formatting: fonts, size, bold/italic/underline/caps/color/vertAlign/lang/charSpacing/position
- Single-section page setup (sectPr: pgSz/pgMar/cols)
- document settings: defaultTabStop

Excluded / not guaranteed for deterministic behavior in this version:
- Tables, drawings/images/shapes, fields, hyperlinks, headers/footers, footnotes/endnotes
- Complex section breaks beyond one terminal sectPr
- Advanced style inheritance beyond what the parser explicitly resolves into RAW
- Run token types softHyphen and noBreakHyphen are allowed by schema but are NOT reconstructed by UltimateReconstructorV10 (see RULE-RUN-COMPATIBILITY and RULE-RUN-UNSUPPORTED)

2. Parsing Rules (DOCX → RAW)

RULE-P-001 — Hanging Indent Preservation
If paragraph contains:
  <w:ind w:hanging="X">
Then RAW MUST store:
  p_format.indentHangingTwip = X
This value MUST be preserved for reconstruction.

RULE-P-002 — Token Order Integrity
During parsing:
- <w:r> elements MUST NOT be merged.
- Token order MUST be preserved exactly for:
  <w:t>, <w:tab/>, <w:br/>, <w:cr/>, <w:sym/>
- Non-run or non-token elements MAY be ignored but MUST NOT affect token order.

RULE-P-003 — Whitespace Preservation (RAW text)
Text nodes MUST preserve exact string content including:
- leading spaces
- trailing spaces
- multiple consecutive spaces
No trimming or normalization is allowed.

RULE-P-004 — Preserve Flag Capture (informational)
If source <w:t> has xml:space="preserve", parser MAY set:
  run.meta.preserve = true
This flag is informational in v0.2 and is not required for reconstruction correctness (see RULE-R-006).

RULE-P-005 — Numbering Mapping
If paragraph contains:
  <w:numPr>
Then RAW MUST store:
  p_format.numbering = { numId: string, ilvl: integer }
If <w:ind w:hanging> exists, it MUST be stored per RULE-P-001.
numbering.xml MUST be mapped into:
  numbering_definitions[numId] = { abstractNumId?, levels, lvl_overrides? }
No synthetic defaults are allowed.

RULE-P-006 — Spacing Preservation (no synthesis)
If paragraph effective spacing is explicitly defined in OOXML (in paragraph pPr, style pPr, or docDefaults pPr):
- Parser MUST preserve it in RAW p_format as the corresponding fields.
- Zero values MUST NOT be dropped (use “is not None”, not truthy checks).
Supported spacing fields in RAW p_format:
  spaceBeforeTwip, spaceAfterTwip,
  spaceBeforeLines, spaceAfterLines,
  beforeAutospacing, afterAutospacing,
  lineTwip, lineRule
If a field is not present in OOXML, it MUST NOT be synthesized into RAW.

RULE-P-007 — Default Base Style Mapping
Parser MUST set `meta.default_style_id` to the internal RAW style_id that corresponds to Word default paragraph style (`w:style` with `w:type="paragraph"` and `w:default="1"/"true"`).
The mapping MUST be based on effective formatting (docDefaults + style chain + that style's own pPr/rPr), then registered through the RAW styles library.
This field does not bind content style_id values to Word `styleId`; it only selects the RAW base style used for generating Word Normal style.


3. Reconstruction Rules (RAW → DOCX)

RULE-R-001 — Deterministic Reconstruction
Reconstruction MUST NOT:
- normalize whitespace
- auto-correct formatting
- reorder tokens
- infer formatting not present in RAW
It MUST follow RAW strictly within current supported subset.

RULE-R-002 — Paragraph Formatting Emission
For each content item:
- A <w:p> MUST be emitted.
- Paragraph properties MUST be emitted from styles[style_id].p_format.
- content[].p_override exists in schema but is NOT applied by UltimateReconstructorV10 (see RULE-R-OVERRIDE-NYI).

RULE-R-003 — Empty Paragraph Preservation
If RAW paragraph has:
  "runs": []
Then reconstruction MUST emit:
  <w:p> (empty paragraph)
Paragraph formatting MUST still be applied.

RULE-R-004 — Paragraph Mark Run Formatting for Empty Paragraphs
If paragraph is empty (runs == []):
- Reconstructor MUST emit paragraph-mark run properties inside <w:pPr>/<w:rPr>
  using styles[style_id].r_format (if any).

RULE-R-005 — Run Formatting Model (style-driven base)
For non-empty paragraphs:
- base run format MUST be taken from styles[style_id].r_format.
- effective run format = merge(base_r, run.diff) where run.diff overrides base_r per-key.
- Keys with value null/None MUST NOT be written into XML.

RULE-R-006 — Space Preservation (xml:space="preserve")
For run.type="text" (and run.type="sym" in this reconstructor):
- Reconstructor MUST set xml:space="preserve" on <w:t> if and only if the text:
  - starts with space, OR
  - ends with space, OR
  - contains two consecutive spaces
(run.meta.preserve is allowed by schema but is not required by UltimateReconstructorV10.)

RULE-R-007 — Token Emission
Supported run types in UltimateReconstructorV10:
- text  -> <w:t>
- tab   -> <w:tab/>
- break -> <w:br/> and optional @w:type from run.break_type if provided and valid
- cr    -> <w:cr/>
- sym   -> emitted as <w:t> with literal text (best-effort in this version)
Token order MUST match RAW order exactly.

RULE-R-008 — Leading Tab Hint (informational)
If RAW run:
  { "type": "tab", "meta": { "leading": true } }
Reconstructor emits <w:tab/> (same as any tab run).
This hint is informational in v0.2 (no tab-to-indent conversion exists).

RULE-R-OVERRIDE-NYI — Paragraph Overrides Not Yet Implemented
content[].p_override is reserved by schema.
UltimateReconstructorV10 does not apply p_override in this version.
Therefore, for deterministic reconstruction with V10, producers SHOULD omit p_override or keep it empty.

RULE-RUN-COMPATIBILITY — Run Type Coverage (Parser v41 + Reconstructor v10)
End-to-end parsed and reconstructed run types:
- text, tab, break, cr, sym
Parsed but not reconstructed:
- none in current supported set
Schema-allowed but not reconstructed:
- softHyphen, noBreakHyphen
Unsupported/non-target run content remains outside current deterministic scope.

RULE-RUN-UNSUPPORTED — Schema-Allowed but Not Reconstructed by V10
Run types allowed by schema but ignored by UltimateReconstructorV10:
- softHyphen, noBreakHyphen
If present, they do not contribute to reconstructed XML in this version.
For strict 1:1 determinism with V10, producers MUST NOT emit these run types.

RULE-R-009 — Package Parts
Reconstructor MUST write:
- word/document.xml
- word/styles.xml (minimal Normal + docDefaults run properties)
  - Reconstructor MUST use meta.default_style_id (if present and valid) to build default paragraph style Normal with pPr from styles[default_style_id].p_format and may include rPr from styles[default_style_id].r_format.
  - If meta.default_style_id is missing or invalid, current minimal fallback behavior is used.
- word/settings.xml (minimal settings.xml with defaultTabStop if present)
- word/numbering.xml only if numbering_definitions non-empty
And the required relationships and [Content_Types].xml deterministically.

4. Forbidden Transformations
Parser/reconstructor MUST NOT:
- merge runs during parsing
- convert tabs to spaces or spaces to tabs
- remove empty paragraphs
- synthesize missing formatting
- drop explicit false/0 values present in RAW
