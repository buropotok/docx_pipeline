Reconstruction Contract v0.2
(DOCX ↔ RAW JSON Core Rules, synced with UltimateReconstructorV11 + RAW JSON Schema v2.8.3)

0. Versioning
- RAW JSON MUST declare:
  meta.schema_version
  meta.rules_version
- This contract corresponds to:
  rules_version = "0.2"
  schema_version = "2.8.3"
  parser_version = "v42"
  reconstructor_version = "v11"

1. General Principles
1.1 Goal
- The goal of RAW is deterministic 1:1 visual reconstruction of a constrained DOCX subset (“forms”).
- RAW JSON is a structural IR. No visual optimization is allowed during parsing.
- Reconstruction must be deterministic and schema-driven.

1.2 No synthesis
- Parser and reconstructor MUST NOT invent formatting that is not present in OOXML or RAW.
- If a value is absent in OOXML, it MUST NOT appear in RAW as a synthesized default.
  (Future normalization/reconstruction policies may be introduced as opt-in modes.)

1.3 Pipeline Requirements
RULE-PIPE-001 — Materialization prerequisite
- Official deterministic pipeline for target documents is:
  0) Word SaveAs materialization: donor.docx -> donor.materialized.docx
  1) XML parsing: donor.materialized.docx -> donor.json
  2) Effective materializer enrichment: donor.materialized.docx + donor.json -> donor.effective.json
  3) Reconstruction: donor.effective.json -> reconstructed.docx
- Visual 1:1 guarantee for target documents applies to reconstruction input donor.effective.json.
- SaveAs + enrichment are required in the official pipeline for cases where source OOXML omits effective Word defaults.

RULE-E-001 — Enrichment semantics (fill holes only)
- Enrichment MAY add only missing values in RAW.
- Enrichment MUST NOT overwrite values already parsed into RAW.

RULE-E-002 — Enrichment proof requirement
- Enrichment MUST write only values deterministically extracted from the effective Word model.
- Enrichment MUST NOT introduce heuristic or guessed values.

RULE-R-VAL-001 — Reconstructor validation/fail-fast
- If required reconstruction value is missing, reconstructor MUST fail with a contract violation error.
- Reconstructor MUST NOT synthesize fallback values for missing required data.

1.4 Scope (v0.2)
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
- Run token types softHyphen and noBreakHyphen are allowed by schema but are NOT reconstructed by UltimateReconstructorV11 (see RULE-RUN-COMPATIBILITY and RULE-RUN-UNSUPPORTED)

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
  numbering_definitions[numId] = { abstractNumId?, multiLevelType?, levels, lvl_overrides? }
For each level, parser MUST preserve (if present in OOXML):
- levels[ilvl].tabPosTwip (from w:lvl/w:tab/@w:val)
- levels[ilvl].lvlJc
- levels[ilvl].suff
- levels[ilvl].pStyle
- levels[ilvl].level_pPr (ind + tabs subset)
- levels[ilvl].level_rPr (number style run properties)
No synthetic defaults are allowed: if a source element is absent, corresponding RAW fields MUST be absent.
In particular, if donor level does not contain w:lvl/w:tab, parser MUST NOT add tabPosTwip.
These fields are required to preserve list geometry (number/text indentation and number style).

RULE-P-INDENT-ORIGIN — Paragraph Indent Origin Tracking
When parser writes paragraph p_format indent fields from document.xml direct paragraph properties (<w:p>/<w:pPr>/<w:ind>):
- indentStartTwipOrigin = "direct" for indentStartTwip
- indentEndTwipOrigin = "direct" for indentEndTwip
- indentFirstLineTwipOrigin = "direct" for indentFirstLineTwip
- indentHangingTwipOrigin = "direct" for indentHangingTwip
Origin MUST be emitted only when corresponding indent value is present.
No synthetic origin defaults are allowed.

Reconstructor MUST suppress style-derived paragraph indents for numbered paragraphs:
- if paragraph has numbering and indent origin is not "direct" (or origin is absent), reconstructor MUST NOT emit corresponding <w:pPr>/<w:ind> attribute for that indent key.
- if numbered paragraph has at least one indent with origin "direct", only those direct indent attributes MAY be emitted.
This avoids paragraph-level indent override of numbering-level geometry.

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

RULE-P-STYLE-META-001 — Style Materialization Metadata (Stage 1)
RAW styles library entries MAY include the following OPTIONAL metadata fields:
  - styles[style_id].title: user-friendly name (e.g. "Обычный", "Стиль 1", ...)
  - styles[style_id].word_style_id: planned Word w:styleId for future styles.xml materialization (e.g. "TF_s0001")
  - styles[style_id].source_word_style_id: informational donor Word paragraph styleId observed while parsing (e.g. "Heading1")
Additionally, content items MAY include:
  - content[].source_word_style_id: donor Word paragraph styleId for that paragraph.
These fields MUST be assigned deterministically (no guessing) and MUST NOT affect formatting de-duplication (style identity remains based on p_format+r_format only).
UltimateReconstructorV11 may ignore these fields until the dedicated styles.xml materialization stage is implemented.

RULE-P-008 — Numbering Level Geometry (lvl pPr subset)
If numbering.xml level contains:
  <w:lvl ...><w:pPr>...</w:pPr></w:lvl>
Parser MUST extract RAW level_pPr subset when present:
  levels[ilvl].level_pPr.indentStartTwip   <- w:ind/@w:left
  levels[ilvl].level_pPr.indentEndTwip     <- w:ind/@w:right
  levels[ilvl].level_pPr.indentFirstLineTwip <- w:ind/@w:firstLine
  levels[ilvl].level_pPr.indentHangingTwip <- w:ind/@w:hanging
  levels[ilvl].level_pPr.tabs[] <- w:tabs/w:tab (posTwip/val/leader)
Fields MUST be written only when present/parseable. No defaults are allowed.
This geometry is required to preserve list number/text alignment.

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
- content[].p_override exists in schema but is NOT applied by UltimateReconstructorV11 (see RULE-R-OVERRIDE-NYI).

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
(run.meta.preserve is allowed by schema but is not required by UltimateReconstructorV11.)

RULE-R-007 — Token Emission
Supported run types in UltimateReconstructorV11:
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
UltimateReconstructorV11 does not apply p_override in this version.
Therefore, for deterministic reconstruction with V11, producers SHOULD omit p_override or keep it empty.

RULE-RUN-COMPATIBILITY — Run Type Coverage (Parser v42 + Reconstructor v11)
Parsed + reconstructed (end-to-end):
- text, tab, break, cr, sym
Parsed only (not reconstructed):
- softHyphen, noBreakHyphen
Unsupported in current scope:
- run token content outside schema-defined run.type set

RULE-RUN-UNSUPPORTED — Schema-Allowed but Not Reconstructed by V11
Run types allowed by schema but ignored by UltimateReconstructorV11:
- softHyphen, noBreakHyphen
If present, they do not contribute to reconstructed XML in this version.
For strict 1:1 determinism with V11, producers MUST NOT emit these run types.

RULE-R-009 — Package Parts
Reconstructor MUST write:
- word/document.xml
- word/styles.xml (minimal Normal + docDefaults run properties)
  - Reconstructor MUST use meta.default_style_id (if present and valid) to build default paragraph style Normal with pPr from styles[default_style_id].p_format and may include rPr from styles[default_style_id].r_format.
  - If meta.default_style_id is missing or invalid, current minimal fallback behavior is used.
- word/settings.xml (minimal settings.xml with defaultTabStop if present)
- word/numbering.xml only if numbering_definitions non-empty
And the required relationships and [Content_Types].xml deterministically.

RULE-R-010 — Numbering Level Geometry and Style Emission
If RAW numbering record contains corresponding fields, reconstructor MUST emit in numbering.xml:
- numbering_definitions[numId].multiLevelType -> <w:abstractNum><w:multiLevelType @w:val>
- levels[ilvl].tabPosTwip -> <w:tab @w:val> as direct child of <w:lvl>
- levels[ilvl].lvlJc -> <w:lvlJc @w:val>
- levels[ilvl].suff -> <w:suff @w:val>
- levels[ilvl].pStyle -> <w:pStyle @w:val>
- levels[ilvl].level_pPr -> <w:pPr> subset:
  - indentStartTwip -> w:ind/@w:left
  - indentEndTwip -> w:ind/@w:right
  - indentFirstLineTwip -> w:ind/@w:firstLine
  - indentHangingTwip -> w:ind/@w:hanging
  - tabs[] -> <w:tabs><w:tab .../></w:tabs>
- levels[ilvl].level_rPr -> <w:rPr> using RAW run-format model keys.
Missing fields MUST NOT be synthesized. If a field is absent in RAW, corresponding elements/attributes MUST NOT be emitted.
In particular, reconstructor MUST NOT emit <w:lvl><w:tab/></w:lvl> when tabPosTwip is absent in RAW.
This rule preserves list geometry (number/text indentation), multi-level behavior, and number style.

4. Forbidden Transformations
Parser/reconstructor MUST NOT:
- merge runs during parsing
- convert tabs to spaces or spaces to tabs
- remove empty paragraphs
- synthesize missing formatting
- drop explicit false/0 values present in RAW

