# AGENTS.md
Project: DOCX → RAW JSON → DOCX Deterministic Pipeline
Status: Active
See VERSION.md for current synchronized versions.


---

## 1. Project Goal

This repository implements a deterministic DOCX ↔ RAW JSON pipeline.

Primary objective:

    Visually 1:1 deterministic reconstruction
    for a constrained “forms” subset of DOCX.

RAW JSON is the only Intermediate Representation (IR).

The system must remain:
- deterministic
- schema-driven
- evolution-safe
- free of synthetic formatting

No visual optimization is allowed.

---

## 2. Core Architecture

Pipeline:

    DOCX → Word SaveAs (materialize) → Parser → Effective Materializer → RAW JSON → Reconstructor → DOCX

Key components:

- src/parser.py
- src/reconstructor.py
- schema/raw.schema.json
- rules/contract.md

All four MUST remain synchronized.

---

## 3. Source of Truth Hierarchy

Priority order:

1. raw.schema.json
2. rules/contract.md
3. parser.py / reconstructor.py

If a mismatch occurs:
- Schema and Rules define expected behavior.
- Code must be updated to match them.
- Schema must NOT be reduced to match implementation gaps.

---

## 4. Hard Constraints (Non-Negotiable)

### 4.1 No Synthetic Defaults
Parser and reconstructor MUST NOT invent values
that are not present in source OOXML or RAW.

Zero values (0) MUST NOT be dropped.

Use explicit `is not None` checks.

---

### 4.2 No Destructive Refactoring
Agents MUST:
- Apply minimal, surgical patches.
- Avoid rewriting working subsystems.
- Preserve backward compatibility.

No “rewrite from scratch”.

---

### 4.3 Determinism
The system MUST NOT:

- normalize whitespace
- merge runs during parsing
- reorder tokens
- auto-correct formatting
- infer missing formatting

Output must be reproducible byte-stable (except non-critical ordering in zip container).

---

## 5. Spacing Rules

When working with paragraph spacing:

- Preserve:
  - spaceBeforeTwip
  - spaceAfterTwip
  - spaceBeforeLines
  - spaceAfterLines
  - beforeAutospacing
  - afterAutospacing
  - lineTwip
  - lineRule

- Do NOT synthesize defaults (e.g. “8pt after” UI defaults).
- Respect explicit zeros.

---

## 6. Run Formatting Model

For non-empty paragraphs:

    base_r = styles[style_id].r_format
    effective_r = merge(base_r, run.diff)

For empty paragraphs:

    base_r applied to paragraph mark (<w:pPr>/<w:rPr>)

No implicit inheritance between runs.

---

## 7. Unsupported (Current Scope)

Do NOT implement unless version is bumped:

- Tables
- Images
- Shapes
- Fields
- Headers/Footers
- Footnotes
- Multi-section logic beyond terminal sectPr
- Advanced style inheritance

Run types allowed by schema but NOT reconstructed:
- softHyphen
- noBreakHyphen

Run type supported end-to-end (parse + reconstruct):
- cr

If implementing them:
- bump Schema + Rules version
- update this file

---

## 8. When Modifying Code

Agent MUST:

1. Explain what is broken.
2. Identify precise failure location.
3. Provide minimal patch.
4. Ensure no regression of existing rules.
5. Update version numbers if behavior changes.

Do NOT remove supported fields from schema.

---

## 9. Versioning Rules

If behavior changes:
- Update rules/contract.md version
- Update raw.schema.json version
- Update this file

Version bumps:

- Patch-level: internal bug fix
- Minor: new supported feature within scope
- Major: scope expansion (tables, headers, etc.)

---

## 10. Testing Expectations

Before finalizing a patch:

- RAW must validate against schema.
- No spacing fields lost.
- No runs merged during parsing.
- Empty paragraphs preserved.
- xml:space behavior deterministic.

---

## 11. Communication Style for Agents

When proposing changes:

Use:

    "In the code, ..."
    "The code does ..."
    "This patch modifies ..."

Avoid:

    "You hardcoded ..."
    "Your code is wrong ..."

Focus on system behavior, not authorship.

---

## 12. Evolution Strategy

All evolution must be:

- incremental
- versioned
- reversible
- deterministic

No breaking changes without explicit version bump.

---

END OF AGENTS MANIFEST
