# DOCX Deterministic Pipeline

Deterministic DOCX ↔ RAW JSON pipeline for "forms" subset documents.

Current Versions:

- Schema: 2.8.1
- Rules: 0.2
- Reconstructor: UltimateReconstructorV10
- Parser: (see src/parser.py version header)

---

# 1. Project Goal

This project implements a deterministic transformation pipeline:

    DOCX → RAW JSON → DOCX

The objective is:

    Visually 1:1 reconstruction
    for structured, certificate-like, form-style documents.

RAW JSON is the only Intermediate Representation (IR).

No visual optimization.
No synthetic defaults.
No heuristic normalization.

Determinism is mandatory.

---

# 2. Repository Structure
docx_pipeline/
│
├── AGENTS.md
├── README.md
├── VERSION.md
│
├── schema/
│ └── raw.schema.json
│
├── rules/
│ └── contract.md
│
├── src/
│ ├── parser.py
│ └── reconstructor.py
│
└── tests/


---

# 3. Core Architecture

## 3.1 Parsing

The parser:

- Reads DOCX using lxml
- Extracts structural information
- Produces RAW JSON compliant with schema
- Preserves:
  - spacing
  - numbering
  - indents
  - run order
  - whitespace

No merging of runs during parsing.

---

## 3.2 Reconstruction

The reconstructor:

- Builds DOCX using lxml (NO python-docx)
- Uses RAW JSON as strict source of truth
- Applies style-driven run model:
  
  base_r = styles[style_id].r_format  
  effective_r = merge(base_r, run.diff)

- Preserves xml:space rules deterministically
- Emits minimal required OpenXML structure

---

# 4. Deterministic Rules

See:
rules/contract.md

Key invariants:

- No synthetic spacing
- Zero values must not be dropped
- No token reordering
- No run merging during parsing
- Empty paragraphs preserved
- No implicit formatting inference

---

# 5. Supported Features (v0.2 Scope)

✔ Paragraph formatting  
✔ Indents (including hanging)  
✔ Spacing (twip + lines + autospacing)  
✔ Numbering with overrides  
✔ Tabs  
✔ Basic run formatting  
✔ Page setup (sectPr)  
✔ defaultTabStop  

---

# 6. Out of Scope

Not supported in this version:

- Tables
- Images
- Shapes
- Fields
- Hyperlinks
- Headers/Footers
- Footnotes
- Multi-section documents beyond final sectPr

Schema may define more fields than currently reconstructed.
Implementation must catch up — schema must not be reduced.

---

# 7. How to Run

## Reconstruct from RAW JSON

```bash
python src/reconstructor.py
