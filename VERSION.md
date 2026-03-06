# VERSION MATRIX

Current synchronized versions:

Schema: 2.16
Rules: 0.4
Parser: v43
Reconstructor: v2.15

---

# Synchronization Contract

These four components MUST always remain synchronized:

- schema/raw.schema.json
- rules/contract.md
- src/parser.py
- src/reconstructor.py

If any behavior changes:

- Update corresponding version.
- Update AGENTS.md.
- Update this file.
