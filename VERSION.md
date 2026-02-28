# VERSION MATRIX

Current synchronized versions:

Schema: 2.12
Rules: 0.3
Parser: v43
Reconstructor: v12

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
