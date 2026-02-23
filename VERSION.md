# VERSION MATRIX

Current synchronized versions:

Schema: 2.8.3
Rules: 0.2
Parser: v42
Reconstructor: v11

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
