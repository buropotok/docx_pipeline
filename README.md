TranslateFactory
1. Purpose

TranslateFactory — детерминированная система подготовки нотариально заверяемых переводов документов.

Система реализует воспроизводимый pipeline:

- Schema: 2.8.1
- Rules: 0.2
- Reconstructor: UltimateReconstructorV11
- Parser: (see src/parser.py version header)

Система проектируется как аудитопригодная и версионируемая.

2. Repository Structure
schema/                # JSON Schema (Canonical IR)
rules/                 # Contract + invariants
src/                   # parser / reconstructor
tf_dagster/            # orchestration layer
tools/                 # CLI utilities
data/runs/<run_id>/    # run artifacts
VERSION.md
README.md
3. Version Matrix (Strict Sync)

Все изменения синхронизируются по матрице:

Schema: 2.8.2

Rules: 0.2

Parser: v41

Reconstructor: v10

Изменение любого из компонентов требует:

Обновления VERSION.md

Проверки contract.md

Проверки совместимости parser/reconstructor

Несогласованность версий запрещена.

4. Canonical Pipeline
4.1 Stage 0 — Word Materialization

Asset:

saveas_materialized

Назначение:

Приведение входного DOCX к стабильному формату через Word COM SaveAs

Нормализация скрытых особенностей Word

Ограничения:

COM изолирован

parser не зависит от COM

4.2 Stage 1 — Deterministic Parsing

Asset:

parse_raw_json

Вход:

materialized.docx

Выход:

raw.json

Гарантии:

Pure

Без догадок

Без synthetic значений

Строго по schema

Полная детерминированность

4.3 Stage 2 — Cleaning (AI)

Артефакт:

raw.cleaned.json

Допустимые действия:

whitespace normalization

```bash
python src/reconstructor.py

---

## Windows pipeline (Word materialize -> parse -> enrich -> reconstruct)

Requirements:

```bash
pip install -r requirements.txt
```

Run:

```bash
python -m tools.run_pipeline donor.docx  # reads ./data/donor.docx by default
```

Note: Windows only, with installed Microsoft Word.

By default pipeline artifacts are written into `/data`:
- `/data/<name>.materialized.docx`
- `/data/<name>.json`
- `/data/<name>.effective.json`
- `/data/<name>.reconstructed.docx`

Raw ZIP contents are also extracted automatically for analysis:
- `/data/raw/donor`
- `/data/raw/materialized`
- `/data/raw/reconstructed`

Each run writes a diagnostic log:
- `/data/logs/run_YYYYMMDD_HHMMSS.log`



## Official deterministic pipeline (Windows target flow)

0) Word SaveAs materialization: `donor.docx -> donor.materialized.docx`  
1) XML Parser: `donor.materialized.docx -> donor.json`  
2) Effective materializer (Word COM enrichment): `donor.materialized.docx + donor.json -> donor.effective.json`  
   - fill holes only, no overwrite of already parsed values  
3) Reconstructor: `donor.effective.json -> donor.reconstructed.docx`

For target documents, visual 1:1 geometry guarantee applies to reconstructor input `donor.effective.json`.
SaveAs and enrichment are mandatory parts of the official pipeline for cases where source OOXML does not serialize effective Word defaults completely.

