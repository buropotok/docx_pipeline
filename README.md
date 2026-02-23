TranslateFactory
1. Purpose

TranslateFactory — детерминированная система подготовки нотариально заверяемых переводов документов.

Система реализует воспроизводимый pipeline:

DOCX / Scan
    → Canonical IR
    → Fingerprints
    → Knowledge Base alignment
    → Filled IR
    → Deterministic Reconstruction
    → Final DOCX

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

удаление технического мусора

структурная очистка

Запрещено:

генерация новых semantic значений

исправление логики документа

4.4 Stage 3 — Organizing (AI)

Артефакт:

raw.cleaned.organized.json

Добавляется:

SG (structural groups)

семантические блоки

логическая структура

4.5 Stage 4 — Docx Fingerprint

Артефакт:

docx_fingerprint.json

Содержит:

структурные сигнатуры

признаки шаблонности

repeat patterns

Используется для:

KB matching

clustering

winner selection

4.6 Stage 5 — Filled Stage

Артефакты:

raw.cleaned.organized.filled.json
docx_fingerprint.filled.json

Инварианты:

fill holes only

no overwrite

no synthetic

no normalization outside contract

После этого выполняется нормализация обратно в:

raw.cleaned.organized.json
docx_fingerprint.json
4.7 Stage 6 — Reconstruction

Asset:

reconstruct_docx

Вход:

organized.json

docx_fingerprint.json

Выход:

reconstructed.docx

Гарантии:

Строго детерминированный

Fail-fast

Не исправляет вход

Не генерирует missing

5. Orchestration (Dagster)

Pipeline реализован через Dagster assets.

Основные assets
saveas_materialized
parse_raw_json
materialize_effective
reconstruct_docx

Full job:

full_run_job

Используется:

Nothing

non_argument_deps

file-based run_dir

6. Run Model

Каждый запуск создаёт:

data/runs/<run_id>/

Типовая структура:

input.docx
materialized.docx
unpacked_docx/
raw.json
raw.cleaned.json
raw.cleaned.organized.json
docx_fingerprint.json
raw.cleaned.organized.filled.json
docx_fingerprint.filled.json
reconstructed.docx
logs/

Все артефакты сохраняются.

Pipeline допускает перезапуск с любого промежуточного этапа.

7. How to Run
7.1 Запуск через Dagster

Пример:

dagster job execute -f tf_dagster/assets.py -j full_run_job

(или через dagster dev UI)

Параметризация input осуществляется через конфигурацию job / env.

7.2 Ручной запуск parser
python src/parser.py input.docx output_raw.json

(если предусмотрен CLI)

7.3 Ручной запуск reconstructor
python src/reconstructor.py organized.json fingerprint.json output.docx
8. Determinism Guarantees

Система обеспечивает:

Полную воспроизводимость

Отсутствие скрытых мутаций

Явные промежуточные состояния

Отсутствие implicit state

Версионирование схем

Любой run можно:

повторить

сравнить

диффировать

аудитировать

9. AI Boundary

AI используется только в:

cleaning

organizing

winner selection

filling

edits

AI никогда не используется в:

parsing

reconstruction

fingerprint, если запрещено контрактом

10. Engineering Constraints

Запрещено:

синтетика вне разрешённого

автоматическое исправление missing

silent fallback

implicit defaults

merging runs

Все изменения схем — только эволюционные.

11. Roadmap (Engineering)

Parameterized input

Batch KB bootstrap

COM concurrency limits

Quarantine mechanism

Retry policy

Full online pipeline

Production notarization mode