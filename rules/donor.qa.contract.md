# donor.qa.contract.md

## QA Pipeline Contract for Donor Preparation

### 1. Purpose

The QA pipeline validates that the **optimized reconstruction of a document preserves the original layout geometry sufficiently to qualify the document as a donor candidate**.

The pipeline performs a **layout comparison between the original document and the optimized reconstructed document** using Word's rendering engine via Word COM automation.

The QA stage acts as a **quality gate** before documents are admitted to the donor pool.

---

# 2. When QA Runs

QA is executed **after corpus optimization**.

Trigger conditions:

```
document.processing_status = 'optimized'
```

QA may run:

* on the entire corpus
* on a selected subset
* incrementally on newly optimized documents

---

# 3. QA Pipeline Overview

Pipeline stages:

```
optimized document
    ↓
Word COM layout audit
    ↓
layout audit JSON artifacts
    ↓
AI evaluation
    ↓
QA metrics stored in DB
    ↓
processing_status update
```

Outputs:

```
original.layout_audit.json
optimized.layout_audit.json
qa_ai_report.json
database QA metrics
updated document.processing_status
```

---

# 4. Document Inputs

For each document the pipeline retrieves from DB:

Required fields:

```
document.id
document.source_abs_path
document.artifacts_abs_path
document.processing_status
```

Required artifacts:

```
original DOCX
optimized reconstructed DOCX
```

If optimized reconstructed DOCX does not exist, the QA pipeline must mark the document as:

```
optimized_fail
```

---

# 5. Word Layout Audit Stage

Word COM is used as the **layout oracle**.

The pipeline must run Word COM in **batch mode**.

Recommended pattern:

```
start Word instance
process N documents
close Word
```

Typical batch size:

```
100–300 documents
```

---

## 5.1 Layout Metrics Extraction

For both documents:

```
original.docx
optimized.docx
```

The audit script extracts layout metrics using Word COM.

### Document-level metrics

```
page_count
paragraph_count
table_count
```

### Paragraph-level metrics

For each paragraph:

```
index
page
top_pt
bottom_pt
height_pt
(optional) text snippet
```

### Table-level metrics

For each table:

```
index
page
rows
columns
top_pt
bottom_pt
height_pt
(optional) text snippet
```

Floating tables may produce unreliable height values.
Downstream paragraph shifts should be used to detect floating object layout drift.

---

# 6. Layout Audit Artifacts

For each document two JSON artifacts must be generated.

### Original document audit

```
original.layout_audit.json
```

### Optimized document audit

```
optimized.layout_audit.json
```

Artifacts must be stored inside the document artifacts directory.

Recommended location:

```
artifacts/uid_xxxxx/qa/
```

Example:

```
qa/original.layout_audit.json
qa/optimized.layout_audit.json
```

---

# 7. AI Analysis Stage

The pipeline sends both audit JSON files to AI.

Inputs:

```
original.layout_audit.json
optimized.layout_audit.json
```

The AI evaluates:

* structural consistency
* layout drift
* table geometry changes
* paragraph shifts
* page count differences

AI must return **strict JSON** according to the contract below.

---

# 8. AI Response Contract

Expected AI response format:

```json
{
  "status": "ok",
  "score": 0.93,
  "confidence": 0.9,
  "decision": "optimized_check",
  "summary": "Layout preserved with minor acceptable drift.",
  "issues": [
    {
      "type": "table_height_delta",
      "severity": "low",
      "location": {
        "table_index": 2
      },
      "details": "Table height changed by 10pt."
    }
  ],
  "metrics": {
    "page_count_changed": false,
    "max_paragraph_top_delta_pt": 3.0,
    "max_table_height_delta_pt": 10.0
  }
}
```

---

# 9. Deterministic Validation Layer

AI evaluation must **not override deterministic failure conditions**.

The pipeline must perform the following hard checks:

### Fatal conditions

```
original DOCX missing
optimized DOCX missing
Word COM failure
layout audit JSON missing
invalid AI response JSON
```

If any fatal condition occurs:

```
processing_status = optimized_fail
```

---

# 10. QA Status Assignment

After deterministic checks and AI evaluation the document receives a new status.

### optimized_check

Document passed QA and is accepted as a donor candidate.

Conditions:

```
AI decision == optimized_check
AND no deterministic failures
```

---

### optimized_warning

Document contains layout deviations requiring manual review.

Examples:

```
table height change above threshold
paragraph drift detected
AI confidence low
AI decision == optimized_warning
```

Documents with this status are candidates for **manual verification**.

---

### optimized_fail

QA pipeline failed or document geometry is clearly broken.

Examples:

```
missing artifacts
Word COM error
AI response invalid
major layout corruption
```

---

# 11. QA Database Table

A dedicated QA table must store metrics and results.

Suggested name:

```
document_layout_qa
```

---

## 11.1 Core Fields

```
id
document_id
qa_status
qa_started_at
qa_finished_at
qa_duration_ms
```

---

## 11.2 Artifact Paths

```
original_audit_abs_path
optimized_audit_abs_path
ai_report_abs_path
raw_ai_response_abs_path
```

---

## 11.3 Layout Metrics

```
page_count_original
page_count_optimized

paragraph_count_original
paragraph_count_optimized

table_count_original
table_count_optimized

max_paragraph_top_delta_pt
max_paragraph_height_delta_pt
max_table_top_delta_pt
max_table_height_delta_pt
```

---

## 11.4 AI Metrics

```
ai_status
ai_score
ai_confidence
ai_summary
ai_decision
model_name
prompt_version
```

---

## 11.5 Error Handling

```
error_message
warning_count
fatal_count
```

---

# 12. Artifact Storage

All QA artifacts must be saved alongside document artifacts.

Recommended structure:

```
artifacts/
  uid_xxxxx/
    qa/
      original.layout_audit.json
      optimized.layout_audit.json
      qa_ai_report.json
      qa_log.json
```

---

# 13. Donor Pool Admission Rule

Only documents with status:

```
processing_status = optimized_check
```

may enter the donor retrieval pool.

Documents with:

```
optimized_warning
```

require manual review.

Documents with:

```
optimized_fail
```

must be excluded from the donor pool.

---

# 14. Performance Requirements

QA must operate in **batch mode**.

Guidelines:

```
Word instance reused within batch
batch size: 100–300 documents
Word restarted between batches
```

Typical processing time:

```
~0.5–1.5 seconds per document
```

---

# 15. Design Principles

The QA system follows these principles:

1. Word rendering is treated as the **ground truth layout engine**.
2. Deterministic checks guard against runtime failures.
3. AI provides **semantic interpretation of layout differences**.
4. All QA decisions are **reproducible through stored artifacts**.
5. Only verified documents enter the donor pool.

---

# 16. Final Objective

The QA pipeline ensures that:

```
optimized donor documents preserve layout integrity
before being used in document reconstruction workflows.
```

This step prevents corrupted or geometrically unstable donors from entering the system.
