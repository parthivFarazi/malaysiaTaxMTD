# Malaysian MTD 2018 PDF → Audit-Ready Excel Conversion Plan (Revision 2)

## Project Overview

Convert the official Malaysian 2018 Monthly Tax Deduction (MTD/PCB) PDF into a structured, searchable, and auditable Excel workbook.

The final dataset must be suitable for business use where every value can be traced back to the official government source document.

The source PDF contains approximately 733 pages and follows a highly structured repeating table format.

The objective is not merely PDF extraction.

The objective is:

1. Extract accurately
2. Map values to the correct columns
3. Validate every value
4. Minimize manual review
5. Preserve auditability

---

# Key Discovery

After inspecting the actual PDF:

- The PDF is machine-readable.
- The PDF is not OCR-based.
- The table structure is highly repetitive.
- The primary risk is column misalignment.
- The primary challenge is parsing blank cells correctly.

Example structure:

```text
Salary Range

Category 1
B

Category 2
K
KA1
KA2
...
KA10

Category 3
K
KA1
KA2
...
KA10
```

This structure repeats throughout the document.

Because of this:

- OCR is unnecessary.
- Image-based extraction is unnecessary.
- AI interpretation is unnecessary.
- Complex multi-engine extraction is unnecessary.

The solution should focus on deterministic parsing.

---

# Success Criteria

## Accuracy

Target:

100% verified final dataset

Every final value must satisfy one of:

- Parsed successfully and passed validation
- Reviewed manually and approved

No value may be silently inferred.

---

# High-Level Architecture

Official PDF
↓
Native Text Extraction
↓
Table Parser
↓
Column State Engine
↓
Validation Engine
↓
Exception Generation
↓
Manual Review
↓
Excel Export
↓
Audit Package

---

# Folder Structure

```text
mtd_2018_conversion/

├── data/
│   ├── official_pdf/
│   │   └── Jadual_PCB_2018.pdf
│   │
│   ├── raw_text/
│   ├── parsed_rows/
│   ├── validation/
│   └── final_excel/
│
├── scripts/
│   ├── 01_extract_text.py
│   ├── 02_parse_rows.py
│   ├── 03_map_columns.py
│   ├── 04_validate_structure.py
│   ├── 05_validate_business_rules.py
│   ├── 06_generate_exceptions.py
│   ├── 07_export_excel.py
│   └── 08_generate_audit_report.py
│
├── logs/
│   ├── parsing_log.csv
│   ├── exception_log.csv
│   ├── review_log.csv
│   └── validation_log.csv
│
└── README.md
```

---

# Phase 1 — Native Text Extraction

## Tool

pdfplumber

Alternative:

pdfminer.six

Goal:

Extract raw rows exactly as they appear in the PDF.

Example:

```text
3141 - 3145 1.00 - - - - - - - - - - - 1.00 - - - - - - - - - -
```

No interpretation.

No transformation.

Store raw extraction.

---

# Phase 2 — Column State Engine

## Why This Exists

This is the most important component in the system.

The PDF contains many blank cells represented by "-".

If blanks are handled incorrectly:

```text
Column 4
↓
moves into
↓
Column 3
```

This creates silent corruption.

The Column State Engine prevents this.

---

## Fixed Schema

Each row must map into:

```text
salary_from
salary_to

B

cat2_K
cat2_KA1
cat2_KA2
cat2_KA3
cat2_KA4
cat2_KA5
cat2_KA6
cat2_KA7
cat2_KA8
cat2_KA9
cat2_KA10

cat3_K
cat3_KA1
cat3_KA2
cat3_KA3
cat3_KA4
cat3_KA5
cat3_KA6
cat3_KA7
cat3_KA8
cat3_KA9
cat3_KA10
```

Total:

24 business columns plus salary range fields.

Every row must conform.

---

# Phase 3 — Parser

Convert raw text rows into structured records.

Example:

Raw:

```text
3661 - 3665 16.60 11.60 6.60 1.60
```

Parsed:

```json
{
  "salary_from": 3661,
  "salary_to": 3665,
  "B": 16.60,
  "cat3_K": 11.60,
  "cat3_KA1": 6.60,
  "cat3_KA2": 1.60
}
```

All missing values stored as NULL.

Never shifted.

Never guessed.

---

# Phase 4 — Structural Validation

## Column Count Validation

Expected:

```text
24 tax columns
```

Any deviation:

```text
EXCEPTION
```

---

## Salary Range Validation

Expected:

```text
3141-3145
3146-3150
3151-3155
```

Check continuity.

Missing range:

```text
EXCEPTION
```

---

## Duplicate Range Validation

Same salary range appearing twice:

```text
EXCEPTION
```

---

# Phase 5 — Business Rule Validation

## Numeric Validation

Allowed:

```text
0
positive numbers
NULL
```

Not allowed:

```text
negative values
text
corrupted values
```

---

## Pattern Validation

Detect sudden anomalies.

Example:

```text
16.60
16.75
16.90
17.05
999.99
17.35
```

Flag.

---

## Category Progression Validation

Many columns increase gradually.

Unexpected jumps should be reviewed.

Not automatically corrected.

Only flagged.

---

# Phase 6 — Row Fingerprinting

Every parsed row generates:

```text
SHA256(
salary_from|
salary_to|
all columns
)
```

Purpose:

- Detect accidental modifications
- Detect export corruption
- Detect manual edits

---

# Phase 7 — Exception Engine

Generate exception records for:

## Parsing Failure

Example:

```text
Unable to map row
```

---

## Missing Columns

Example:

```text
Expected 24
Found 23
```

---

## Duplicate Salary Ranges

Example:

```text
3661-3665 appears twice
```

---

## Validation Failures

Example:

```text
Unexpected numeric anomaly
```

---

# Phase 8 — Manual Review

Goal:

Review only exceptions.

Never review entire PDF manually.

Each exception should contain:

```text
Page Number
Salary Range
Column Name
Raw Row
Parsed Row
Failure Reason
```

Reviewer actions:

```text
Approve
Correct
Escalate
```

Every action logged.

---

# Expected Manual Review Volume

Estimated:

```text
Total rows:
30,000 - 40,000

Manual review:
< 100 rows
```

if parser and validations are implemented correctly.

---

# Excel Workbook Structure

## Sheet 1

Raw_Extract

Exact extracted text.

---

## Sheet 2

Parsed_Data

Fully structured dataset.

---

## Sheet 3

Validation_Log

Validation results.

---

## Sheet 4

Exceptions

Rows requiring review.

---

## Sheet 5

Manual_Review_Log

Reviewer decisions.

---

## Sheet 6

Final_MTD

Approved production dataset.

---

## Sheet 7

README

Methodology and audit notes.

---

# Audit Requirements

Every final value must be traceable to:

```text
Excel Cell
↓
Parsed Row
↓
Raw Text Row
↓
PDF Page
↓
Official PDF
```

---

# Non-Negotiable Rules

1. No silent corrections.
2. No AI-generated values.
3. No column shifting.
4. No inferred tax amounts.
5. No automatic fixes.
6. All exceptions logged.
7. All manual corrections logged.
8. Official PDF remains the source of truth.
9. Every final value must be reproducible from the source document.