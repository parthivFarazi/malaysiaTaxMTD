#!/usr/bin/env python3
"""Phase 4 canonical long dataset generation and validation.

This script converts the full Parsed_Grid into Canonical_Long and validates the
canonical dataset. It does not generate Excel or any final user-facing workbook.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any


DEFAULT_PARSED_GRID = "phase3_full_parse_output/parsed_grid_full.csv"
DEFAULT_METADATA = "phase1_discovery_output/pdf_metadata.json"
DEFAULT_OUTPUT_DIR = "phase4_canonical_output"
SCRIPT_VERSION = "phase4-canonicalize-0.1.0"

CANONICAL_FIELDS = [
    "record_id",
    "source_pdf_sha256",
    "source_filename",
    "source_page",
    "template_family",
    "section_code",
    "source_row_index",
    "source_column_label",
    "salary_from",
    "salary_to",
    "category_group",
    "spouse_working_status",
    "dependent_code",
    "dependent_number",
    "raw_value",
    "normalized_value",
    "blank_status",
    "validation_status",
]

EXCEPTION_FIELDS = [
    "exception_id",
    "severity",
    "validation_layer",
    "rule_name",
    "record_id",
    "source_page",
    "source_row_index",
    "source_column_label",
    "category_group",
    "dependent_code",
    "raw_value",
    "details",
]

SECTION_ALLOWED_COLUMNS = {
    "KA1_KA10": [
        "B",
        "cat2_K",
        *[f"cat2_KA{index}" for index in range(1, 11)],
        "cat3_K",
        *[f"cat3_KA{index}" for index in range(1, 11)],
    ],
    "KA11_KA20": [
        *[f"cat2_KA{index}" for index in range(11, 21)],
        *[f"cat3_KA{index}" for index in range(11, 21)],
    ],
}

VALID_SECTION_DEPENDENTS = {
    "KA1_KA10": {"B", "K", *[f"KA{index}" for index in range(1, 11)]},
    "KA11_KA20": {f"KA{index}" for index in range(11, 21)},
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate and validate Canonical_Long from the full parsed grid."
    )
    parser.add_argument(
        "--project-root",
        type=Path,
        default=Path.cwd(),
        help="Project folder. Defaults to current directory.",
    )
    parser.add_argument(
        "--parsed-grid",
        type=Path,
        default=None,
        help=f"Parsed grid CSV. Defaults to {DEFAULT_PARSED_GRID}.",
    )
    parser.add_argument(
        "--metadata",
        type=Path,
        default=None,
        help=f"PDF metadata JSON. Defaults to {DEFAULT_METADATA}.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help=f"Output directory. Defaults to {DEFAULT_OUTPUT_DIR}.",
    )
    return parser.parse_args()


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def load_metadata(path: Path) -> dict[str, Any]:
    with path.open(encoding="utf-8") as handle:
        metadata = json.load(handle)
    required = ["sha256", "filename"]
    missing = [key for key in required if not metadata.get(key)]
    if missing:
        raise RuntimeError(f"Missing required metadata keys in {path}: {missing}")
    return metadata


def column_mapping(column_label: str) -> dict[str, Any]:
    if column_label == "B":
        return {
            "category_group": "category_1",
            "spouse_working_status": "single",
            "dependent_code": "B",
            "dependent_number": "",
        }

    match = re.fullmatch(r"cat([23])_(K|KA\d+)", column_label)
    if not match:
        raise ValueError(f"Unsupported source column label: {column_label}")

    category_number = match.group(1)
    dependent_code = match.group(2)
    return {
        "category_group": f"category_{category_number}",
        "spouse_working_status": (
            "spouse_not_working" if category_number == "2" else "spouse_working"
        ),
        "dependent_code": dependent_code,
        "dependent_number": dependent_number(dependent_code),
    }


def dependent_number(dependent_code: str) -> str:
    if dependent_code.startswith("KA"):
        return dependent_code[2:]
    return ""


def normalize_raw_value(raw_value: str) -> tuple[str, str, str]:
    if raw_value == "-":
        return "", "observed_dash", ""
    if raw_value == "":
        return "", "observed_blank", "blank raw value in parsed grid"
    try:
        normalized = Decimal(raw_value)
    except InvalidOperation:
        return "", "malformed", "malformed numeric value"
    return format(normalized, "f"), "populated", ""


def make_exception(
    exceptions: list[dict[str, Any]],
    severity: str,
    validation_layer: str,
    rule_name: str,
    record: dict[str, Any],
    details: str,
) -> None:
    exceptions.append(
        {
            "exception_id": f"EXC-{len(exceptions) + 1:07d}",
            "severity": severity,
            "validation_layer": validation_layer,
            "rule_name": rule_name,
            "record_id": record.get("record_id", ""),
            "source_page": record.get("source_page", ""),
            "source_row_index": record.get("source_row_index", ""),
            "source_column_label": record.get("source_column_label", ""),
            "category_group": record.get("category_group", ""),
            "dependent_code": record.get("dependent_code", ""),
            "raw_value": record.get("raw_value", ""),
            "details": details,
        }
    )


def build_canonical_records(
    parsed_grid_path: Path, metadata: dict[str, Any]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    canonical: list[dict[str, Any]] = []
    exceptions: list[dict[str, Any]] = []
    provenance_rows: list[dict[str, Any]] = []
    source_pdf_sha256 = metadata["sha256"]
    source_filename = metadata["filename"]

    with parsed_grid_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        for parsed_row in reader:
            section_code = parsed_row["section_code"]
            allowed_columns = SECTION_ALLOWED_COLUMNS.get(section_code)
            if allowed_columns is None:
                placeholder = {
                    "source_page": parsed_row.get("page_number", ""),
                    "source_row_index": parsed_row.get("row_index", ""),
                }
                make_exception(
                    exceptions,
                    "critical",
                    "structural",
                    "unexpected_section_code",
                    placeholder,
                    f"Unexpected section_code: {section_code}",
                )
                continue

            provenance_rows.append(
                {
                    "source_page": parsed_row["page_number"],
                    "template_family": parsed_row["template_family"],
                    "section_code": section_code,
                    "source_row_index": parsed_row["row_index"],
                    "salary_from": parsed_row["salary_from"],
                    "salary_to": parsed_row["salary_to"],
                    "source_column_count": len(allowed_columns),
                    "canonical_record_count": len(allowed_columns),
                }
            )

            for column_label in allowed_columns:
                raw_value = parsed_row.get(column_label, "")
                mapping = column_mapping(column_label)
                normalized_value, blank_status, normalize_error = normalize_raw_value(raw_value)
                record_id = f"CL-{len(canonical) + 1:09d}"
                record = {
                    "record_id": record_id,
                    "source_pdf_sha256": source_pdf_sha256,
                    "source_filename": source_filename,
                    "source_page": parsed_row["page_number"],
                    "template_family": parsed_row["template_family"],
                    "section_code": section_code,
                    "source_row_index": parsed_row["row_index"],
                    "source_column_label": column_label,
                    "salary_from": parsed_row["salary_from"],
                    "salary_to": parsed_row["salary_to"],
                    "category_group": mapping["category_group"],
                    "spouse_working_status": mapping["spouse_working_status"],
                    "dependent_code": mapping["dependent_code"],
                    "dependent_number": mapping["dependent_number"],
                    "raw_value": raw_value,
                    "normalized_value": normalized_value,
                    "blank_status": blank_status,
                    "validation_status": "valid",
                }
                canonical.append(record)

                if normalize_error:
                    make_exception(
                        exceptions,
                        "critical",
                        "business",
                        "malformed_numeric_value",
                        record,
                        normalize_error,
                    )

    return canonical, exceptions, provenance_rows


def validate_structural(
    canonical: list[dict[str, Any]], exceptions: list[dict[str, Any]]
) -> None:
    seen_keys: set[tuple[str, str, str, str]] = set()
    required_fields = [
        "record_id",
        "source_pdf_sha256",
        "source_filename",
        "source_page",
        "template_family",
        "section_code",
        "source_row_index",
        "source_column_label",
        "salary_from",
        "salary_to",
        "category_group",
        "spouse_working_status",
        "dependent_code",
        "raw_value",
        "blank_status",
    ]
    valid_category_pairs = {
        ("category_1", "single"),
        ("category_2", "spouse_not_working"),
        ("category_3", "spouse_working"),
    }

    for record in canonical:
        missing = [field for field in required_fields if record.get(field) == ""]
        if missing:
            make_exception(
                exceptions,
                "critical",
                "structural",
                "missing_required_metadata",
                record,
                f"Missing required fields: {', '.join(missing)}",
            )

        try:
            salary_from = Decimal(record["salary_from"])
            salary_to = Decimal(record["salary_to"])
            if salary_from > salary_to:
                make_exception(
                    exceptions,
                    "critical",
                    "structural",
                    "salary_from_gt_salary_to",
                    record,
                    "salary_from is greater than salary_to",
                )
        except InvalidOperation:
            make_exception(
                exceptions,
                "critical",
                "structural",
                "malformed_salary_range",
                record,
                "salary_from or salary_to is not numeric",
            )

        allowed_dependents = VALID_SECTION_DEPENDENTS.get(record["section_code"], set())
        if record["dependent_code"] not in allowed_dependents:
            make_exception(
                exceptions,
                "critical",
                "structural",
                "unexpected_dependent_code_for_section",
                record,
                (
                    f"{record['dependent_code']} is not expected in section "
                    f"{record['section_code']}"
                ),
            )

        category_pair = (record["category_group"], record["spouse_working_status"])
        if category_pair not in valid_category_pairs:
            make_exception(
                exceptions,
                "critical",
                "structural",
                "invalid_category_mapping",
                record,
                f"Invalid category/spouse mapping: {category_pair}",
            )
        if record["category_group"] == "category_1" and record["dependent_code"] != "B":
            make_exception(
                exceptions,
                "critical",
                "structural",
                "invalid_category_1_dependent",
                record,
                "Category 1 must use dependent_code B",
            )

        unique_key = (
            record["source_page"],
            record["source_row_index"],
            record["source_column_label"],
            record["section_code"],
        )
        if unique_key in seen_keys:
            make_exception(
                exceptions,
                "critical",
                "structural",
                "duplicate_canonical_record",
                record,
                f"Duplicate provenance key: {unique_key}",
            )
        seen_keys.add(unique_key)


def validate_business(
    canonical: list[dict[str, Any]], exceptions: list[dict[str, Any]]
) -> None:
    for record in canonical:
        if record["blank_status"] != "populated":
            continue
        try:
            value = Decimal(record["normalized_value"])
        except InvalidOperation:
            make_exception(
                exceptions,
                "warning",
                "business",
                "malformed_numeric_value",
                record,
                "normalized_value could not be parsed as Decimal",
            )
            continue
        if value < 0:
            make_exception(
                exceptions,
                "warning",
                "business",
                "negative_deduction_value",
                record,
                "Deduction value is negative",
            )

    series: dict[tuple[str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for record in canonical:
        series[
            (
                record["category_group"],
                record["spouse_working_status"],
                record["dependent_code"],
            )
        ].append(record)

    for records in series.values():
        ordered = sorted(
            records,
            key=lambda item: (
                Decimal(item["salary_from"]),
                Decimal(item["salary_to"]),
                int(item["source_page"]),
                int(item["source_row_index"]),
            ),
        )
        previous_salary_from: Decimal | None = None
        previous_salary_to: Decimal | None = None
        previous_value: Decimal | None = None
        positive_deltas: list[Decimal] = []
        value_pairs: list[tuple[dict[str, Any], Decimal, Decimal]] = []

        for record in ordered:
            salary_from = Decimal(record["salary_from"])
            salary_to = Decimal(record["salary_to"])
            if previous_salary_from is not None and salary_from < previous_salary_from:
                make_exception(
                    exceptions,
                    "warning",
                    "business",
                    "unexpected_salary_range_ordering",
                    record,
                    "Salary ranges are not sorted within dependent-code series",
                )
            if previous_salary_to is not None and salary_from <= previous_salary_to:
                make_exception(
                    exceptions,
                    "warning",
                    "business",
                    "overlapping_salary_range",
                    record,
                    "Salary range overlaps prior range within dependent-code series",
                )
            previous_salary_from = salary_from
            previous_salary_to = salary_to

            if record["blank_status"] != "populated":
                continue
            value = Decimal(record["normalized_value"])
            if previous_value is not None:
                delta = value - previous_value
                value_pairs.append((record, value, delta))
                if delta < 0:
                    make_exception(
                        exceptions,
                        "warning",
                        "business",
                        "suspicious_deduction_decrease",
                        record,
                        f"Deduction decreased by {abs(delta)} from previous populated value",
                    )
                elif delta > 0:
                    positive_deltas.append(delta)
            previous_value = value

        if len(positive_deltas) < 10:
            continue
        sorted_deltas = sorted(positive_deltas)
        median_delta = sorted_deltas[len(sorted_deltas) // 2]
        extreme_threshold = max(Decimal("1000"), median_delta * Decimal("25"))
        for record, _value, delta in value_pairs:
            if delta > extreme_threshold:
                make_exception(
                    exceptions,
                    "warning",
                    "business",
                    "extreme_deduction_jump",
                    record,
                    f"Deduction jump {delta} exceeds threshold {extreme_threshold}",
                )


def apply_validation_status(
    canonical: list[dict[str, Any]], exceptions: list[dict[str, Any]]
) -> None:
    status_by_record: dict[str, str] = {}
    for exception in exceptions:
        record_id = exception.get("record_id", "")
        if not record_id:
            continue
        if exception["severity"] == "critical":
            status_by_record[record_id] = "failed"
        elif status_by_record.get(record_id) != "failed":
            status_by_record[record_id] = "warning"

    for record in canonical:
        record["validation_status"] = status_by_record.get(record["record_id"], "valid")


def validation_summary_rows(
    canonical: list[dict[str, Any]], exceptions: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    exception_counter = Counter((ex["validation_layer"], ex["rule_name"], ex["severity"]) for ex in exceptions)
    status_counter = Counter(record["validation_status"] for record in canonical)
    blank_counter = Counter(record["blank_status"] for record in canonical)
    rows = [
        {
            "validation_name": "total_canonical_records",
            "validation_layer": "summary",
            "severity": "",
            "status": "info",
            "checked_records": len(canonical),
            "issue_count": "",
            "details": "Total records in Canonical_Long",
        },
        {
            "validation_name": "validation_status_valid",
            "validation_layer": "summary",
            "severity": "",
            "status": "info",
            "checked_records": len(canonical),
            "issue_count": status_counter["valid"],
            "details": "Records with validation_status=valid",
        },
        {
            "validation_name": "validation_status_warning",
            "validation_layer": "summary",
            "severity": "warning",
            "status": "pass" if status_counter["warning"] == 0 else "warning",
            "checked_records": len(canonical),
            "issue_count": status_counter["warning"],
            "details": "Records with validation_status=warning",
        },
        {
            "validation_name": "validation_status_failed",
            "validation_layer": "summary",
            "severity": "critical",
            "status": "pass" if status_counter["failed"] == 0 else "fail",
            "checked_records": len(canonical),
            "issue_count": status_counter["failed"],
            "details": "Records with validation_status=failed",
        },
    ]

    for blank_status, count in sorted(blank_counter.items()):
        rows.append(
            {
                "validation_name": f"blank_status_{blank_status}",
                "validation_layer": "summary",
                "severity": "",
                "status": "info",
                "checked_records": len(canonical),
                "issue_count": count,
                "details": f"Records with blank_status={blank_status}",
            }
        )

    expected_rules = [
        ("structural", "salary_from_gt_salary_to", "critical"),
        ("structural", "unexpected_dependent_code_for_section", "critical"),
        ("structural", "invalid_category_mapping", "critical"),
        ("structural", "duplicate_canonical_record", "critical"),
        ("structural", "missing_required_metadata", "critical"),
        ("business", "negative_deduction_value", "warning"),
        ("business", "malformed_numeric_value", "warning"),
        ("business", "unexpected_salary_range_ordering", "warning"),
        ("business", "suspicious_deduction_decrease", "warning"),
        ("business", "extreme_deduction_jump", "warning"),
    ]
    for layer, rule_name, severity in expected_rules:
        issue_count = exception_counter[(layer, rule_name, severity)]
        rows.append(
            {
                "validation_name": rule_name,
                "validation_layer": layer,
                "severity": severity,
                "status": "pass" if issue_count == 0 else ("fail" if severity == "critical" else "warning"),
                "checked_records": len(canonical),
                "issue_count": issue_count,
                "details": f"{rule_name} issues found",
            }
        )
    return rows


def provenance_summary_rows(
    canonical: list[dict[str, Any]], parsed_row_provenance: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    by_page: dict[tuple[str, str, str], dict[str, int]] = defaultdict(lambda: Counter())
    for record in canonical:
        key = (record["source_page"], record["template_family"], record["section_code"])
        by_page[key]["canonical_records"] += 1
        if record["blank_status"] == "observed_dash":
            by_page[key]["observed_dash_records"] += 1
        elif record["blank_status"] == "populated":
            by_page[key]["populated_records"] += 1

    parsed_rows_by_page = Counter(row["source_page"] for row in parsed_row_provenance)
    rows: list[dict[str, Any]] = []
    for key in sorted(by_page, key=lambda item: int(item[0])):
        source_page, template_family, section_code = key
        counts = by_page[key]
        rows.append(
            {
                "source_page": source_page,
                "template_family": template_family,
                "section_code": section_code,
                "parsed_grid_rows": parsed_rows_by_page[source_page],
                "canonical_records": counts["canonical_records"],
                "populated_records": counts["populated_records"],
                "observed_dash_records": counts["observed_dash_records"],
                "provenance_path": "PDF Page -> Parsed Grid Row -> Column Label -> Canonical Record",
            }
        )
    return rows


def write_phase4_summary(
    path: Path,
    canonical: list[dict[str, Any]],
    exceptions: list[dict[str, Any]],
    metadata: dict[str, Any],
) -> None:
    blank_counter = Counter(record["blank_status"] for record in canonical)
    status_counter = Counter(record["validation_status"] for record in canonical)
    severity_counter = Counter(exception["severity"] for exception in exceptions)
    populated = blank_counter["populated"]
    observed_dash = blank_counter["observed_dash"]
    validation_failures = severity_counter["critical"]
    business_warnings = sum(
        1
        for exception in exceptions
        if exception["validation_layer"] == "business" and exception["severity"] == "warning"
    )
    production_ready = validation_failures == 0

    lines = [
        "# Phase 4 Canonical Dataset Summary",
        "",
        f"Source file: `{metadata['filename']}`",
        f"Source SHA-256: `{metadata['sha256']}`",
        "",
        "This phase generated Canonical_Long from Parsed_Grid. It did not export Excel.",
        "",
        "## Totals",
        "",
        f"- Total canonical records: `{len(canonical)}`",
        f"- Populated deduction records: `{populated}`",
        f"- Observed dash records: `{observed_dash}`",
        f"- Validation failures: `{validation_failures}`",
        f"- Business-rule warnings: `{business_warnings}`",
        f"- Exception count: `{len(exceptions)}`",
        f"- Records valid: `{status_counter['valid']}`",
        f"- Records warning: `{status_counter['warning']}`",
        f"- Records failed: `{status_counter['failed']}`",
        "",
        "## Readiness",
        "",
        (
            "Canonical_Long is production-ready as the source-of-truth dataset for the "
            "next Excel/reporting phase."
            if production_ready
            else "Canonical_Long is not production-ready until critical validation failures are resolved."
        ),
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    project_root = args.project_root.expanduser().resolve()
    parsed_grid_path = (
        args.parsed_grid.expanduser().resolve()
        if args.parsed_grid is not None
        else project_root / DEFAULT_PARSED_GRID
    )
    metadata_path = (
        args.metadata.expanduser().resolve()
        if args.metadata is not None
        else project_root / DEFAULT_METADATA
    )
    output_dir = (
        args.output_dir.expanduser().resolve()
        if args.output_dir is not None
        else project_root / DEFAULT_OUTPUT_DIR
    )

    if not parsed_grid_path.exists():
        raise FileNotFoundError(f"Missing parsed grid CSV: {parsed_grid_path}")
    if not metadata_path.exists():
        raise FileNotFoundError(f"Missing metadata JSON: {metadata_path}")

    metadata = load_metadata(metadata_path)
    canonical, exceptions, parsed_row_provenance = build_canonical_records(
        parsed_grid_path, metadata
    )
    validate_structural(canonical, exceptions)
    validate_business(canonical, exceptions)
    apply_validation_status(canonical, exceptions)

    validation_summary = validation_summary_rows(canonical, exceptions)
    provenance_summary = provenance_summary_rows(canonical, parsed_row_provenance)

    output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(output_dir / "canonical_long.csv", canonical, CANONICAL_FIELDS)
    write_csv(
        output_dir / "canonical_validation_summary.csv",
        validation_summary,
        [
            "validation_name",
            "validation_layer",
            "severity",
            "status",
            "checked_records",
            "issue_count",
            "details",
        ],
    )
    write_csv(output_dir / "canonical_exceptions.csv", exceptions, EXCEPTION_FIELDS)
    write_csv(
        output_dir / "provenance_summary.csv",
        provenance_summary,
        [
            "source_page",
            "template_family",
            "section_code",
            "parsed_grid_rows",
            "canonical_records",
            "populated_records",
            "observed_dash_records",
            "provenance_path",
        ],
    )
    write_phase4_summary(output_dir / "phase4_summary.md", canonical, exceptions, metadata)

    blank_counter = Counter(record["blank_status"] for record in canonical)
    severity_counter = Counter(exception["severity"] for exception in exceptions)
    business_warnings = sum(
        1
        for exception in exceptions
        if exception["validation_layer"] == "business" and exception["severity"] == "warning"
    )

    print(f"Canonical records: {len(canonical)}")
    print(f"Populated deduction records: {blank_counter['populated']}")
    print(f"Observed dash records: {blank_counter['observed_dash']}")
    print(f"Validation failures: {severity_counter['critical']}")
    print(f"Business-rule warnings: {business_warnings}")
    print(f"Exception count: {len(exceptions)}")
    print(f"Output: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
