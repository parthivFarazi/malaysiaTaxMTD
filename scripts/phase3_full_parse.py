#!/usr/bin/env python3
"""Phase 3 full-document parsed grid generation and validation.

This script runs the accepted geometry parser across the full PDF, using the
Phase 1 page classification result. It produces parsed-grid and validation
artifacts only. It does not normalize deduction values, generate Canonical_Long,
or export Excel.
"""

from __future__ import annotations

import argparse
import csv
from collections import Counter
from pathlib import Path
from typing import Any

import pdfplumber

from phase2_parser_prototype import (
    DATA_ROW_PATTERN,
    DEFAULT_CLASSIFICATION_PATH,
    extract_words,
    group_words_into_lines,
    load_classification,
    locate_pdf,
    output_grid_columns,
    parse_page,
    token_assignment_row,
    write_csv,
)


DEFAULT_OUTPUT_DIR = "phase3_full_parse_output"

EXPECTED_ROWS_BY_FAMILY = {
    "intro_data_ka1_ka10": 28,
    "normal_ka1_ka10": 47,
    "short_final_ka1_ka10": 17,
    "cover_ka11_ka20": 0,
    "intro_data_ka11_ka20": 24,
    "normal_ka11_ka20": 44,
    "short_final_ka11_ka20": 32,
    "non_data_or_unknown": 0,
}

DATA_FAMILIES = {
    "intro_data_ka1_ka10",
    "normal_ka1_ka10",
    "short_final_ka1_ka10",
    "intro_data_ka11_ka20",
    "normal_ka11_ka20",
    "short_final_ka11_ka20",
}

ZERO_RECORD_FAMILIES = {
    "cover_ka11_ka20",
    "non_data_or_unknown",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run Phase 3 full-document parsed grid generation and validation."
    )
    parser.add_argument(
        "--project-root",
        type=Path,
        default=Path.cwd(),
        help="Project folder containing the source PDF. Defaults to current directory.",
    )
    parser.add_argument(
        "--pdf",
        type=Path,
        default=None,
        help="Optional explicit PDF path. If omitted, exactly one PDF must exist under project root.",
    )
    parser.add_argument(
        "--classification-csv",
        type=Path,
        default=None,
        help=f"Page classification CSV. Defaults to {DEFAULT_CLASSIFICATION_PATH}.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help=f"Output directory. Defaults to {DEFAULT_OUTPUT_DIR}/ under project root.",
    )
    return parser.parse_args()


def section_from_classification(row: dict[str, str]) -> str:
    section = row.get("detected_section_code", "")
    if section:
        return section
    family = row.get("detected_family", "")
    if "ka1_ka10" in family:
        return "KA1_KA10"
    if "ka11_ka20" in family:
        return "KA11_KA20"
    return ""


def data_like_row_count(page: Any) -> tuple[int, int, list[dict[str, Any]]]:
    words = extract_words(page)
    lines = group_words_into_lines(words)
    count = sum(1 for line in lines if DATA_ROW_PATTERN.match(line.text))
    token_rows = [
        token_assignment_row(
            page_number=page.page_number,
            word=word,
            assignment_type="non_data_token",
            row=None,
            column_name="",
            notes="zero-record page",
        )
        for word in words
    ]
    return count, len(words), token_rows


def exception_row(
    exception_type: str,
    severity: str,
    page_number: int,
    details: str,
    row_index: str = "",
    column_name: str = "",
    token_index: str = "",
    text: str = "",
) -> dict[str, Any]:
    return {
        "exception_type": exception_type,
        "severity": severity,
        "page_number": page_number,
        "row_index": row_index,
        "column_name": column_name,
        "token_index": token_index,
        "text": text,
        "details": details,
    }


def summarize_page_numbers(page_numbers: list[int]) -> str:
    if not page_numbers:
        return "none"
    ranges: list[str] = []
    start = previous = page_numbers[0]
    for page_number in page_numbers[1:]:
        if page_number == previous + 1:
            previous = page_number
            continue
        ranges.append(str(start) if start == previous else f"{start}-{previous}")
        start = previous = page_number
    ranges.append(str(start) if start == previous else f"{start}-{previous}")
    return ",".join(ranges)


def write_summary_markdown(
    path: Path,
    pdf_path: Path,
    validation_rows: list[dict[str, Any]],
    exception_rows: list[dict[str, Any]],
) -> None:
    family_counts = Counter(row["detected_family"] for row in validation_rows)
    status_counts = Counter(row["page_status"] for row in validation_rows)
    unexpected_row_pages = [
        int(row["page_number"])
        for row in validation_rows
        if str(row["row_count_status"]) != "ok"
    ]
    low_pages = [
        int(row["page_number"])
        for row in validation_rows
        if int(row["unassigned_data_like_tokens"]) > 0
        or int(row["duplicate_cells"]) > 0
        or int(row["grid_review_rows"]) > 0
    ]
    total_rows = sum(int(row["parsed_rows"]) for row in validation_rows)
    total_assigned = sum(int(row["assigned_cell_tokens"]) for row in validation_rows)
    total_unassigned = sum(int(row["unassigned_data_like_tokens"]) for row in validation_rows)
    total_duplicates = sum(int(row["duplicate_cells"]) for row in validation_rows)

    lines = [
        "# Phase 3 Full Parse Summary",
        "",
        f"PDF: `{pdf_path.name}`",
        "",
        "This run generated a full Parsed_Grid dataset only. It did not normalize "
        "deduction values, generate Canonical_Long, or export Excel.",
        "",
        "## Totals",
        "",
        f"- Pages processed: `{len(validation_rows)}`",
        f"- Total parsed rows: `{total_rows}`",
        f"- Total assigned cell values: `{total_assigned}`",
        f"- Unassigned data-like tokens: `{total_unassigned}`",
        f"- Duplicate cells: `{total_duplicates}`",
        f"- Exception count: `{len(exception_rows)}`",
        f"- Pages with unexpected row counts: `{summarize_page_numbers(unexpected_row_pages)}`",
        f"- Pages with cell assignment issues: `{summarize_page_numbers(low_pages)}`",
        "",
        "## Pages By Family",
        "",
        "| Family | Pages |",
        "| --- | ---: |",
    ]
    for family in EXPECTED_ROWS_BY_FAMILY:
        lines.append(f"| `{family}` | {family_counts[family]} |")

    lines.extend(
        [
            "",
            "## Page Status",
            "",
            "| Status | Pages |",
            "| --- | ---: |",
        ]
    )
    for status, count in sorted(status_counts.items()):
        lines.append(f"| `{status}` | {count} |")

    promotion_safe = (
        len(exception_rows) == 0
        and total_unassigned == 0
        and total_duplicates == 0
        and not unexpected_row_pages
    )
    lines.extend(
        [
            "",
            "## Promotion Assessment",
            "",
            (
                "The parsed-grid parser is safe to promote from prototype to production-parser "
                "implementation."
                if promotion_safe
                else "The parsed-grid parser is not safe to promote until exceptions are resolved."
            ),
        ]
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    project_root = args.project_root.expanduser().resolve()
    pdf_path = locate_pdf(project_root, args.pdf)
    classification_path = (
        args.classification_csv.expanduser().resolve()
        if args.classification_csv is not None
        else project_root / DEFAULT_CLASSIFICATION_PATH
    )
    output_dir = (
        args.output_dir.expanduser().resolve()
        if args.output_dir is not None
        else project_root / DEFAULT_OUTPUT_DIR
    )

    if not classification_path.exists():
        raise FileNotFoundError(
            f"Missing classification CSV: {classification_path}. "
            "Run scripts/phase1_classify_pages.py first."
        )

    classification = load_classification(classification_path)
    parsed_grid_rows: list[dict[str, Any]] = []
    cell_assignment_rows: list[dict[str, Any]] = []
    validation_rows: list[dict[str, Any]] = []
    exceptions: list[dict[str, Any]] = []

    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)
        missing_classification_pages = [
            page_number
            for page_number in range(1, page_count + 1)
            if page_number not in classification
        ]
        for page_number in missing_classification_pages:
            exceptions.append(
                exception_row(
                    "missing_page_classification",
                    "critical",
                    page_number,
                    "Page is missing from Phase 1 page_classification.csv",
                )
            )

        for page in pdf.pages:
            page_number = page.page_number
            class_row = classification.get(page_number)
            if class_row is None:
                validation_rows.append(
                    {
                        "page_number": page_number,
                        "detected_family": "",
                        "section_code": "",
                        "expected_rows": "",
                        "observed_data_like_rows": 0,
                        "parsed_rows": 0,
                        "assigned_cell_tokens": 0,
                        "unassigned_data_like_tokens": 0,
                        "duplicate_cells": 0,
                        "grid_review_rows": 0,
                        "word_count": 0,
                        "row_count_status": "exception",
                        "page_status": "exception",
                        "notes": "missing classification",
                    }
                )
                continue

            family = class_row["detected_family"]
            section_code = section_from_classification(class_row)
            expected_rows = EXPECTED_ROWS_BY_FAMILY.get(family)
            page_exceptions_before = len(exceptions)

            if expected_rows is None:
                exceptions.append(
                    exception_row(
                        "unexpected_template_family",
                        "critical",
                        page_number,
                        f"Unexpected detected_family from classification: {family}",
                    )
                )
                expected_rows = 0

            if family in ZERO_RECORD_FAMILIES:
                observed_rows, word_count, page_token_rows = data_like_row_count(page)
                cell_assignment_rows.extend(page_token_rows)
                parsed_rows = 0
                assigned_cell_tokens = 0
                unassigned_tokens = observed_rows
                duplicate_cells = 0
                grid_review_rows = 0
                page_status = "zero_record_logged"
                notes = "cover/non-data page explicitly logged"
                if observed_rows != expected_rows:
                    exceptions.append(
                        exception_row(
                            "zero_record_page_has_data_like_rows",
                            "critical",
                            page_number,
                            (
                                f"Expected {expected_rows} data-like rows for {family}, "
                                f"observed {observed_rows}"
                            ),
                        )
                    )
            elif family in DATA_FAMILIES:
                try:
                    (
                        page_grid,
                        _page_columns,
                        _page_rows,
                        page_tokens,
                        page_summary,
                    ) = parse_page(page, family, section_code)
                    parsed_grid_rows.extend(page_grid)
                    cell_assignment_rows.extend(page_tokens)
                    parsed_rows = int(page_summary["row_count"])
                    observed_rows = parsed_rows
                    assigned_cell_tokens = int(page_summary["assigned_cell_tokens"])
                    unassigned_tokens = int(page_summary["unassigned_data_tokens"])
                    duplicate_cells = int(page_summary["duplicate_cells"])
                    grid_review_rows = int(page_summary["grid_review_rows"])
                    word_count = int(page_summary["word_count"])
                    page_status = "parsed"
                    notes = ""

                    if unassigned_tokens:
                        exceptions.append(
                            exception_row(
                                "unassigned_data_like_tokens",
                                "critical",
                                page_number,
                                f"{unassigned_tokens} data-like tokens were not assigned",
                            )
                        )
                    if duplicate_cells:
                        exceptions.append(
                            exception_row(
                                "duplicate_cell_assignments",
                                "critical",
                                page_number,
                                f"{duplicate_cells} cells had multiple token assignments",
                            )
                        )
                    if grid_review_rows:
                        exceptions.append(
                            exception_row(
                                "grid_review_rows",
                                "major",
                                page_number,
                                f"{grid_review_rows} parsed rows require review",
                            )
                        )

                    for grid_row in page_grid:
                        if not grid_row.get("salary_from") or not grid_row.get("salary_to"):
                            exceptions.append(
                                exception_row(
                                    "missing_salary_range",
                                    "critical",
                                    page_number,
                                    "Parsed row missing salary_from or salary_to",
                                    row_index=str(grid_row.get("row_index", "")),
                                )
                            )
                except Exception as exc:  # noqa: BLE001 - full validation records parser failures.
                    parsed_rows = 0
                    observed_rows = 0
                    assigned_cell_tokens = 0
                    unassigned_tokens = 0
                    duplicate_cells = 0
                    grid_review_rows = 0
                    word_count = 0
                    page_status = "exception"
                    notes = repr(exc)
                    exceptions.append(
                        exception_row(
                            "page_parse_failure",
                            "critical",
                            page_number,
                            repr(exc),
                        )
                    )
            else:
                observed_rows, word_count, page_token_rows = data_like_row_count(page)
                cell_assignment_rows.extend(page_token_rows)
                parsed_rows = 0
                assigned_cell_tokens = 0
                unassigned_tokens = observed_rows
                duplicate_cells = 0
                grid_review_rows = 0
                page_status = "exception"
                notes = f"parsed row outside expected page families: {family}"
                exceptions.append(
                    exception_row(
                        "parsed_row_outside_expected_family",
                        "critical",
                        page_number,
                        f"Unsupported page family: {family}",
                    )
                )

            if observed_rows != expected_rows:
                exceptions.append(
                    exception_row(
                        "unexpected_row_count",
                        "critical",
                        page_number,
                        f"Expected {expected_rows} rows for {family}, observed {observed_rows}",
                    )
                )
                row_count_status = "exception"
            else:
                row_count_status = "ok"

            page_exception_count = len(exceptions) - page_exceptions_before
            validation_rows.append(
                {
                    "page_number": page_number,
                    "detected_family": family,
                    "section_code": section_code,
                    "expected_rows": expected_rows,
                    "observed_data_like_rows": observed_rows,
                    "parsed_rows": parsed_rows,
                    "assigned_cell_tokens": assigned_cell_tokens,
                    "unassigned_data_like_tokens": unassigned_tokens,
                    "duplicate_cells": duplicate_cells,
                    "grid_review_rows": grid_review_rows,
                    "word_count": word_count,
                    "row_count_status": row_count_status,
                    "page_status": page_status,
                    "page_exception_count": page_exception_count,
                    "notes": notes,
                }
            )

    output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(output_dir / "parsed_grid_full.csv", parsed_grid_rows, output_grid_columns())
    write_csv(
        output_dir / "cell_assignment_report_full.csv",
        cell_assignment_rows,
        [
            "page_number",
            "token_index",
            "text",
            "x0",
            "x1",
            "top",
            "bottom",
            "x_center",
            "y_center",
            "assignment_type",
            "row_index",
            "column_name",
            "notes",
        ],
    )
    write_csv(
        output_dir / "parse_validation_summary.csv",
        validation_rows,
        [
            "page_number",
            "detected_family",
            "section_code",
            "expected_rows",
            "observed_data_like_rows",
            "parsed_rows",
            "assigned_cell_tokens",
            "unassigned_data_like_tokens",
            "duplicate_cells",
            "grid_review_rows",
            "word_count",
            "row_count_status",
            "page_status",
            "page_exception_count",
            "notes",
        ],
    )
    write_csv(
        output_dir / "parse_exceptions.csv",
        exceptions,
        [
            "exception_type",
            "severity",
            "page_number",
            "row_index",
            "column_name",
            "token_index",
            "text",
            "details",
        ],
    )
    write_summary_markdown(
        output_dir / "full_parse_summary.md",
        pdf_path,
        validation_rows,
        exceptions,
    )

    total_rows = sum(int(row["parsed_rows"]) for row in validation_rows)
    total_assigned = sum(int(row["assigned_cell_tokens"]) for row in validation_rows)
    total_unassigned = sum(int(row["unassigned_data_like_tokens"]) for row in validation_rows)
    total_duplicates = sum(int(row["duplicate_cells"]) for row in validation_rows)
    unexpected_row_pages = [
        int(row["page_number"])
        for row in validation_rows
        if row["row_count_status"] != "ok"
    ]

    print(f"PDF: {pdf_path}")
    print(f"Pages processed: {len(validation_rows)}")
    print(f"Total parsed rows: {total_rows}")
    print(f"Total assigned cell values: {total_assigned}")
    print(f"Unassigned data-like tokens: {total_unassigned}")
    print(f"Duplicate cells: {total_duplicates}")
    print(f"Exception count: {len(exceptions)}")
    print(f"Pages with unexpected row counts: {summarize_page_numbers(unexpected_row_pages)}")
    print(f"Output: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
