#!/usr/bin/env python3
"""Phase 5 independent audit verification.

This script verifies Canonical_Long against an independent pdfminer text-line
extraction path. It is read-only with respect to Canonical_Long and Parsed_Grid.
It does not generate Excel.
"""

from __future__ import annotations

import argparse
import csv
import random
import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from pdfminer.high_level import extract_pages
from pdfminer.layout import LAParams, LTTextContainer, LTTextLine


DEFAULT_CANONICAL = "phase4_canonical_output/canonical_long.csv"
DEFAULT_CLASSIFICATION = "phase1_discovery_output/page_classification.csv"
DEFAULT_OUTPUT_DIR = "phase5_independent_audit_output"
DEFAULT_RANDOM_SEED = 2018
DATA_ROW_PATTERN = re.compile(r"^\s*(\d+)\s*-\s*(\d+)\b")


KA1_KA10_COLUMNS = [
    "B",
    "cat2_K",
    *[f"cat2_KA{index}" for index in range(1, 11)],
    "cat3_K",
    *[f"cat3_KA{index}" for index in range(1, 11)],
]

KA11_KA20_COLUMNS = [
    *[f"cat2_KA{index}" for index in range(11, 21)],
    *[f"cat3_KA{index}" for index in range(11, 21)],
]


@dataclass(frozen=True)
class IndependentRecord:
    source_page: int
    source_row_index: int
    source_column_label: str
    salary_from: str
    salary_to: str
    dependent_code: str
    raw_value: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run Phase 5 independent audit verification against Canonical_Long."
    )
    parser.add_argument(
        "--project-root",
        type=Path,
        default=Path.cwd(),
        help="Project folder. Defaults to current directory.",
    )
    parser.add_argument(
        "--pdf",
        type=Path,
        default=None,
        help="Optional explicit PDF path. If omitted, exactly one PDF must exist under project root.",
    )
    parser.add_argument(
        "--canonical",
        type=Path,
        default=None,
        help=f"Canonical_Long CSV. Defaults to {DEFAULT_CANONICAL}.",
    )
    parser.add_argument(
        "--classification",
        type=Path,
        default=None,
        help=f"Page classification CSV. Defaults to {DEFAULT_CLASSIFICATION}.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help=f"Output directory. Defaults to {DEFAULT_OUTPUT_DIR}.",
    )
    parser.add_argument(
        "--seed",
        type=int,
        default=DEFAULT_RANDOM_SEED,
        help="Deterministic random seed for page sampling.",
    )
    return parser.parse_args()


def locate_pdf(project_root: Path, explicit_pdf: Path | None) -> Path:
    if explicit_pdf is not None:
        pdf_path = explicit_pdf.expanduser().resolve()
        if not pdf_path.exists():
            raise FileNotFoundError(f"PDF path does not exist: {pdf_path}")
        if pdf_path.suffix.lower() != ".pdf":
            raise ValueError(f"Expected a .pdf file, got: {pdf_path}")
        return pdf_path

    pdfs = sorted(
        path.resolve()
        for path in project_root.rglob("*.pdf")
        if not any(part.startswith(".") for part in path.relative_to(project_root).parts)
    )
    if not pdfs:
        raise FileNotFoundError(f"No PDF files found under project root: {project_root}")
    if len(pdfs) > 1:
        formatted = "\n".join(f"- {path}" for path in pdfs)
        raise RuntimeError(
            "Multiple PDF files found. Re-run with --pdf to avoid guessing:\n" + formatted
        )
    return pdfs[0]


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def load_classification(path: Path) -> dict[int, dict[str, str]]:
    with path.open(newline="", encoding="utf-8") as handle:
        return {int(row["page_number"]): row for row in csv.DictReader(handle)}


def expected_columns_for_section(section_code: str) -> list[str]:
    if section_code == "KA1_KA10":
        return KA1_KA10_COLUMNS
    if section_code == "KA11_KA20":
        return KA11_KA20_COLUMNS
    return []


def dependent_code_from_column(column_label: str) -> str:
    if column_label == "B":
        return "B"
    return column_label.split("_", 1)[1]


def grouped_pdfminer_lines_from_chunks(chunks: list[dict[str, Any]]) -> list[str]:
    groups: list[dict[str, Any]] = []
    for chunk in sorted(chunks, key=lambda item: (-item["y"], item["x"])):
        for group in groups:
            if abs(float(group["y"]) - float(chunk["y"])) <= 2.0:
                group["items"].append(chunk)
                group["ys"].append(chunk["y"])
                group["y"] = sum(group["ys"]) / len(group["ys"])
                break
        else:
            groups.append({"y": chunk["y"], "ys": [chunk["y"]], "items": [chunk]})

    return [
        " ".join(item["text"] for item in sorted(group["items"], key=lambda item: item["x"]))
        for group in groups
    ]


def grouped_pdfminer_lines_all(pdf_path: Path) -> dict[int, list[str]]:
    lines_by_page: dict[int, list[str]] = {}
    for page_number, page_layout in enumerate(
        extract_pages(str(pdf_path), laparams=LAParams()), start=1
    ):
        chunks: list[dict[str, Any]] = []
        for element in page_layout:
            if not isinstance(element, LTTextContainer):
                continue
            for obj in element:
                if isinstance(obj, LTTextLine):
                    text = obj.get_text().strip()
                    if text:
                        chunks.append({"x": float(obj.x0), "y": float(obj.y0), "text": text})
        lines_by_page[page_number] = grouped_pdfminer_lines_from_chunks(chunks)
    return lines_by_page


def grouped_pdfminer_lines(pdf_path: Path, page_number: int) -> list[str]:
    chunks: list[dict[str, Any]] = []
    for page_layout in extract_pages(
        str(pdf_path), page_numbers=[page_number - 1], laparams=LAParams()
    ):
        for element in page_layout:
            if not isinstance(element, LTTextContainer):
                continue
            for obj in element:
                if isinstance(obj, LTTextLine):
                    text = obj.get_text().strip()
                    if text:
                        chunks.append({"x": float(obj.x0), "y": float(obj.y0), "text": text})

    return grouped_pdfminer_lines_from_chunks(chunks)


def parse_independent_page(
    pdf_path: Path, page_number: int, section_code: str
) -> tuple[list[IndependentRecord], list[dict[str, Any]]]:
    return parse_independent_page_lines(
        grouped_pdfminer_lines(pdf_path, page_number), page_number, section_code
    )


def parse_independent_page_lines(
    lines: list[str], page_number: int, section_code: str
) -> tuple[list[IndependentRecord], list[dict[str, Any]]]:
    expected_columns = expected_columns_for_section(section_code)
    if not expected_columns:
        return [], []

    records: list[IndependentRecord] = []
    exceptions: list[dict[str, Any]] = []
    data_lines = [line for line in lines if DATA_ROW_PATTERN.match(line)]
    for row_index, line in enumerate(data_lines, start=1):
        tokens = line.split()
        if len(tokens) < 3 or tokens[1] != "-":
            exceptions.append(
                {
                    "mismatch_type": "independent_row_parse_failure",
                    "source_page": page_number,
                    "source_row_index": row_index,
                    "source_column_label": "",
                    "canonical_value": "",
                    "independent_value": line,
                    "details": "Salary row did not follow expected token shape",
                }
            )
            continue
        salary_from = tokens[0]
        salary_to = tokens[2]
        values = tokens[3:]
        if len(values) != len(expected_columns):
            exceptions.append(
                {
                    "mismatch_type": "independent_value_count_mismatch",
                    "source_page": page_number,
                    "source_row_index": row_index,
                    "source_column_label": "",
                    "canonical_value": len(expected_columns),
                    "independent_value": len(values),
                    "details": line,
                }
            )
            continue
        for column_label, raw_value in zip(expected_columns, values, strict=True):
            records.append(
                IndependentRecord(
                    source_page=page_number,
                    source_row_index=row_index,
                    source_column_label=column_label,
                    salary_from=salary_from,
                    salary_to=salary_to,
                    dependent_code=dependent_code_from_column(column_label),
                    raw_value=raw_value,
                )
            )
    return records, exceptions


def load_canonical(path: Path) -> tuple[
    dict[tuple[int, int, str], dict[str, str]],
    dict[int, list[dict[str, str]]],
    list[dict[str, str]],
]:
    by_key: dict[tuple[int, int, str], dict[str, str]] = {}
    by_page: dict[int, list[dict[str, str]]] = defaultdict(list)
    all_records: list[dict[str, str]] = []
    with path.open(newline="", encoding="utf-8") as handle:
        for record in csv.DictReader(handle):
            page = int(record["source_page"])
            row_index = int(record["source_row_index"])
            key = (page, row_index, record["source_column_label"])
            by_key[key] = record
            by_page[page].append(record)
            all_records.append(record)
    return by_key, by_page, all_records


def sample_pages(classification: dict[int, dict[str, str]], seed: int) -> tuple[list[int], list[int]]:
    ka1_pages = [
        page
        for page, row in classification.items()
        if row.get("detected_section_code") == "KA1_KA10"
        and row.get("detected_family") != "cover_ka11_ka20"
    ]
    ka11_pages = [
        page
        for page, row in classification.items()
        if row.get("detected_section_code") == "KA11_KA20"
        and row.get("detected_family") != "cover_ka11_ka20"
    ]
    rng = random.Random(seed)
    return sorted(rng.sample(ka1_pages, 50)), sorted(rng.sample(ka11_pages, 50))


def compare_sample_records(
    independent_records: list[IndependentRecord],
    canonical_by_key: dict[tuple[int, int, str], dict[str, str]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    results: list[dict[str, Any]] = []
    mismatches: list[dict[str, Any]] = []
    for independent in independent_records:
        key = (
            independent.source_page,
            independent.source_row_index,
            independent.source_column_label,
        )
        canonical = canonical_by_key.get(key)
        if canonical is None:
            result = result_row(independent, None, False, "missing_canonical_record")
            results.append(result)
            mismatches.append(mismatch_from_result(result))
            continue
        checks = {
            "salary_from": canonical["salary_from"] == independent.salary_from,
            "salary_to": canonical["salary_to"] == independent.salary_to,
            "dependent_code": canonical["dependent_code"] == independent.dependent_code,
            "raw_value": canonical["raw_value"] == independent.raw_value,
        }
        exact_match = all(checks.values())
        mismatch_reason = "|".join(name for name, passed in checks.items() if not passed)
        result = result_row(independent, canonical, exact_match, mismatch_reason)
        results.append(result)
        if not exact_match:
            mismatches.append(mismatch_from_result(result))
    return results, mismatches


def result_row(
    independent: IndependentRecord,
    canonical: dict[str, str] | None,
    exact_match: bool,
    mismatch_reason: str,
) -> dict[str, Any]:
    return {
        "source_page": independent.source_page,
        "source_row_index": independent.source_row_index,
        "source_column_label": independent.source_column_label,
        "dependent_code": independent.dependent_code,
        "canonical_salary_from": "" if canonical is None else canonical["salary_from"],
        "independent_salary_from": independent.salary_from,
        "canonical_salary_to": "" if canonical is None else canonical["salary_to"],
        "independent_salary_to": independent.salary_to,
        "canonical_raw_value": "" if canonical is None else canonical["raw_value"],
        "independent_raw_value": independent.raw_value,
        "canonical_blank_status": "" if canonical is None else canonical["blank_status"],
        "exact_match": exact_match,
        "mismatch_reason": mismatch_reason,
    }


def mismatch_from_result(result: dict[str, Any]) -> dict[str, Any]:
    return {
        "mismatch_scope": "sample_record",
        "mismatch_type": result["mismatch_reason"],
        "source_page": result["source_page"],
        "source_row_index": result["source_row_index"],
        "source_column_label": result["source_column_label"],
        "canonical_value": (
            f"{result['canonical_salary_from']}-{result['canonical_salary_to']}="
            f"{result['canonical_raw_value']}"
        ),
        "independent_value": (
            f"{result['independent_salary_from']}-{result['independent_salary_to']}="
            f"{result['independent_raw_value']}"
        ),
        "details": "Exact sample record comparison failed",
    }


def aggregate_canonical(records: list[dict[str, str]]) -> dict[str, Any]:
    by_section = Counter(record["section_code"] for record in records)
    by_column = Counter(
        (record["section_code"], record["source_column_label"]) for record in records
    )
    salary_ranges = {
        (record["section_code"], record["salary_from"], record["salary_to"])
        for record in records
    }
    parsed_rows = {
        (record["source_page"], record["source_row_index"]) for record in records
    }
    return {
        "record_count": len(records),
        "section_counts": by_section,
        "dependent_column_counts": by_column,
        "salary_range_count": len(salary_ranges),
        "parsed_row_count": len(parsed_rows),
    }


def aggregate_independent(records: list[IndependentRecord]) -> dict[str, Any]:
    by_section = Counter(
        "KA1_KA10" if record.source_column_label in KA1_KA10_COLUMNS else "KA11_KA20"
        for record in records
    )
    by_column = Counter(
        (
            "KA1_KA10" if record.source_column_label in KA1_KA10_COLUMNS else "KA11_KA20",
            record.source_column_label,
        )
        for record in records
    )
    salary_ranges = {
        (
            "KA1_KA10" if record.source_column_label in KA1_KA10_COLUMNS else "KA11_KA20",
            record.salary_from,
            record.salary_to,
        )
        for record in records
    }
    parsed_rows = {(record.source_page, record.source_row_index) for record in records}
    return {
        "record_count": len(records),
        "section_counts": by_section,
        "dependent_column_counts": by_column,
        "salary_range_count": len(salary_ranges),
        "parsed_row_count": len(parsed_rows),
    }


def audit_summary_rows(
    canonical_agg: dict[str, Any],
    independent_agg: dict[str, Any],
    sample_results: list[dict[str, Any]],
    sampled_pages: list[int],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []

    def add(metric: str, scope: str, canonical_count: int, independent_count: int, details: str) -> None:
        rows.append(
            {
                "metric": metric,
                "scope": scope,
                "canonical_count": canonical_count,
                "independent_count": independent_count,
                "status": "pass" if canonical_count == independent_count else "fail",
                "details": details,
            }
        )

    add(
        "record_count",
        "full_dataset",
        canonical_agg["record_count"],
        independent_agg["record_count"],
        "Total deduction records",
    )
    add(
        "parsed_row_count",
        "full_dataset",
        canonical_agg["parsed_row_count"],
        independent_agg["parsed_row_count"],
        "Unique PDF table rows",
    )
    add(
        "salary_range_count",
        "full_dataset",
        canonical_agg["salary_range_count"],
        independent_agg["salary_range_count"],
        "Distinct salary ranges by section",
    )
    for section in ["KA1_KA10", "KA11_KA20"]:
        add(
            "section_record_count",
            section,
            canonical_agg["section_counts"][section],
            independent_agg["section_counts"][section],
            "Records by section",
        )
    for key in sorted(canonical_agg["dependent_column_counts"]):
        add(
            "dependent_column_count",
            f"{key[0]}:{key[1]}",
            canonical_agg["dependent_column_counts"][key],
            independent_agg["dependent_column_counts"][key],
            "Records by section and dependent column",
        )

    exact_matches = sum(1 for row in sample_results if row["exact_match"])
    add(
        "sample_record_exact_matches",
        "random_sample",
        len(sample_results),
        exact_matches,
        f"Exact matches across {len(sampled_pages)} sampled pages",
    )
    return rows


def write_methodology(path: Path, seed: int, ka1_pages: list[int], ka11_pages: list[int]) -> None:
    lines = [
        "# Phase 5 Independent Audit Methodology",
        "",
        "## Method",
        "",
        "Tabula and Camelot were not available in the local Python environment. "
        "The audit therefore used pdfminer.six directly as the independent extractor.",
        "",
        "The independent extraction path differs from the production parser:",
        "",
        "- production parser: pdfplumber word extraction plus geometry-derived row/column bands",
        "- audit parser: pdfminer text-line extraction, y-line grouping, and left-to-right token sequence parsing",
        "- no Parsed_Grid or Canonical_Long values are used to parse PDF rows",
        "",
        "## Sampling",
        "",
        f"Random seed: `{seed}`",
        f"KA1-KA10 sampled pages: `{','.join(str(page) for page in ka1_pages)}`",
        f"KA11-KA20 sampled pages: `{','.join(str(page) for page in ka11_pages)}`",
        "",
        "## Match Rules",
        "",
        "- salary_from must match exactly",
        "- salary_to must match exactly",
        "- dependent_code must match exactly",
        "- raw deduction value must match exactly",
        "- `-` is compared as a literal dash, not zero",
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_phase_summary(
    path: Path,
    sampled_pages: list[int],
    sample_results: list[dict[str, Any]],
    mismatches: list[dict[str, Any]],
    aggregate_rows: list[dict[str, Any]],
) -> None:
    exact_matches = sum(1 for row in sample_results if row["exact_match"])
    mismatch_count = len(mismatches)
    mismatch_rate = (mismatch_count / len(sample_results)) if sample_results else 0
    aggregate_failures = [row for row in aggregate_rows if row["status"] != "pass"]
    systematic_shift = any(
        "raw_value" in str(row.get("mismatch_type", ""))
        or "dependent" in str(row.get("mismatch_type", ""))
        for row in mismatches
    )
    passes = mismatch_count == 0 and not aggregate_failures
    lines = [
        "# Phase 5 Independent Audit Summary",
        "",
        f"- Pages audited: `{len(sampled_pages)}`",
        f"- Records audited: `{len(sample_results)}`",
        f"- Exact matches: `{exact_matches}`",
        f"- Mismatch count: `{mismatch_count}`",
        f"- Mismatch rate: `{mismatch_rate:.8%}`",
        f"- Aggregate check failures: `{len(aggregate_failures)}`",
        f"- Systematic column-shift evidence: `{'yes' if systematic_shift else 'no'}`",
        "",
        "## Result",
        "",
        (
            "Canonical_Long passes independent verification."
            if passes
            else "Canonical_Long does not pass independent verification until mismatches are reviewed."
        ),
    ]
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    args = parse_args()
    project_root = args.project_root.expanduser().resolve()
    pdf_path = locate_pdf(project_root, args.pdf)
    canonical_path = (
        args.canonical.expanduser().resolve()
        if args.canonical is not None
        else project_root / DEFAULT_CANONICAL
    )
    classification_path = (
        args.classification.expanduser().resolve()
        if args.classification is not None
        else project_root / DEFAULT_CLASSIFICATION
    )
    output_dir = (
        args.output_dir.expanduser().resolve()
        if args.output_dir is not None
        else project_root / DEFAULT_OUTPUT_DIR
    )

    if not canonical_path.exists():
        raise FileNotFoundError(f"Missing canonical CSV: {canonical_path}")
    if not classification_path.exists():
        raise FileNotFoundError(f"Missing classification CSV: {classification_path}")

    classification = load_classification(classification_path)
    canonical_by_key, _canonical_by_page, canonical_records = load_canonical(canonical_path)
    ka1_sample, ka11_sample = sample_pages(classification, args.seed)
    sampled_pages = sorted(ka1_sample + ka11_sample)
    pdfminer_lines_by_page = grouped_pdfminer_lines_all(pdf_path)

    independent_all: list[IndependentRecord] = []
    independent_sample: list[IndependentRecord] = []
    mismatches: list[dict[str, Any]] = []

    for page_number, class_row in sorted(classification.items()):
        section_code = class_row.get("detected_section_code", "")
        if not section_code:
            continue
        if class_row.get("detected_family") == "cover_ka11_ka20":
            continue
        page_records, page_errors = parse_independent_page_lines(
            pdfminer_lines_by_page.get(page_number, []), page_number, section_code
        )
        independent_all.extend(page_records)
        if page_number in sampled_pages:
            independent_sample.extend(page_records)
        for error in page_errors:
            mismatches.append({"mismatch_scope": "independent_extraction", **error})

    sample_results, sample_mismatches = compare_sample_records(
        independent_sample, canonical_by_key
    )
    mismatches.extend(sample_mismatches)

    canonical_agg = aggregate_canonical(canonical_records)
    independent_agg = aggregate_independent(independent_all)
    summary_rows = audit_summary_rows(
        canonical_agg, independent_agg, sample_results, sampled_pages
    )
    for row in summary_rows:
        if row["status"] != "pass":
            mismatches.append(
                {
                    "mismatch_scope": "aggregate",
                    "mismatch_type": row["metric"],
                    "source_page": "",
                    "source_row_index": "",
                    "source_column_label": row["scope"],
                    "canonical_value": row["canonical_count"],
                    "independent_value": row["independent_count"],
                    "details": row["details"],
                }
            )

    output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(
        output_dir / "audit_sample_results.csv",
        sample_results,
        [
            "source_page",
            "source_row_index",
            "source_column_label",
            "dependent_code",
            "canonical_salary_from",
            "independent_salary_from",
            "canonical_salary_to",
            "independent_salary_to",
            "canonical_raw_value",
            "independent_raw_value",
            "canonical_blank_status",
            "exact_match",
            "mismatch_reason",
        ],
    )
    write_csv(
        output_dir / "audit_mismatches.csv",
        mismatches,
        [
            "mismatch_scope",
            "mismatch_type",
            "source_page",
            "source_row_index",
            "source_column_label",
            "canonical_value",
            "independent_value",
            "details",
        ],
    )
    write_csv(
        output_dir / "audit_summary.csv",
        summary_rows,
        ["metric", "scope", "canonical_count", "independent_count", "status", "details"],
    )
    write_methodology(output_dir / "audit_methodology.md", args.seed, ka1_sample, ka11_sample)
    write_phase_summary(
        output_dir / "phase5_audit_summary.md",
        sampled_pages,
        sample_results,
        mismatches,
        summary_rows,
    )

    exact_matches = sum(1 for row in sample_results if row["exact_match"])
    mismatch_rate = (len(mismatches) / len(sample_results)) if sample_results else 0
    systematic_shift = any(
        "raw_value" in str(row.get("mismatch_type", ""))
        or "dependent" in str(row.get("mismatch_type", ""))
        for row in mismatches
    )
    print(f"Pages audited: {len(sampled_pages)}")
    print(f"Records audited: {len(sample_results)}")
    print(f"Exact matches: {exact_matches}")
    print(f"Mismatch count: {len(mismatches)}")
    print(f"Mismatch rate: {mismatch_rate:.8%}")
    print(f"Systematic column-shift evidence: {'yes' if systematic_shift else 'no'}")
    print(f"Output: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
