#!/usr/bin/env python3
"""Phase 2 geometry parser prototype for representative pages.

This script proves section-aware geometry assignment on selected pages only.
It does not process the full PDF, normalize tax values, emit Canonical_Long, or
export Excel.
"""

from __future__ import annotations

import argparse
import csv
import re
import statistics
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    import pdfplumber
except ImportError as exc:  # pragma: no cover - exercised only when missing locally
    raise SystemExit(
        "Missing dependency: pdfplumber. Install Phase 1 dependencies with "
        "`python3 -m pip install -r requirements.txt`."
    ) from exc


DEFAULT_OUTPUT_DIR = "phase2_parser_prototype_output"
DEFAULT_CLASSIFICATION_PATH = "phase1_discovery_output/page_classification.csv"
TARGET_PAGES = [2, 3, 10, 355, 357, 358, 400, 733]
LINE_Y_TOLERANCE = 3.0
DATA_ROW_PATTERN = re.compile(r"^\s*\d+\s*-\s*\d+\b")
INTEGER_PATTERN = re.compile(r"^\d+$")


@dataclass(frozen=True)
class Word:
    page_number: int
    token_index: int
    text: str
    x0: float
    x1: float
    top: float
    bottom: float

    @property
    def x_center(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def y_center(self) -> float:
        return (self.top + self.bottom) / 2


@dataclass(frozen=True)
class Line:
    y_center: float
    top: float
    bottom: float
    words: tuple[Word, ...]

    @property
    def text(self) -> str:
        return " ".join(word.text for word in self.words)

    @property
    def upper_tokens(self) -> list[str]:
        return [word.text.upper() for word in self.words]


@dataclass(frozen=True)
class ColumnBand:
    column_name: str
    source_label: str
    source_type: str
    center: float
    left: float
    right: float
    header_x0: float | None = None
    header_x1: float | None = None


@dataclass(frozen=True)
class RowBand:
    row_index: int
    center: float
    top: float
    bottom: float
    salary_from: str
    salary_to: str
    line_text: str
    line_words: tuple[Word, ...]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run Phase 2 geometry parser prototype on representative pages only."
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
        "--output-dir",
        type=Path,
        default=None,
        help=f"Output directory. Defaults to {DEFAULT_OUTPUT_DIR}/ under the project root.",
    )
    parser.add_argument(
        "--classification-csv",
        type=Path,
        default=None,
        help=f"Page classification CSV. Defaults to {DEFAULT_CLASSIFICATION_PATH}.",
    )
    parser.add_argument(
        "--pages",
        default=",".join(str(page) for page in TARGET_PAGES),
        help="Comma-separated target pages or ranges.",
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


def parse_page_ranges(raw_ranges: str, page_count: int) -> list[int]:
    pages: set[int] = set()
    for raw_part in raw_ranges.split(","):
        part = raw_part.strip()
        if not part:
            continue
        if "-" in part:
            start_text, end_text = part.split("-", 1)
            start = int(start_text.strip())
            end = int(end_text.strip())
            if start > end:
                raise ValueError(f"Invalid page range: {part}")
            pages.update(range(max(1, start), min(page_count, end) + 1))
        else:
            page = int(part)
            if page < 1 or page > page_count:
                raise ValueError(f"Page out of bounds: {page}; page count is {page_count}")
            pages.add(page)
    return sorted(pages)


def page_rotation(page: Any) -> int | None:
    rotation = getattr(page, "rotation", None)
    if rotation is not None:
        return int(rotation)
    attrs = getattr(getattr(page, "page_obj", None), "attrs", {})
    raw_rotation = attrs.get("Rotate") if isinstance(attrs, dict) else None
    return int(raw_rotation) if raw_rotation is not None else None


def expected_columns(section_code: str) -> list[str]:
    if section_code == "KA1_KA10":
        return [
            "salary_from",
            "salary_to",
            "B",
            "cat2_K",
            *[f"cat2_KA{index}" for index in range(1, 11)],
            "cat3_K",
            *[f"cat3_KA{index}" for index in range(1, 11)],
        ]
    if section_code == "KA11_KA20":
        return [
            "salary_from",
            "salary_to",
            *[f"cat2_KA{index}" for index in range(11, 21)],
            *[f"cat3_KA{index}" for index in range(11, 21)],
        ]
    raise ValueError(f"Unsupported section_code: {section_code}")


def output_grid_columns() -> list[str]:
    return [
        "page_number",
        "template_family",
        "section_code",
        "row_index",
        "source_row_y_center",
        "salary_from",
        "salary_to",
        "B",
        "cat2_K",
        *[f"cat2_KA{index}" for index in range(1, 21)],
        "cat3_K",
        *[f"cat3_KA{index}" for index in range(1, 21)],
        "row_assignment_status",
        "row_notes",
    ]


def extract_words(page: Any) -> list[Word]:
    raw_words = page.extract_words(
        x_tolerance=1,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=False,
    )
    words: list[Word] = []
    for token_index, word in enumerate(raw_words, start=1):
        text = str(word.get("text", "")).strip()
        if not text:
            continue
        words.append(
            Word(
                page_number=page.page_number,
                token_index=token_index,
                text=text,
                x0=float(word["x0"]),
                x1=float(word["x1"]),
                top=float(word["top"]),
                bottom=float(word["bottom"]),
            )
        )
    return words


def group_words_into_lines(words: list[Word]) -> list[Line]:
    grouped: list[dict[str, Any]] = []
    for word in sorted(words, key=lambda item: (item.y_center, item.x0)):
        for line in grouped:
            if abs(float(line["y_center"]) - word.y_center) <= LINE_Y_TOLERANCE:
                line["words"].append(word)
                line["ys"].append(word.y_center)
                line["y_center"] = sum(line["ys"]) / len(line["ys"])
                break
        else:
            grouped.append({"y_center": word.y_center, "ys": [word.y_center], "words": [word]})

    lines: list[Line] = []
    for line in grouped:
        line_words = tuple(sorted(line["words"], key=lambda item: item.x0))
        lines.append(
            Line(
                y_center=float(line["y_center"]),
                top=min(word.top for word in line_words),
                bottom=max(word.bottom for word in line_words),
                words=line_words,
            )
        )
    return lines


def load_classification(path: Path) -> dict[int, dict[str, str]]:
    with path.open(newline="", encoding="utf-8") as handle:
        return {int(row["page_number"]): row for row in csv.DictReader(handle)}


def infer_section_from_family(template_family: str) -> str:
    if "ka1_ka10" in template_family:
        return "KA1_KA10"
    if "ka11_ka20" in template_family:
        return "KA11_KA20"
    raise ValueError(f"Cannot infer section from template family: {template_family}")


def header_line_for_section(lines: list[Line], section_code: str) -> Line:
    best_line: Line | None = None
    best_score = -1
    for line in lines:
        tokens = line.upper_tokens
        if section_code == "KA1_KA10":
            score = sum(token in {"B", "K", *[f"KA{i}" for i in range(1, 11)]} for token in tokens)
            has_required = {"B", "K", "KA1", "KA10"}.issubset(set(tokens))
        else:
            score = sum(token in {f"KA{i}" for i in range(11, 21)} for token in tokens)
            has_required = {"KA11", "KA12", "KA20"}.issubset(set(tokens))
        if has_required and score > best_score:
            best_score = score
            best_line = line

    if best_line is None:
        raise RuntimeError(f"Could not detect header line for section {section_code}")
    return best_line


def header_specs(section_code: str) -> list[tuple[str, str]]:
    if section_code == "KA1_KA10":
        return [
            ("B", "B"),
            ("cat2_K", "K"),
            *[(f"cat2_KA{index}", f"KA{index}") for index in range(1, 11)],
            ("cat3_K", "K"),
            *[(f"cat3_KA{index}", f"KA{index}") for index in range(1, 11)],
        ]
    return [
        *[(f"cat2_KA{index}", f"KA{index}") for index in range(11, 21)],
        *[(f"cat3_KA{index}", f"KA{index}") for index in range(11, 21)],
    ]


def derive_row_bands(lines: list[Line]) -> list[RowBand]:
    data_lines = [line for line in lines if DATA_ROW_PATTERN.match(line.text)]
    if not data_lines:
        raise RuntimeError("No data-like salary-range rows detected")

    row_centers = [line.y_center for line in data_lines]
    gaps = [
        row_centers[index + 1] - row_centers[index]
        for index in range(len(row_centers) - 1)
    ]
    fallback_half_gap = (statistics.median(gaps) / 2) if gaps else 4.0

    row_bands: list[RowBand] = []
    for index, line in enumerate(data_lines):
        top = (
            (row_centers[index - 1] + line.y_center) / 2
            if index > 0
            else line.y_center - fallback_half_gap
        )
        bottom = (
            (line.y_center + row_centers[index + 1]) / 2
            if index < len(data_lines) - 1
            else line.y_center + fallback_half_gap
        )
        words = line.words
        if len(words) < 3 or not INTEGER_PATTERN.match(words[0].text) or words[1].text != "-":
            raise RuntimeError(f"Unexpected salary row token shape: {line.text}")
        row_bands.append(
            RowBand(
                row_index=index + 1,
                center=line.y_center,
                top=top,
                bottom=bottom,
                salary_from=words[0].text,
                salary_to=words[2].text,
                line_text=line.text,
                line_words=words,
            )
        )
    return row_bands


def median_word_center(row_bands: list[RowBand], word_index: int) -> float:
    return statistics.median(row.line_words[word_index].x_center for row in row_bands)


def derive_column_bands(
    lines: list[Line], row_bands: list[RowBand], section_code: str, page_width: float
) -> list[ColumnBand]:
    header_line = header_line_for_section(lines, section_code)
    specs = header_specs(section_code)
    header_words = [
        word
        for word in header_line.words
        if word.text.upper() in {label for _, label in specs}
    ]
    if len(header_words) != len(specs):
        raise RuntimeError(
            f"Header label count mismatch for {section_code}: "
            f"expected {len(specs)}, observed {len(header_words)} on line: {header_line.text}"
        )

    centers: list[dict[str, Any]] = [
        {
            "column_name": "salary_from",
            "source_label": "salary_from_observed",
            "source_type": "salary_observed",
            "center": median_word_center(row_bands, 0),
            "header_x0": None,
            "header_x1": None,
        },
        {
            "column_name": "salary_to",
            "source_label": "salary_to_observed",
            "source_type": "salary_observed",
            "center": median_word_center(row_bands, 2),
            "header_x0": None,
            "header_x1": None,
        },
    ]

    for (column_name, source_label), word in zip(specs, header_words, strict=True):
        centers.append(
            {
                "column_name": column_name,
                "source_label": source_label,
                "source_type": "header_left_anchor",
                "center": word.x0,
                "header_x0": word.x0,
                "header_x1": word.x1,
            }
        )

    centers.sort(key=lambda item: float(item["center"]))
    bands: list[ColumnBand] = []
    for index, item in enumerate(centers):
        center = float(item["center"])
        if index == 0:
            next_center = float(centers[index + 1]["center"])
            left = center - ((next_center - center) / 2)
        else:
            left = (float(centers[index - 1]["center"]) + center) / 2
        if index == len(centers) - 1:
            right = page_width
        else:
            right = (center + float(centers[index + 1]["center"])) / 2
        bands.append(
            ColumnBand(
                column_name=str(item["column_name"]),
                source_label=str(item["source_label"]),
                source_type=str(item["source_type"]),
                center=center,
                left=left,
                right=right,
                header_x0=item["header_x0"],
                header_x1=item["header_x1"],
            )
        )
    return bands


def find_band_for_x(column_bands: list[ColumnBand], x_center: float) -> ColumnBand | None:
    for band in column_bands:
        if band.left <= x_center < band.right:
            return band
    return None


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def page_context(page_number: int, classification: dict[int, dict[str, str]]) -> tuple[str, str]:
    row = classification.get(page_number)
    if not row:
        raise RuntimeError(
            f"Page {page_number} is missing from page classification CSV. "
            "Run scripts/phase1_classify_pages.py first."
        )
    template_family = row["detected_family"]
    section_code = row.get("detected_section_code") or infer_section_from_family(template_family)
    return template_family, section_code


def parse_page(
    page: Any, template_family: str, section_code: str
) -> tuple[
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
    dict[str, Any],
]:
    words = extract_words(page)
    lines = group_words_into_lines(words)
    row_bands = derive_row_bands(lines)
    column_bands = derive_column_bands(lines, row_bands, section_code, float(page.width))
    expected = expected_columns(section_code)
    expected_set = set(expected)

    row_by_word_token: dict[int, RowBand] = {}
    for row in row_bands:
        for word in row.line_words:
            row_by_word_token[word.token_index] = row

    grid_rows: list[dict[str, Any]] = []
    cell_tokens: dict[tuple[int, str], list[Word]] = defaultdict(list)
    token_assignment_rows: list[dict[str, Any]] = []
    unassigned_data_tokens: list[Word] = []
    duplicate_cells: list[tuple[int, str, list[str]]] = []

    for row in row_bands:
        grid = {
            "page_number": page.page_number,
            "template_family": template_family,
            "section_code": section_code,
            "row_index": row.row_index,
            "source_row_y_center": f"{row.center:.4f}",
            "row_assignment_status": "ok",
            "row_notes": "",
        }
        for column in output_grid_columns():
            grid.setdefault(column, "")
        grid["salary_from"] = row.salary_from
        grid["salary_to"] = row.salary_to

        for word_position, word in enumerate(row.line_words):
            if word_position == 1 and word.text == "-":
                token_assignment_rows.append(
                    token_assignment_row(
                        page.page_number,
                        word,
                        "salary_range_separator",
                        row,
                        "",
                        "salary separator between salary_from and salary_to",
                    )
                )
                continue

            if word_position == 0:
                column_name = "salary_from"
            elif word_position == 2:
                column_name = "salary_to"
            else:
                band = find_band_for_x(column_bands, word.x0)
                column_name = band.column_name if band is not None else ""

            if not column_name or column_name not in expected_set:
                unassigned_data_tokens.append(word)
                token_assignment_rows.append(
                    token_assignment_row(
                        page.page_number,
                        word,
                        "unassigned_data_token",
                        row,
                        column_name,
                        "data-row token did not map to expected column band",
                    )
                )
                continue

            cell_tokens[(row.row_index, column_name)].append(word)
            token_assignment_rows.append(
                token_assignment_row(
                    page.page_number,
                    word,
                    "assigned_cell",
                    row,
                    column_name,
                    "",
                )
            )

        for column_name in expected:
            tokens = cell_tokens.get((row.row_index, column_name), [])
            if tokens:
                grid[column_name] = " ".join(token.text for token in tokens)
                if len(tokens) > 1:
                    duplicate_cells.append(
                        (row.row_index, column_name, [token.text for token in tokens])
                    )
                    grid["row_assignment_status"] = "review"
                    grid["row_notes"] = append_note(
                        grid["row_notes"], f"multiple_tokens_in_{column_name}"
                    )
            else:
                grid[column_name] = ""
                grid["row_assignment_status"] = "review"
                grid["row_notes"] = append_note(grid["row_notes"], f"missing_{column_name}")

        grid_rows.append(grid)

    for word in words:
        if word.token_index in row_by_word_token:
            continue
        assignment_type = classify_non_data_token(word, lines, row_bands)
        token_assignment_rows.append(
            token_assignment_row(
                page.page_number,
                word,
                assignment_type,
                None,
                "",
                "",
            )
        )

    column_report_rows = [
        {
            "page_number": page.page_number,
            "template_family": template_family,
            "section_code": section_code,
            "column_name": band.column_name,
            "source_label": band.source_label,
            "source_type": band.source_type,
            "center": f"{band.center:.4f}",
            "left_boundary": f"{band.left:.4f}",
            "right_boundary": f"{band.right:.4f}",
            "header_x0": "" if band.header_x0 is None else f"{band.header_x0:.4f}",
            "header_x1": "" if band.header_x1 is None else f"{band.header_x1:.4f}",
            "assigned_token_count": sum(
                1
                for row in row_bands
                for token in cell_tokens.get((row.row_index, band.column_name), [])
            ),
        }
        for band in column_bands
    ]

    row_report_rows = [
        {
            "page_number": page.page_number,
            "template_family": template_family,
            "section_code": section_code,
            "row_index": row.row_index,
            "center": f"{row.center:.4f}",
            "top_boundary": f"{row.top:.4f}",
            "bottom_boundary": f"{row.bottom:.4f}",
            "salary_from": row.salary_from,
            "salary_to": row.salary_to,
            "line_token_count": len(row.line_words),
            "expected_cell_count": len(expected),
            "assigned_cell_token_count": sum(
                len(cell_tokens.get((row.row_index, column_name), []))
                for column_name in expected
            ),
            "line_text": row.line_text,
        }
        for row in row_bands
    ]

    summary = {
        "page_number": page.page_number,
        "template_family": template_family,
        "section_code": section_code,
        "page_width": float(page.width),
        "page_height": float(page.height),
        "rotation": page_rotation(page),
        "row_count": len(row_bands),
        "column_count": len(column_bands),
        "word_count": len(words),
        "assigned_cell_tokens": sum(
            1 for row in token_assignment_rows if row["assignment_type"] == "assigned_cell"
        ),
        "salary_separator_tokens": sum(
            1
            for row in token_assignment_rows
            if row["assignment_type"] == "salary_range_separator"
        ),
        "non_data_tokens": sum(
            1
            for row in token_assignment_rows
            if row["assignment_type"] in {"header_token", "footer_token", "non_data_token"}
        ),
        "unassigned_data_tokens": len(unassigned_data_tokens),
        "duplicate_cells": len(duplicate_cells),
        "grid_review_rows": sum(1 for row in grid_rows if row["row_assignment_status"] != "ok"),
    }
    return grid_rows, column_report_rows, row_report_rows, token_assignment_rows, summary


def append_note(existing: str, note: str) -> str:
    return note if not existing else f"{existing};{note}"


def classify_non_data_token(word: Word, lines: list[Line], row_bands: list[RowBand]) -> str:
    if row_bands and word.y_center > row_bands[-1].bottom:
        return "footer_token"
    return "header_token"


def token_assignment_row(
    page_number: int,
    word: Word,
    assignment_type: str,
    row: RowBand | None,
    column_name: str,
    notes: str,
) -> dict[str, Any]:
    return {
        "page_number": page_number,
        "token_index": word.token_index,
        "text": word.text,
        "x0": f"{word.x0:.4f}",
        "x1": f"{word.x1:.4f}",
        "top": f"{word.top:.4f}",
        "bottom": f"{word.bottom:.4f}",
        "x_center": f"{word.x_center:.4f}",
        "y_center": f"{word.y_center:.4f}",
        "assignment_type": assignment_type,
        "row_index": "" if row is None else row.row_index,
        "column_name": column_name,
        "notes": notes,
    }


def markdown_table(rows: list[dict[str, Any]], columns: list[str], max_rows: int = 5) -> str:
    selected = rows[:max_rows]
    lines = [
        "| " + " | ".join(columns) + " |",
        "| " + " | ".join("---" for _ in columns) + " |",
    ]
    for row in selected:
        values = [str(row.get(column, "")).replace("|", "\\|") for column in columns]
        lines.append("| " + " | ".join(values) + " |")
    return "\n".join(lines)


def write_summary_markdown(
    path: Path,
    pdf_path: Path,
    target_pages: list[int],
    page_summaries: list[dict[str, Any]],
    parsed_grid_rows: list[dict[str, Any]],
) -> None:
    lines = [
        "# Phase 2 Parser Prototype Summary",
        "",
        f"PDF: `{pdf_path.name}`",
        f"Target pages: `{', '.join(str(page) for page in target_pages)}`",
        "",
        "This prototype derives section-aware column and row bands from page-local geometry.",
        "It preserves raw observed values and does not normalize tax amounts.",
        "",
        "## Page Results",
        "",
        "| Page | Family | Section | Rows | Columns | Unassigned Data Tokens | Duplicate Cells | Review Rows |",
        "| --- | --- | --- | ---: | ---: | ---: | ---: | ---: |",
    ]
    for summary in page_summaries:
        lines.append(
            "| {page_number} | {template_family} | {section_code} | {row_count} | "
            "{column_count} | {unassigned_data_tokens} | {duplicate_cells} | "
            "{grid_review_rows} |".format(**summary)
        )

    total_unassigned = sum(int(summary["unassigned_data_tokens"]) for summary in page_summaries)
    total_duplicates = sum(int(summary["duplicate_cells"]) for summary in page_summaries)
    lines.extend(
        [
            "",
            "## Validation",
            "",
            f"- Total unassigned data-like tokens: `{total_unassigned}`",
            f"- Total duplicate cells: `{total_duplicates}`",
            "- Deduction dashes are preserved as raw `-` cell values.",
            "- Salary range separator dashes are classified as non-data `salary_range_separator` tokens.",
            "",
            "## Parsed Grid Preview",
            "",
        ]
    )

    by_page: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in parsed_grid_rows:
        by_page[int(row["page_number"])].append(row)

    for page_number in target_pages:
        page_rows = by_page.get(page_number, [])
        if not page_rows:
            continue
        if page_rows[0]["section_code"] == "KA1_KA10":
            columns = ["page_number", "row_index", "salary_from", "salary_to", "B", "cat2_K", "cat2_KA1", "cat3_K", "cat3_KA1"]
        else:
            columns = ["page_number", "row_index", "salary_from", "salary_to", "cat2_KA11", "cat2_KA12", "cat3_KA11", "cat3_KA12"]
        lines.extend([f"### Page {page_number}", "", markdown_table(page_rows, columns), ""])

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    args = parse_args()
    project_root = args.project_root.expanduser().resolve()
    output_dir = (
        args.output_dir.expanduser().resolve()
        if args.output_dir is not None
        else project_root / DEFAULT_OUTPUT_DIR
    )
    classification_path = (
        args.classification_csv.expanduser().resolve()
        if args.classification_csv is not None
        else project_root / DEFAULT_CLASSIFICATION_PATH
    )
    if not classification_path.exists():
        raise FileNotFoundError(
            f"Missing classification CSV: {classification_path}. "
            "Run scripts/phase1_classify_pages.py first."
        )

    pdf_path = locate_pdf(project_root, args.pdf)
    classification = load_classification(classification_path)

    parsed_grid_rows: list[dict[str, Any]] = []
    column_report_rows: list[dict[str, Any]] = []
    row_report_rows: list[dict[str, Any]] = []
    token_assignment_rows: list[dict[str, Any]] = []
    page_summaries: list[dict[str, Any]] = []

    with pdfplumber.open(pdf_path) as pdf:
        target_pages = parse_page_ranges(args.pages, len(pdf.pages))
        for page_number in target_pages:
            page = pdf.pages[page_number - 1]
            template_family, section_code = page_context(page_number, classification)
            if template_family == "non_data_or_unknown":
                raise RuntimeError(f"Target page {page_number} is not a data page")
            page_grid, page_columns, page_rows, page_tokens, page_summary = parse_page(
                page, template_family, section_code
            )
            parsed_grid_rows.extend(page_grid)
            column_report_rows.extend(page_columns)
            row_report_rows.extend(page_rows)
            token_assignment_rows.extend(page_tokens)
            page_summaries.append(page_summary)

    output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(
        output_dir / "parsed_grid_sample.csv",
        parsed_grid_rows,
        output_grid_columns(),
    )
    write_csv(
        output_dir / "column_band_report.csv",
        column_report_rows,
        [
            "page_number",
            "template_family",
            "section_code",
            "column_name",
            "source_label",
            "source_type",
            "center",
            "left_boundary",
            "right_boundary",
            "header_x0",
            "header_x1",
            "assigned_token_count",
        ],
    )
    write_csv(
        output_dir / "row_band_report.csv",
        row_report_rows,
        [
            "page_number",
            "template_family",
            "section_code",
            "row_index",
            "center",
            "top_boundary",
            "bottom_boundary",
            "salary_from",
            "salary_to",
            "line_token_count",
            "expected_cell_count",
            "assigned_cell_token_count",
            "line_text",
        ],
    )
    write_csv(
        output_dir / "cell_assignment_report.csv",
        token_assignment_rows,
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
    write_summary_markdown(
        output_dir / "parser_prototype_summary.md",
        pdf_path,
        target_pages,
        page_summaries,
        parsed_grid_rows,
    )

    print(f"PDF: {pdf_path}")
    print(f"Target pages: {', '.join(str(page) for page in target_pages)}")
    print(f"Parsed grid rows: {len(parsed_grid_rows)}")
    print(f"Column band rows: {len(column_report_rows)}")
    print(f"Row band rows: {len(row_report_rows)}")
    print(f"Token assignment rows: {len(token_assignment_rows)}")
    print(
        "Unassigned data-like tokens: "
        f"{sum(summary['unassigned_data_tokens'] for summary in page_summaries)}"
    )
    print(f"Output: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
