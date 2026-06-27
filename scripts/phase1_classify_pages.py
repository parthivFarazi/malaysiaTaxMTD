#!/usr/bin/env python3
"""Classify every page into Phase 1 template families.

This is a header/page-shape validation script. It does not parse tax values,
assign data cells, or produce Excel output.
"""

from __future__ import annotations

import argparse
import csv
import re
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


DEFAULT_OUTPUT_DIR = "phase1_discovery_output"
LOW_CONFIDENCE_THRESHOLD = 0.80
LINE_Y_TOLERANCE = 3.0
DATA_ROW_PATTERN = re.compile(r"^\s*\d+\s*-\s*\d+\b")

FAMILIES = [
    "intro_data_ka1_ka10",
    "normal_ka1_ka10",
    "short_final_ka1_ka10",
    "cover_ka11_ka20",
    "intro_data_ka11_ka20",
    "normal_ka11_ka20",
    "short_final_ka11_ka20",
    "non_data_or_unknown",
]

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

KA1_KA10_CODES = ["B", "K", *[f"KA{index}" for index in range(1, 11)]]
KA11_KA20_CODES = [f"KA{index}" for index in range(11, 21)]


@dataclass(frozen=True)
class Line:
    y_center: float
    text: str
    tokens: tuple[str, ...]

    @property
    def upper_text(self) -> str:
        return self.text.upper()

    @property
    def upper_tokens(self) -> tuple[str, ...]:
        return tuple(token.upper() for token in self.tokens)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Classify every PDF page into discovered Phase 1 template families."
    )
    parser.add_argument(
        "--project-root",
        type=Path,
        default=Path.cwd(),
        help="Project folder containing the source PDF. Defaults to the current directory.",
    )
    parser.add_argument(
        "--pdf",
        type=Path,
        default=None,
        help="Optional explicit PDF path. If omitted, exactly one PDF must exist under the project root.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help=f"Output directory. Defaults to {DEFAULT_OUTPUT_DIR}/ under the project root.",
    )
    parser.add_argument(
        "--low-confidence-threshold",
        type=float,
        default=LOW_CONFIDENCE_THRESHOLD,
        help="Pages below this confidence are written to unknown_or_low_confidence_pages.csv.",
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


def page_rotation(page: Any) -> int | None:
    rotation = getattr(page, "rotation", None)
    if rotation is not None:
        return int(rotation)
    attrs = getattr(getattr(page, "page_obj", None), "attrs", {})
    raw_rotation = attrs.get("Rotate") if isinstance(attrs, dict) else None
    return int(raw_rotation) if raw_rotation is not None else None


def extract_words(page: Any) -> list[dict[str, Any]]:
    return page.extract_words(
        x_tolerance=1,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=False,
    )


def group_words_into_lines(words: list[dict[str, Any]]) -> list[Line]:
    lines: list[dict[str, Any]] = []
    for word in sorted(words, key=lambda item: ((float(item["top"]) + float(item["bottom"])) / 2, float(item["x0"]))):
        text = str(word.get("text", "")).strip()
        if not text:
            continue
        y_center = (float(word["top"]) + float(word["bottom"])) / 2
        word_payload = {
            "text": text,
            "x0": float(word["x0"]),
            "y_center": y_center,
        }
        for line in lines:
            if abs(float(line["y_center"]) - y_center) <= LINE_Y_TOLERANCE:
                line["items"].append(word_payload)
                line["ys"].append(y_center)
                line["y_center"] = sum(line["ys"]) / len(line["ys"])
                break
        else:
            lines.append({"y_center": y_center, "ys": [y_center], "items": [word_payload]})

    output: list[Line] = []
    for line in lines:
        sorted_items = sorted(line["items"], key=lambda item: item["x0"])
        tokens = tuple(item["text"] for item in sorted_items)
        output.append(
            Line(
                y_center=float(line["y_center"]),
                text=" ".join(tokens),
                tokens=tokens,
            )
        )
    return output


def expected_family_for_page(page_number: int) -> str:
    if page_number == 2:
        return "intro_data_ka1_ka10"
    if 3 <= page_number <= 354:
        return "normal_ka1_ka10"
    if page_number == 355:
        return "short_final_ka1_ka10"
    if page_number == 356:
        return "cover_ka11_ka20"
    if page_number == 357:
        return "intro_data_ka11_ka20"
    if 358 <= page_number <= 732:
        return "normal_ka11_ka20"
    if page_number == 733:
        return "short_final_ka11_ka20"
    return "non_data_or_unknown"


def best_header_line(lines: list[Line]) -> tuple[str, list[str], float, str]:
    best_family = ""
    best_labels: list[str] = []
    best_score = 0.0
    best_text = ""

    for line in lines:
        tokens = list(line.upper_tokens)
        ka11_count = sum(1 for code in KA11_KA20_CODES if code in tokens)
        ka1_count = sum(1 for code in KA1_KA10_CODES if code in tokens)
        has_repeated_ka11 = sum(1 for token in tokens if token in KA11_KA20_CODES) >= 15
        has_repeated_ka1 = sum(1 for token in tokens if token in KA1_KA10_CODES) >= 15

        if ka11_count >= 8 or has_repeated_ka11:
            score = ka11_count + (5 if has_repeated_ka11 else 0)
            if score > best_score:
                best_family = "KA11_KA20"
                best_labels = [token for token in tokens if token in KA11_KA20_CODES]
                best_score = float(score)
                best_text = line.text

        if {"B", "K", "KA1", "KA10"}.issubset(set(tokens)) or has_repeated_ka1:
            score = ka1_count + (5 if has_repeated_ka1 else 0)
            if score > best_score:
                best_family = "KA1_KA10"
                best_labels = [token for token in tokens if token in KA1_KA10_CODES]
                best_score = float(score)
                best_text = line.text

    confidence = min(1.0, best_score / 20.0) if best_score else 0.0
    return best_family, best_labels, confidence, best_text


def data_like_lines(lines: list[Line]) -> list[Line]:
    return [line for line in lines if DATA_ROW_PATTERN.match(line.text)]


def detect_section_code(
    joined_text: str, header_family: str, header_confidence: float
) -> tuple[str, float, list[str]]:
    reasons: list[str] = []
    if header_family == "KA11_KA20":
        reasons.append("header_labels_ka11_ka20")
        return "KA11_KA20", max(0.90, header_confidence), reasons
    if header_family == "KA1_KA10":
        reasons.append("header_labels_ka1_ka10")
        return "KA1_KA10", max(0.90, header_confidence), reasons

    if "KA11-KA20" in joined_text or "(KA11-KA20)" in joined_text:
        reasons.append("cover_or_title_text_ka11_ka20")
        return "KA11_KA20", 0.85, reasons
    if "BUJANG" in joined_text and "KA10" in joined_text:
        reasons.append("intro_text_ka1_ka10")
        return "KA1_KA10", 0.75, reasons

    return "", 0.0, reasons


def has_intro_evidence(joined_text: str) -> bool:
    return (
        "LEMBAGA HASIL DALAM NEGERI" in joined_text
        and "DEDUCTION/RELIEF" in joined_text
        and "MONTHLY TAX DEDUCTIONS" in joined_text
    )


def classify_page(page: Any) -> dict[str, Any]:
    words = extract_words(page)
    lines = group_words_into_lines(words)
    joined_text = " ".join(line.upper_text for line in lines)
    header_family, header_labels, header_confidence, header_text = best_header_line(lines)
    section_code, section_confidence, section_reasons = detect_section_code(
        joined_text, header_family, header_confidence
    )
    observed_data_rows = len(data_like_lines(lines))
    intro_evidence = has_intro_evidence(joined_text)
    cover_evidence = "KA11-KA20" in joined_text or "(KA11-KA20)" in joined_text
    notes: list[str] = []

    if not words:
        detected_family = "non_data_or_unknown"
        confidence = 0.10
        notes.append("no_positioned_words_extracted")
    elif observed_data_rows == 0:
        if section_code == "KA11_KA20" and cover_evidence:
            detected_family = "cover_ka11_ka20"
            confidence = 0.98
            notes.append("cover_text_detected_zero_data_rows")
        else:
            detected_family = "non_data_or_unknown"
            confidence = 0.70 if section_code else 0.40
            notes.append("zero_data_like_rows")
    elif section_code == "KA1_KA10":
        if intro_evidence:
            detected_family = "intro_data_ka1_ka10"
            confidence = 0.97
            notes.append("intro_text_and_ka1_ka10_header_detected")
        elif observed_data_rows < 40:
            detected_family = "short_final_ka1_ka10"
            confidence = 0.95
            notes.append("ka1_ka10_header_with_short_data_row_count")
        else:
            detected_family = "normal_ka1_ka10"
            confidence = 0.98
            notes.append("ka1_ka10_header_with_normal_data_row_count")
    elif section_code == "KA11_KA20":
        if intro_evidence:
            detected_family = "intro_data_ka11_ka20"
            confidence = 0.97
            notes.append("intro_text_and_ka11_ka20_header_detected")
        elif observed_data_rows < 40:
            detected_family = "short_final_ka11_ka20"
            confidence = 0.95
            notes.append("ka11_ka20_header_with_short_data_row_count")
        else:
            detected_family = "normal_ka11_ka20"
            confidence = 0.98
            notes.append("ka11_ka20_header_with_normal_data_row_count")
    else:
        detected_family = "non_data_or_unknown"
        confidence = 0.45
        notes.append("data_like_rows_without_recognized_section_header")

    if section_reasons:
        notes.extend(section_reasons)
    if header_text:
        evidence_text = header_text
    else:
        evidence_text = " | ".join(line.text for line in lines[:5])

    expected_family = expected_family_for_page(page.page_number)
    disagrees = detected_family != expected_family
    if disagrees:
        notes.append(f"expected_family_disagreement:{expected_family}")

    return {
        "page_number": page.page_number,
        "detected_family": detected_family,
        "confidence": f"{confidence:.2f}",
        "evidence_text": evidence_text,
        "detected_section_code": section_code,
        "detected_header_labels": "|".join(header_labels),
        "expected_data_rows": EXPECTED_ROWS_BY_FAMILY[detected_family],
        "observed_data_like_rows": observed_data_rows,
        "page_width": float(page.width),
        "page_height": float(page.height),
        "rotation": page_rotation(page),
        "expected_family": expected_family,
        "expected_range_disagreement": disagrees,
        "is_low_confidence": confidence < LOW_CONFIDENCE_THRESHOLD,
        "notes": "; ".join(notes),
    }


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def summary_rows(classification_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_family: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in classification_rows:
        by_family[str(row["detected_family"])].append(row)

    rows: list[dict[str, Any]] = []
    for family in FAMILIES:
        family_rows = by_family.get(family, [])
        observed_counts = [int(row["observed_data_like_rows"]) for row in family_rows]
        rows.append(
            {
                "detected_family": family,
                "page_count": len(family_rows),
                "page_numbers": "|".join(str(row["page_number"]) for row in family_rows),
                "expected_data_rows": EXPECTED_ROWS_BY_FAMILY[family],
                "observed_data_rows_min": min(observed_counts) if observed_counts else "",
                "observed_data_rows_max": max(observed_counts) if observed_counts else "",
                "observed_data_row_counts": "|".join(
                    str(count) for count in sorted(Counter(observed_counts))
                ),
                "low_confidence_pages": "|".join(
                    str(row["page_number"]) for row in family_rows if row["is_low_confidence"]
                ),
                "expected_range_disagreement_pages": "|".join(
                    str(row["page_number"])
                    for row in family_rows
                    if row["expected_range_disagreement"]
                ),
            }
        )
    return rows


def compact_page_ranges(page_numbers: list[int]) -> str:
    if not page_numbers:
        return ""
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


def main() -> int:
    args = parse_args()
    project_root = args.project_root.expanduser().resolve()
    output_dir = (
        args.output_dir.expanduser().resolve()
        if args.output_dir is not None
        else project_root / DEFAULT_OUTPUT_DIR
    )
    pdf_path = locate_pdf(project_root, args.pdf)

    classification_rows: list[dict[str, Any]] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            classification_rows.append(classify_page(page))

    for row in classification_rows:
        confidence = float(row["confidence"])
        row["is_low_confidence"] = confidence < args.low_confidence_threshold

    page_classification_fields = [
        "page_number",
        "detected_family",
        "confidence",
        "evidence_text",
        "detected_section_code",
        "detected_header_labels",
        "expected_data_rows",
        "observed_data_like_rows",
        "page_width",
        "page_height",
        "rotation",
        "expected_family",
        "expected_range_disagreement",
        "is_low_confidence",
        "notes",
    ]
    write_csv(
        output_dir / "page_classification.csv",
        classification_rows,
        page_classification_fields,
    )

    summary = summary_rows(classification_rows)
    write_csv(
        output_dir / "page_classification_summary.csv",
        summary,
        [
            "detected_family",
            "page_count",
            "page_numbers",
            "expected_data_rows",
            "observed_data_rows_min",
            "observed_data_rows_max",
            "observed_data_row_counts",
            "low_confidence_pages",
            "expected_range_disagreement_pages",
        ],
    )

    attention_rows = [
        row
        for row in classification_rows
        if row["detected_family"] == "non_data_or_unknown"
        or row["is_low_confidence"]
        or row["expected_range_disagreement"]
    ]
    write_csv(
        output_dir / "unknown_or_low_confidence_pages.csv",
        attention_rows,
        page_classification_fields,
    )

    family_counts = Counter(row["detected_family"] for row in classification_rows)
    low_confidence_pages = [
        int(row["page_number"]) for row in classification_rows if row["is_low_confidence"]
    ]
    disagreement_pages = [
        int(row["page_number"])
        for row in classification_rows
        if row["expected_range_disagreement"]
    ]

    print(f"PDF: {pdf_path}")
    print(f"Pages classified: {len(classification_rows)}")
    for family in FAMILIES:
        print(f"{family}: {family_counts[family]}")
    print(f"Low-confidence pages: {compact_page_ranges(low_confidence_pages) or 'none'}")
    print(f"Expected-range disagreements: {compact_page_ranges(disagreement_pages) or 'none'}")
    print(f"Output: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
