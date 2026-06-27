#!/usr/bin/env python3
"""Phase 1 discovery for the Malaysian 2018 MTD/PCB PDF.

This script intentionally stops before production parsing. It samples selected
pages, extracts positioned text tokens with pdfplumber, and writes raw evidence
files for geometry review.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import platform
import statistics
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import UTC, datetime
from importlib import metadata as importlib_metadata
from pathlib import Path
from typing import Any, Iterable

try:
    import pdfplumber
except ImportError as exc:  # pragma: no cover - exercised only when missing locally
    raise SystemExit(
        "Missing dependency: pdfplumber. Install Phase 1 dependencies with "
        "`python3 -m pip install -r requirements.txt`."
    ) from exc


SCRIPT_VERSION = "phase1-discovery-0.1.0"
DEFAULT_OUTPUT_DIR = "phase1_discovery_output"


@dataclass(frozen=True)
class PageBox:
    page_number: int
    width: float
    height: float
    rotation: int | None

    @property
    def signature(self) -> str:
        return f"{self.width:.2f}x{self.height:.2f}/rot={self.rotation}"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run Phase 1 proof-of-extractability discovery on a sample of PDF pages."
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
        "--cluster-tolerance",
        type=float,
        default=3.0,
        help="Point tolerance for x/y coordinate clustering.",
    )
    parser.add_argument(
        "--extra-pages",
        default="",
        help=(
            "Optional comma-separated 1-based page numbers or ranges to include "
            "in the sample, for example: 359-363,733."
        ),
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


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def package_version(package_name: str) -> str | None:
    try:
        return importlib_metadata.version(package_name)
    except importlib_metadata.PackageNotFoundError:
        return None


def page_rotation(page: Any) -> int | None:
    rotation = getattr(page, "rotation", None)
    if rotation is not None:
        return int(rotation)
    attrs = getattr(getattr(page, "page_obj", None), "attrs", {})
    raw_rotation = attrs.get("Rotate") if isinstance(attrs, dict) else None
    return int(raw_rotation) if raw_rotation is not None else None


def collect_page_boxes(pdf: Any) -> list[PageBox]:
    boxes: list[PageBox] = []
    for page in pdf.pages:
        boxes.append(
            PageBox(
                page_number=page.page_number,
                width=float(page.width),
                height=float(page.height),
                rotation=page_rotation(page),
            )
        )
    return boxes


def contiguous_pages(start: int, end: int, page_count: int) -> list[int]:
    bounded_start = max(1, start)
    bounded_end = min(page_count, end)
    if bounded_start > bounded_end:
        return []
    return list(range(bounded_start, bounded_end + 1))


def parse_page_ranges(raw_ranges: str, page_count: int) -> set[int]:
    pages: set[int] = set()
    if not raw_ranges.strip():
        return pages

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
            pages.update(contiguous_pages(start, end, page_count))
        else:
            page_number = int(part)
            if page_number < 1 or page_number > page_count:
                raise ValueError(f"Page out of bounds: {page_number}; page count is {page_count}")
            pages.add(page_number)
    return pages


def choose_sample_pages(
    page_count: int, page_boxes: list[PageBox], extra_pages: set[int]
) -> dict[int, set[str]]:
    sample_reasons: dict[int, set[str]] = defaultdict(set)

    for page_number in contiguous_pages(2, 10, page_count):
        sample_reasons[page_number].add("first_data_pages_2_10")

    middle = max(1, round(page_count / 2))
    for page_number in contiguous_pages(middle - 4, middle + 4, page_count):
        sample_reasons[page_number].add("middle_pages")

    for page_number in contiguous_pages(page_count - 8, page_count, page_count):
        sample_reasons[page_number].add("final_data_pages")

    signature_counts = Counter(box.signature for box in page_boxes)
    dominant_signature = signature_counts.most_common(1)[0][0] if signature_counts else None
    for box in page_boxes:
        if dominant_signature is not None and box.signature != dominant_signature:
            sample_reasons[box.page_number].add("page_box_or_rotation_outlier")

    for page_number in sorted(extra_pages):
        sample_reasons[page_number].add("manual_extra_pages")

    return dict(sorted(sample_reasons.items()))


def clean_token_text(text: Any) -> str:
    return str(text).strip()


def extract_page_tokens(page: Any) -> list[dict[str, Any]]:
    words = page.extract_words(
        x_tolerance=1,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=False,
    )

    tokens: list[dict[str, Any]] = []
    for index, word in enumerate(words, start=1):
        text = clean_token_text(word.get("text", ""))
        x0 = float(word["x0"])
        x1 = float(word["x1"])
        top = float(word["top"])
        bottom = float(word["bottom"])
        tokens.append(
            {
                "page_number": page.page_number,
                "token_index": index,
                "text": text,
                "x0": x0,
                "x1": x1,
                "top": top,
                "bottom": bottom,
                "width": x1 - x0,
                "height": bottom - top,
                "x_center": (x0 + x1) / 2,
                "y_center": (top + bottom) / 2,
            }
        )
    return tokens


def cluster_values(items: Iterable[dict[str, Any]], metric: str, tolerance: float) -> list[dict[str, Any]]:
    sorted_items = sorted(items, key=lambda item: (float(item[metric]), item["page_number"]))
    clusters: list[list[dict[str, Any]]] = []

    for item in sorted_items:
        value = float(item[metric])
        if not clusters:
            clusters.append([item])
            continue
        current_values = [float(existing[metric]) for existing in clusters[-1]]
        current_mean = statistics.fmean(current_values)
        if abs(value - current_mean) <= tolerance:
            clusters[-1].append(item)
        else:
            clusters.append([item])

    output: list[dict[str, Any]] = []
    for cluster_id, cluster in enumerate(clusters, start=1):
        metric_values = [float(item[metric]) for item in cluster]
        texts = []
        seen_texts = set()
        for item in cluster:
            text = str(item["text"])
            if text and text not in seen_texts:
                texts.append(text)
                seen_texts.add(text)
            if len(texts) >= 8:
                break

        output.append(
            {
                "cluster_id": cluster_id,
                f"{metric}_mean": statistics.fmean(metric_values),
                f"{metric}_min": min(metric_values),
                f"{metric}_max": max(metric_values),
                f"{metric}_range": max(metric_values) - min(metric_values),
                "token_count": len(cluster),
                "page_count": len({item["page_number"] for item in cluster}),
                "sample_pages": "|".join(str(page) for page in sorted({item["page_number"] for item in cluster})),
                "sample_texts": "|".join(texts),
                "_items": cluster,
            }
        )
    return output


def column_candidates_for_metric(
    tokens: list[dict[str, Any]], tolerance: float
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    by_page: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for token in tokens:
        by_page[int(token["page_number"])].append(token)

    for metric in ("x0", "x_center", "x1"):
        for page_number in sorted(by_page):
            clusters = cluster_values(by_page[page_number], metric, tolerance)
            for cluster in clusters:
                items = cluster.pop("_items")
                rows.append(
                    {
                        "scope": "sample_page",
                        "page_number": page_number,
                        "metric": metric,
                        "cluster_id": cluster["cluster_id"],
                        "position_mean": cluster[f"{metric}_mean"],
                        "position_min": cluster[f"{metric}_min"],
                        "position_max": cluster[f"{metric}_max"],
                        "position_range": cluster[f"{metric}_range"],
                        "x0_min": min(float(item["x0"]) for item in items),
                        "x1_max": max(float(item["x1"]) for item in items),
                        "token_count": cluster["token_count"],
                        "page_count": cluster["page_count"],
                        "sample_pages": cluster["sample_pages"],
                        "sample_texts": cluster["sample_texts"],
                    }
                )

    for metric in ("x0", "x_center", "x1"):
        page_cluster_summaries = [
            {
                "page_number": row["page_number"],
                "text": row["sample_texts"],
                metric: row["position_mean"],
                "x0": row["x0_min"],
                "x1": row["x1_max"],
            }
            for row in rows
            if row["scope"] == "sample_page" and row["metric"] == metric
        ]
        aggregate_clusters = cluster_values(page_cluster_summaries, metric, tolerance)
        next_cluster_id = 1
        for cluster in aggregate_clusters:
            items = cluster.pop("_items")
            rows.append(
                {
                    "scope": "sample_aggregate",
                    "page_number": "",
                    "metric": metric,
                    "cluster_id": next_cluster_id,
                    "position_mean": cluster[f"{metric}_mean"],
                    "position_min": cluster[f"{metric}_min"],
                    "position_max": cluster[f"{metric}_max"],
                    "position_range": cluster[f"{metric}_range"],
                    "x0_min": min(float(item["x0"]) for item in items),
                    "x1_max": max(float(item["x1"]) for item in items),
                    "token_count": cluster["token_count"],
                    "page_count": cluster["page_count"],
                    "sample_pages": cluster["sample_pages"],
                    "sample_texts": cluster["sample_texts"],
                }
            )
            next_cluster_id += 1

    return rows


def row_candidates(tokens: list[dict[str, Any]], tolerance: float) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    by_page: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for token in tokens:
        by_page[int(token["page_number"])].append(token)

    for page_number in sorted(by_page):
        clusters = cluster_values(by_page[page_number], "y_center", tolerance)
        for cluster in clusters:
            items = cluster.pop("_items")
            rows.append(
                {
                    "scope": "sample_page",
                    "page_number": page_number,
                    "cluster_id": cluster["cluster_id"],
                    "y_center_mean": cluster["y_center_mean"],
                    "y_center_min": cluster["y_center_min"],
                    "y_center_max": cluster["y_center_max"],
                    "y_center_range": cluster["y_center_range"],
                    "top_min": min(float(item["top"]) for item in items),
                    "bottom_max": max(float(item["bottom"]) for item in items),
                    "token_count": cluster["token_count"],
                    "page_count": cluster["page_count"],
                    "sample_pages": cluster["sample_pages"],
                    "sample_texts": cluster["sample_texts"],
                }
            )
    return rows


def build_page_geometry_rows(
    page_boxes: list[PageBox],
    sample_reasons: dict[int, set[str]],
    tokens: list[dict[str, Any]],
    extraction_errors: dict[int, str],
) -> list[dict[str, Any]]:
    tokens_by_page: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for token in tokens:
        tokens_by_page[int(token["page_number"])].append(token)

    signature_counts = Counter(box.signature for box in page_boxes)
    dominant_signature = signature_counts.most_common(1)[0][0] if signature_counts else None

    rows: list[dict[str, Any]] = []
    for box in page_boxes:
        page_tokens = tokens_by_page.get(box.page_number, [])
        sampled = box.page_number in sample_reasons
        x_values = [float(token["x_center"]) for token in page_tokens]
        y_values = [float(token["y_center"]) for token in page_tokens]
        rows.append(
            {
                "page_number": box.page_number,
                "page_width": box.width,
                "page_height": box.height,
                "rotation": box.rotation,
                "page_box_signature": box.signature,
                "is_page_box_or_rotation_outlier": box.signature != dominant_signature,
                "is_sampled": sampled,
                "sample_reasons": "|".join(sorted(sample_reasons.get(box.page_number, set()))),
                "token_count": len(page_tokens) if sampled else "",
                "x_center_min": min(x_values) if x_values else "",
                "x_center_max": max(x_values) if x_values else "",
                "y_center_min": min(y_values) if y_values else "",
                "y_center_max": max(y_values) if y_values else "",
                "extraction_error": extraction_errors.get(box.page_number, ""),
            }
        )
    return rows


def write_csv(path: Path, rows: list[dict[str, Any]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2, sort_keys=True)
        handle.write("\n")


def median_numeric(values: list[Any]) -> float | None:
    numeric = [float(value) for value in values if value != "" and value is not None]
    if not numeric:
        return None
    return statistics.median(numeric)


def main() -> int:
    args = parse_args()
    project_root = args.project_root.expanduser().resolve()
    output_dir = (
        args.output_dir.expanduser().resolve()
        if args.output_dir is not None
        else project_root / DEFAULT_OUTPUT_DIR
    )
    pdf_path = locate_pdf(project_root, args.pdf)

    extraction_timestamp_utc = datetime.now(UTC).isoformat()
    extraction_timestamp_local = datetime.now().astimezone().isoformat()
    file_hash = sha256_file(pdf_path)

    all_tokens: list[dict[str, Any]] = []
    extraction_errors: dict[int, str] = {}

    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)
        page_boxes = collect_page_boxes(pdf)
        extra_pages = parse_page_ranges(args.extra_pages, page_count)
        sample_reasons = choose_sample_pages(page_count, page_boxes, extra_pages)

        for page_number in sorted(sample_reasons):
            page = pdf.pages[page_number - 1]
            try:
                all_tokens.extend(extract_page_tokens(page))
            except Exception as exc:  # noqa: BLE001 - discovery should record page-level failures.
                extraction_errors[page_number] = repr(exc)

        pdf_metadata = {
            "filename": pdf_path.name,
            "source_pdf_path": str(pdf_path),
            "file_size_bytes": pdf_path.stat().st_size,
            "sha256": file_hash,
            "page_count": page_count,
            "extraction_timestamp_utc": extraction_timestamp_utc,
            "extraction_timestamp_local": extraction_timestamp_local,
            "script_version": SCRIPT_VERSION,
            "tool_versions": {
                "python": sys.version.replace("\n", " "),
                "platform": platform.platform(),
                "pdfplumber": package_version("pdfplumber"),
                "pdfminer.six": package_version("pdfminer.six"),
                "pypdfium2": package_version("pypdfium2"),
            },
            "pdf_metadata": dict(pdf.metadata or {}),
            "sample_pages": [
                {
                    "page_number": page_number,
                    "reasons": sorted(reasons),
                }
                for page_number, reasons in sorted(sample_reasons.items())
            ],
            "sample_page_count": len(sample_reasons),
            "sampled_token_count": len(all_tokens),
            "extraction_errors": extraction_errors,
        }

    column_rows = column_candidates_for_metric(all_tokens, args.cluster_tolerance)
    row_rows = row_candidates(all_tokens, args.cluster_tolerance)
    geometry_rows = build_page_geometry_rows(
        page_boxes=page_boxes,
        sample_reasons=sample_reasons,
        tokens=all_tokens,
        extraction_errors=extraction_errors,
    )

    sampled_geometry = [row for row in geometry_rows if row["is_sampled"]]
    token_counts = [row["token_count"] for row in sampled_geometry]
    pdf_metadata["sample_summary"] = {
        "sampled_pages": len(sampled_geometry),
        "sampled_token_count": len(all_tokens),
        "sampled_token_count_min": min(token_counts) if token_counts else None,
        "sampled_token_count_max": max(token_counts) if token_counts else None,
        "sampled_token_count_median": median_numeric(token_counts),
        "page_box_signatures": dict(Counter(row["page_box_signature"] for row in geometry_rows)),
    }

    output_dir.mkdir(parents=True, exist_ok=True)
    write_json(output_dir / "pdf_metadata.json", pdf_metadata)

    write_csv(
        output_dir / "raw_tokens_sample.csv",
        all_tokens,
        [
            "page_number",
            "token_index",
            "text",
            "x0",
            "x1",
            "top",
            "bottom",
            "width",
            "height",
        ],
    )
    write_csv(
        output_dir / "page_geometry_summary.csv",
        geometry_rows,
        [
            "page_number",
            "page_width",
            "page_height",
            "rotation",
            "page_box_signature",
            "is_page_box_or_rotation_outlier",
            "is_sampled",
            "sample_reasons",
            "token_count",
            "x_center_min",
            "x_center_max",
            "y_center_min",
            "y_center_max",
            "extraction_error",
        ],
    )
    write_csv(
        output_dir / "detected_column_candidates.csv",
        column_rows,
        [
            "scope",
            "page_number",
            "metric",
            "cluster_id",
            "position_mean",
            "position_min",
            "position_max",
            "position_range",
            "x0_min",
            "x1_max",
            "token_count",
            "page_count",
            "sample_pages",
            "sample_texts",
        ],
    )
    write_csv(
        output_dir / "detected_row_candidates.csv",
        row_rows,
        [
            "scope",
            "page_number",
            "cluster_id",
            "y_center_mean",
            "y_center_min",
            "y_center_max",
            "y_center_range",
            "top_min",
            "bottom_max",
            "token_count",
            "page_count",
            "sample_pages",
            "sample_texts",
        ],
    )

    print(f"PDF: {pdf_path}")
    print(f"SHA-256: {file_hash}")
    print(f"Pages: {pdf_metadata['page_count']}")
    print(f"Sample pages: {len(pdf_metadata['sample_pages'])}")
    print(f"Sample tokens: {len(all_tokens)}")
    print(f"Output: {output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
