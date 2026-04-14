#!/usr/bin/env python3
"""
Inspect page-level table extraction patterns across a corpus of Esfer@ PDFs.

This helper is intended for cautious performance work: it summarizes where table
pages stop, how many trailing empty pages remain, and how much time hypothetical
early-exit rules might save without changing the parser itself.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import pdfplumber

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.pdf_processor import DEFAULT_TABLE_OPTS


def analyze_pdf(path: Path) -> dict[str, object]:
    with pdfplumber.open(path) as pdf:
        page_info: list[dict[str, object]] = []
        for index, page in enumerate(pdf.pages, start=1):
            started_at = time.perf_counter()
            raw = page.extract_table(DEFAULT_TABLE_OPTS)
            page_info.append(
                {
                    "page": index,
                    "has_table": bool(raw),
                    "rows": 0 if not raw else len(raw) - 1,
                    "extract_seconds": round(time.perf_counter() - started_at, 6),
                }
            )

    seen_table = False
    current_gap = 0
    internal_gaps_after_table: list[int] = []
    last_table_page = 0
    last_table_rows = 0
    tables_resume_after_gap = False

    for page in page_info:
        if page["has_table"]:
            if seen_table and current_gap:
                internal_gaps_after_table.append(current_gap)
                tables_resume_after_gap = True
            seen_table = True
            current_gap = 0
            last_table_page = int(page["page"])
            last_table_rows = int(page["rows"])
        elif seen_table:
            current_gap += 1

    trailing_gap = len(page_info) - last_table_page if last_table_page else len(page_info)
    return {
        "file": path.name,
        "pages": len(page_info),
        "table_pages": sum(1 for page in page_info if page["has_table"]),
        "first_table_page": next((page["page"] for page in page_info if page["has_table"]), None),
        "last_table_page": last_table_page or None,
        "last_table_rows": last_table_rows or None,
        "trailing_gap": trailing_gap,
        "tables_resume_after_gap": tables_resume_after_gap,
        "internal_gaps_after_table": internal_gaps_after_table,
        "page_info": page_info,
    }


def simulate_rule(
    docs: list[dict[str, object]],
    name: str,
    empty_threshold: int,
    min_progress: float = 0.0,
    max_remaining_pages: int | None = None,
    require_small_last_table: bool = False,
) -> dict[str, object]:
    per_file: list[dict[str, object]] = []
    for doc in docs:
        seen_table = False
        consecutive_empty = 0
        last_table_rows = None
        exit_page = None
        skipped_pages: list[dict[str, object]] = []
        page_info = doc["page_info"]
        page_count = int(doc["pages"])

        for next_index, page in enumerate(page_info, start=1):
            progress = (next_index - 1) / page_count
            remaining_pages = page_count - next_index + 1
            small_last_table_ok = (last_table_rows or 0) <= 3 if require_small_last_table else True
            remaining_ok = (
                True if max_remaining_pages is None else remaining_pages <= max_remaining_pages
            )
            if (
                seen_table
                and consecutive_empty >= empty_threshold
                and progress >= min_progress
                and remaining_ok
                and small_last_table_ok
            ):
                exit_page = next_index
                skipped_pages = page_info[next_index - 1 :]
                break

            if page["has_table"]:
                seen_table = True
                consecutive_empty = 0
                last_table_rows = int(page["rows"])
            elif seen_table:
                consecutive_empty += 1

        per_file.append(
            {
                "file": doc["file"],
                "exit_page": exit_page,
                "skipped_pages": len(skipped_pages),
                "skipped_table_pages": sum(1 for page in skipped_pages if page["has_table"]),
                "saved_seconds": round(
                    sum(float(page["extract_seconds"]) for page in skipped_pages), 6
                ),
            }
        )

    return {
        "rule": name,
        "files_with_early_exit": sum(1 for row in per_file if row["exit_page"] is not None),
        "files_with_wrong_skip": [row["file"] for row in per_file if row["skipped_table_pages"]],
        "total_skipped_pages": sum(row["skipped_pages"] for row in per_file),
        "total_saved_seconds": round(sum(row["saved_seconds"] for row in per_file), 6),
        "mean_saved_seconds_per_file": round(
            sum(row["saved_seconds"] for row in per_file) / len(per_file), 6
        ),
        "max_saved_seconds_single_file": round(
            max(row["saved_seconds"] for row in per_file),
            6,
        ),
        "sample": per_file[:5],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze table/empty-page patterns in PDFs.")
    parser.add_argument(
        "pdf_root",
        nargs="?",
        default="testing_files",
        type=Path,
        help="Directory containing representative PDFs",
    )
    args = parser.parse_args()

    docs = [analyze_pdf(path) for path in sorted(args.pdf_root.glob("*.pdf"))]
    if not docs:
        raise SystemExit("No PDF files found.")

    result = {
        "file_count": len(docs),
        "resume_after_gap_files": [doc["file"] for doc in docs if doc["tables_resume_after_gap"]],
        "trailing_gap_distribution": sorted({doc["trailing_gap"] for doc in docs}),
        "last_table_row_distribution": sorted(
            {doc["last_table_rows"] for doc in docs if doc["last_table_rows"] is not None}
        ),
        "candidate_rules": [
            simulate_rule(docs, "after_2_empty_after_table", empty_threshold=2),
            simulate_rule(docs, "after_3_empty_after_table", empty_threshold=3),
            simulate_rule(
                docs,
                "after_3_empty_after_table_and_small_final_table",
                empty_threshold=3,
                require_small_last_table=True,
            ),
            simulate_rule(
                docs,
                "after_4_empty_after_table_and_at_most_2_pages_left",
                empty_threshold=4,
                max_remaining_pages=2,
            ),
        ],
        "docs": docs,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
