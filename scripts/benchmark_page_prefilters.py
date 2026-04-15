#!/usr/bin/env python3
"""
Profile cheap page-structure signals and benchmark conservative prefilter candidates.

This helper keeps prefilter experiments reproducible without changing live extraction
behavior. It compares candidate skip rules against the current reference extractor on
the real local corpus.
"""

from __future__ import annotations

import argparse
import json
import statistics
import sys
import time
from pathlib import Path

import pandas as pd
import pdfplumber

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.pdf_processor import DEFAULT_TABLE_OPTS, _extract_group_code_from_page


RECT_THRESHOLD = 14
EDGE_THRESHOLD = 56


def summarize(values: list[float | int]) -> dict[str, float | int]:
    return {
        "min": min(values),
        "median": statistics.median(values),
        "max": max(values),
    }


def profile_structure(pdf_root: Path) -> dict[str, object]:
    rows: list[dict[str, object]] = []
    for path in sorted(pdf_root.glob("*.pdf")):
        with pdfplumber.open(path) as pdf:
            for index, page in enumerate(pdf.pages, start=1):
                rect_count = len(page.rects)
                objects = page.objects
                edge_count = len(page.edges)
                text_len = len(page.extract_text() or "")
                raw = page.extract_table(DEFAULT_TABLE_OPTS)
                rows.append(
                    {
                        "file": path.name,
                        "page": index,
                        "has_table": bool(raw),
                        "rect_count": rect_count,
                        "char_count": len(objects.get("char", [])),
                        "edge_count": edge_count,
                        "text_len": text_len,
                    }
                )

    table_pages = [row for row in rows if row["has_table"]]
    non_table_pages = [row for row in rows if not row["has_table"]]
    return {
        "page_count": len(rows),
        "table_pages": len(table_pages),
        "non_table_pages": len(non_table_pages),
        "rect_count": {
            "table": summarize([int(row["rect_count"]) for row in table_pages]),
            "non_table": summarize([int(row["rect_count"]) for row in non_table_pages]),
        },
        "edge_count": {
            "table": summarize([int(row["edge_count"]) for row in table_pages]),
            "non_table": summarize([int(row["edge_count"]) for row in non_table_pages]),
        },
        "char_count": {
            "table": summarize([int(row["char_count"]) for row in table_pages]),
            "non_table": summarize([int(row["char_count"]) for row in non_table_pages]),
        },
        "text_len": {
            "table": summarize([int(row["text_len"]) for row in table_pages]),
            "non_table": summarize([int(row["text_len"]) for row in non_table_pages]),
        },
    }


def extract_reference(path: Path) -> tuple[str, list[pd.DataFrame], float]:
    started_at = time.perf_counter()
    tables: list[pd.DataFrame] = []
    with pdfplumber.open(path) as pdf:
        group_code = _extract_group_code_from_page(pdf.pages[0]) if pdf.pages else "unknown_group"
        for page in pdf.pages:
            raw = page.extract_table(DEFAULT_TABLE_OPTS)
            if not raw:
                continue
            headers, *rows = raw
            tables.append(pd.DataFrame(rows, columns=headers))
    return group_code, tables, time.perf_counter() - started_at


def extract_candidate(path: Path, mode: str) -> tuple[str, list[pd.DataFrame], float, int]:
    started_at = time.perf_counter()
    skipped_pages = 0
    tables: list[pd.DataFrame] = []
    with pdfplumber.open(path) as pdf:
        group_code = _extract_group_code_from_page(pdf.pages[0]) if pdf.pages else "unknown_group"
        seen_table = False
        seen_empty_after_table = False
        page_count = len(pdf.pages)

        for index, page in enumerate(pdf.pages, start=1):
            if mode == "post_gap_rect" and seen_empty_after_table:
                if len(page.rects) <= RECT_THRESHOLD:
                    skipped_pages += 1
                    continue
            elif mode == "halfway_rect" and seen_table and index / page_count >= 0.5:
                if len(page.rects) <= RECT_THRESHOLD:
                    skipped_pages += 1
                    continue
            elif mode == "post_gap_edge" and seen_empty_after_table:
                if len(page.edges) <= EDGE_THRESHOLD:
                    skipped_pages += 1
                    continue

            raw = page.extract_table(DEFAULT_TABLE_OPTS)
            if not raw:
                if seen_table:
                    seen_empty_after_table = True
                continue

            seen_table = True
            headers, *rows = raw
            tables.append(pd.DataFrame(rows, columns=headers))

    return group_code, tables, time.perf_counter() - started_at, skipped_pages


def compare_tables(reference: list[pd.DataFrame], candidate: list[pd.DataFrame]) -> tuple[bool, str | None]:
    if len(reference) != len(candidate):
        return False, f"table_count:{len(reference)}!={len(candidate)}"
    for index, (ref_df, cand_df) in enumerate(zip(reference, candidate), start=1):
        if not ref_df.equals(cand_df):
            return False, f"table_{index}_mismatch"
    return True, None


def benchmark_candidates(pdf_root: Path) -> dict[str, object]:
    pdfs = sorted(pdf_root.glob("*.pdf"))
    reference: dict[str, dict[str, object]] = {}
    reference_total = 0.0
    for path in pdfs:
        group_code, tables, elapsed = extract_reference(path)
        reference[path.name] = {
            "group_code": group_code,
            "tables": tables,
            "elapsed": elapsed,
        }
        reference_total += elapsed

    candidates = ["post_gap_rect", "halfway_rect", "post_gap_edge"]
    results: list[dict[str, object]] = []
    for mode in candidates:
        total_elapsed = 0.0
        total_skipped_pages = 0
        mismatches: list[dict[str, object]] = []
        per_file: list[dict[str, object]] = []
        for path in pdfs:
            group_code, tables, elapsed, skipped_pages = extract_candidate(path, mode)
            ref = reference[path.name]
            total_elapsed += elapsed
            total_skipped_pages += skipped_pages

            same_group = group_code == ref["group_code"]
            same_tables, reason = compare_tables(ref["tables"], tables)
            if not same_group or not same_tables:
                mismatches.append(
                    {
                        "file": path.name,
                        "reason": "group_code_mismatch" if not same_group else reason,
                        "reference_group_code": ref["group_code"],
                        "candidate_group_code": group_code,
                        "reference_table_count": len(ref["tables"]),
                        "candidate_table_count": len(tables),
                    }
                )

            per_file.append(
                {
                    "file": path.name,
                    "elapsed_seconds": round(elapsed, 6),
                    "skipped_pages": skipped_pages,
                    "delta_vs_reference_seconds": round(ref["elapsed"] - elapsed, 6),
                }
            )

        delta = reference_total - total_elapsed
        results.append(
            {
                "candidate": mode,
                "total_elapsed_seconds": round(total_elapsed, 6),
                "delta_vs_reference_seconds": round(delta, 6),
                "delta_vs_reference_pct": round((delta / reference_total) * 100, 2),
                "skipped_pages": total_skipped_pages,
                "mismatch_count": len(mismatches),
                "mismatches": mismatches[:20],
                "per_file_sample": per_file[:8],
            }
        )

    return {
        "reference_total_seconds": round(reference_total, 6),
        "candidates": results,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Profile and benchmark structural page prefilters.")
    parser.add_argument(
        "pdf_root",
        nargs="?",
        default="testing_files",
        type=Path,
        help="Directory containing representative PDFs",
    )
    args = parser.parse_args()

    result = {
        "structure_profile": profile_structure(args.pdf_root),
        "prefilter_benchmark": benchmark_candidates(args.pdf_root),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
