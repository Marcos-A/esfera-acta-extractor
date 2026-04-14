#!/usr/bin/env python3
"""
Benchmark a small set of pdfplumber table-extraction candidates on a real PDF corpus.

The goal is conservative comparison, not exhaustive search: each candidate is measured
against the current reference behavior and rejected if it changes any extracted table.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import pandas as pd
import pdfplumber

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.pdf_processor import DEFAULT_TABLE_OPTS, _extract_group_code_from_page


CANDIDATES: dict[str, dict[str, object]] = {
    "baseline": dict(DEFAULT_TABLE_OPTS),
    "lines_strict_both": {
        "vertical_strategy": "lines_strict",
        "horizontal_strategy": "lines_strict",
        "snap_tolerance": 3,
    },
    "lines_strict_horizontal": {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines_strict",
        "snap_tolerance": 3,
    },
    "edge_min_length_5": {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 3,
        "edge_min_length": 5,
    },
    "tolerances_2": {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 2,
        "join_tolerance": 2,
        "intersection_tolerance": 2,
    },
    "tolerances_2_edge5": {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 2,
        "join_tolerance": 2,
        "intersection_tolerance": 2,
        "edge_min_length": 5,
    },
}


def extract_with_settings(pdf_path: Path, table_settings: dict[str, object]) -> tuple[str, list[pd.DataFrame], float]:
    started_at = time.perf_counter()
    tables: list[pd.DataFrame] = []
    with pdfplumber.open(pdf_path) as pdf:
        group_code = _extract_group_code_from_page(pdf.pages[0]) if pdf.pages else "unknown_group"
        for page in pdf.pages:
            raw = page.extract_table(table_settings)
            if not raw:
                continue
            headers, *rows = raw
            tables.append(pd.DataFrame(rows, columns=headers))
    return group_code, tables, time.perf_counter() - started_at


def compare_tables(reference: list[pd.DataFrame], candidate: list[pd.DataFrame]) -> tuple[bool, str | None]:
    if len(reference) != len(candidate):
        return False, f"table_count:{len(reference)}!={len(candidate)}"
    for index, (ref_df, cand_df) in enumerate(zip(reference, candidate), start=1):
        if not ref_df.equals(cand_df):
            return False, f"table_{index}_mismatch"
    return True, None


def benchmark_corpus(pdf_root: Path, candidates: list[str]) -> dict[str, object]:
    pdfs = sorted(pdf_root.glob("*.pdf"))
    if not pdfs:
        raise FileNotFoundError(f"No PDFs found in {pdf_root}")

    reference_results: dict[str, dict[str, object]] = {}
    baseline_total = 0.0
    for pdf_path in pdfs:
        group_code, tables, elapsed = extract_with_settings(pdf_path, CANDIDATES["baseline"])
        reference_results[pdf_path.name] = {
            "group_code": group_code,
            "tables": tables,
            "elapsed": elapsed,
        }
        baseline_total += elapsed

    candidate_results: list[dict[str, object]] = []
    for name in candidates:
        settings = CANDIDATES[name]
        total_elapsed = 0.0
        mismatches: list[dict[str, object]] = []
        per_file: list[dict[str, object]] = []
        for pdf_path in pdfs:
            ref = reference_results[pdf_path.name]
            group_code, tables, elapsed = extract_with_settings(pdf_path, settings)
            total_elapsed += elapsed
            same_group = ref["group_code"] == group_code
            same_tables, reason = compare_tables(ref["tables"], tables)
            if not same_group or not same_tables:
                mismatches.append(
                    {
                        "file": pdf_path.name,
                        "reason": "group_code_mismatch" if not same_group else reason,
                        "reference_group_code": ref["group_code"],
                        "candidate_group_code": group_code,
                        "reference_table_count": len(ref["tables"]),
                        "candidate_table_count": len(tables),
                    }
                )
            per_file.append(
                {
                    "file": pdf_path.name,
                    "elapsed_seconds": round(elapsed, 6),
                    "delta_vs_baseline_seconds": round(ref["elapsed"] - elapsed, 6),
                }
            )

        delta = baseline_total - total_elapsed
        candidate_results.append(
            {
                "candidate": name,
                "settings": settings,
                "total_elapsed_seconds": round(total_elapsed, 6),
                "delta_vs_baseline_seconds": round(delta, 6),
                "delta_vs_baseline_pct": round((delta / baseline_total) * 100, 2),
                "mismatch_count": len(mismatches),
                "mismatches": mismatches[:20],
                "per_file_sample": per_file[:8],
            }
        )

    return {
        "pdf_count": len(pdfs),
        "baseline_total_seconds": round(baseline_total, 6),
        "candidates": candidate_results,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Benchmark a small set of pdfplumber candidates.")
    parser.add_argument(
        "pdf_root",
        nargs="?",
        default="testing_files",
        type=Path,
        help="Directory containing representative PDFs",
    )
    parser.add_argument(
        "--candidates",
        nargs="*",
        default=[name for name in CANDIDATES if name != "baseline"],
        choices=[name for name in CANDIDATES if name != "baseline"],
        help="Candidate names to benchmark",
    )
    args = parser.parse_args()

    result = benchmark_corpus(args.pdf_root, args.candidates)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
