#!/usr/bin/env python3
"""
Benchmark the PDF extraction hot path on one or more representative Esfer@ PDFs.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.pdf_processor import extract_group_code_and_tables


def benchmark(pdf_paths: list[Path], runs: int) -> list[dict[str, object]]:
    results: list[dict[str, object]] = []
    for pdf_path in pdf_paths:
        durations: list[float] = []
        last_group_code = "unknown_group"
        last_table_count = 0
        last_row_count = 0
        for _ in range(runs):
            started_at = time.perf_counter()
            group_code, tables = extract_group_code_and_tables(str(pdf_path))
            durations.append(time.perf_counter() - started_at)
            last_group_code = group_code
            last_table_count = len(tables)
            last_row_count = sum(len(table) for table in tables)

        results.append(
            {
                "file": pdf_path.name,
                "runs": runs,
                "durations_seconds": [round(value, 6) for value in durations],
                "average_seconds": round(sum(durations) / len(durations), 6),
                "group_code": last_group_code,
                "table_count": last_table_count,
                "row_count": last_row_count,
            }
        )
    return results


def main() -> None:
    parser = argparse.ArgumentParser(description="Benchmark PDF extraction timings.")
    parser.add_argument("pdfs", nargs="+", type=Path, help="PDF files to benchmark")
    parser.add_argument("--runs", type=int, default=3, help="Number of runs per PDF")
    args = parser.parse_args()

    results = benchmark(args.pdfs, runs=args.runs)
    print(json.dumps(results, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
