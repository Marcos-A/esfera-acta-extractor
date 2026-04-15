#!/usr/bin/env python3
"""
Research harness for side-by-side alternative PDF table extractor evaluation.

This script keeps all prototype work off the live app path. It compares the current
pdfplumber-based extractor against a small set of practical alternatives on the real
local corpus and reports both speed and output compatibility.
"""

from __future__ import annotations

import argparse
import json
import re
import statistics
import time
import warnings
from collections import Counter
from pathlib import Path
from typing import Callable

import camelot
import fitz
import pandas as pd

from src.pdf_processor import extract_group_code_and_tables


warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning, module=r"camelot(\.|$)")


CandidateFn = Callable[[Path], list[pd.DataFrame]]


def dataframe_from_rows(rows: list[list[object]]) -> pd.DataFrame:
    if not rows:
        return pd.DataFrame()
    headers = ["" if value is None else str(value) for value in rows[0]]
    body = [[("" if value is None else str(value)) for value in row] for row in rows[1:]]
    return pd.DataFrame(body, columns=headers)


def choose_largest_table(dfs: list[pd.DataFrame]) -> pd.DataFrame | None:
    if not dfs:
        return None
    return max(dfs, key=lambda df: (df.shape[0] * max(df.shape[1], 1), df.shape[0], df.shape[1]))


def extract_reference(pdf_path: Path) -> tuple[str, list[pd.DataFrame]]:
    group_code, tables = extract_group_code_and_tables(str(pdf_path))
    return group_code, tables


def extract_with_pymupdf(pdf_path: Path) -> list[pd.DataFrame]:
    doc = fitz.open(pdf_path)
    try:
        tables: list[pd.DataFrame] = []
        for page in doc:
            found = page.find_tables()
            dfs = [dataframe_from_rows(table.extract()) for table in found.tables]
            chosen = choose_largest_table([df for df in dfs if not df.empty])
            if chosen is not None:
                tables.append(chosen)
        return tables
    finally:
        doc.close()


def extract_with_camelot(pdf_path: Path, flavor: str) -> list[pd.DataFrame]:
    results = camelot.read_pdf(str(pdf_path), pages="all", flavor=flavor)
    tables_by_page: dict[int, list[pd.DataFrame]] = {}
    for table in results:
        page_number = int(table.parsing_report["page"])
        df = table.df.copy()
        if df.empty:
            continue
        df.columns = ["" if value is None else str(value) for value in df.iloc[0].tolist()]
        df = df.iloc[1:].reset_index(drop=True)
        tables_by_page.setdefault(page_number, []).append(df)

    chosen_tables: list[pd.DataFrame] = []
    for page_number in sorted(tables_by_page):
        chosen = choose_largest_table([df for df in tables_by_page[page_number] if not df.empty])
        if chosen is not None:
            chosen_tables.append(chosen)
    return chosen_tables


def stringify_df(df: pd.DataFrame) -> pd.DataFrame:
    return df.fillna("").astype(str)


def normalize_df(df: pd.DataFrame) -> pd.DataFrame:
    string_df = stringify_df(df)
    return string_df.map(lambda value: re.sub(r"\s+", " ", value).strip())


def compare_table_sets(reference: list[pd.DataFrame], candidate: list[pd.DataFrame]) -> tuple[str, str | None]:
    if len(reference) != len(candidate):
        return "unusable", f"table_count:{len(reference)}!={len(candidate)}"

    exact = True
    normalized = True
    for index, (ref_df, cand_df) in enumerate(zip(reference, candidate), start=1):
        ref_exact = stringify_df(ref_df)
        cand_exact = stringify_df(cand_df)
        if not ref_exact.equals(cand_exact):
            exact = False

        ref_norm = normalize_df(ref_df)
        cand_norm = normalize_df(cand_df)
        if not ref_norm.equals(cand_norm):
            normalized = False
            return "unusable", f"table_{index}_normalized_mismatch"

    if exact:
        return "exact", None
    if normalized:
        return "close", "normalized_match_only"
    return "unusable", "unknown"


def benchmark_candidate(
    pdfs: list[Path],
    reference: dict[str, dict[str, object]],
    name: str,
    extractor: CandidateFn,
) -> dict[str, object]:
    total_elapsed = 0.0
    exact_count = 0
    close_count = 0
    unusable_count = 0
    mismatch_patterns: Counter[str] = Counter()
    failures: list[dict[str, object]] = []
    per_file: list[dict[str, object]] = []

    for pdf_path in pdfs:
        started_at = time.perf_counter()
        try:
            tables = extractor(pdf_path)
            status, reason = compare_table_sets(reference[pdf_path.name]["tables"], tables)
            error = None
        except Exception as exc:  # pragma: no cover - research harness
            tables = []
            status = "unusable"
            reason = f"{type(exc).__name__}:{exc}"
            error = reason
        elapsed = time.perf_counter() - started_at
        total_elapsed += elapsed

        if status == "exact":
            exact_count += 1
        elif status == "close":
            close_count += 1
            mismatch_patterns[reason or "close"] += 1
        else:
            unusable_count += 1
            mismatch_patterns[reason or "unusable"] += 1
            failures.append(
                {
                    "file": pdf_path.name,
                    "reason": reason,
                    "error": error,
                    "reference_table_count": len(reference[pdf_path.name]["tables"]),
                    "candidate_table_count": len(tables),
                }
            )

        per_file.append(
            {
                "file": pdf_path.name,
                "elapsed_seconds": round(elapsed, 6),
                "status": status,
                "reason": reason,
                "reference_table_count": len(reference[pdf_path.name]["tables"]),
                "candidate_table_count": len(tables),
            }
        )

    return {
        "candidate": name,
        "total_elapsed_seconds": round(total_elapsed, 6),
        "success_rate": {
            "exact": exact_count,
            "close": close_count,
            "unusable": unusable_count,
        },
        "mismatch_patterns": dict(mismatch_patterns.most_common()),
        "per_file_summary": per_file,
        "failures": failures[:20],
    }


def run_benchmark(pdf_root: Path, candidate_names: list[str] | None = None) -> dict[str, object]:
    pdfs = sorted(pdf_root.glob("*.pdf"))
    if not pdfs:
        raise FileNotFoundError(f"No PDFs found in {pdf_root}")

    reference: dict[str, dict[str, object]] = {}
    reference_total = 0.0
    for pdf_path in pdfs:
        started_at = time.perf_counter()
        group_code, tables = extract_reference(pdf_path)
        elapsed = time.perf_counter() - started_at
        reference[pdf_path.name] = {
            "group_code": group_code,
            "tables": tables,
            "elapsed": elapsed,
        }
        reference_total += elapsed

    all_candidates: list[tuple[str, CandidateFn]] = [
        ("pymupdf_find_tables", extract_with_pymupdf),
        ("camelot_lattice", lambda path: extract_with_camelot(path, "lattice")),
        ("camelot_stream", lambda path: extract_with_camelot(path, "stream")),
    ]
    if candidate_names is None:
        candidates = all_candidates
    else:
        wanted = set(candidate_names)
        candidates = [item for item in all_candidates if item[0] in wanted]

    results = [
        benchmark_candidate(pdfs, reference, name, extractor)
        for name, extractor in candidates
    ]

    per_file_reference = [
        {
            "file": pdf_path.name,
            "group_code": reference[pdf_path.name]["group_code"],
            "table_count": len(reference[pdf_path.name]["tables"]),
            "elapsed_seconds": round(float(reference[pdf_path.name]["elapsed"]), 6),
        }
        for pdf_path in pdfs
    ]

    return {
        "pdf_count": len(pdfs),
        "reference": {
            "name": "pdfplumber_current",
            "total_elapsed_seconds": round(reference_total, 6),
            "per_file_summary": per_file_reference,
        },
        "candidates": results,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Benchmark alternative PDF table extractors.")
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
        choices=["pymupdf_find_tables", "camelot_lattice", "camelot_stream"],
        help="Optional subset of candidates to benchmark",
    )
    args = parser.parse_args()

    result = run_benchmark(args.pdf_root, candidate_names=args.candidates)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
