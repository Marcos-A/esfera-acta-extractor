#!/usr/bin/env python3
"""
Profile the internal phases of pdfplumber's table-finding pipeline on real PDFs.

This is a research-only helper. It reproduces the installed pdfplumber 0.10.x
extract_table flow stage by stage so we can time the geometry pipeline without
changing app or library code.
"""

from __future__ import annotations

import argparse
import itertools
import json
import statistics
import sys
import time
from operator import itemgetter
from pathlib import Path
from typing import Any, Optional

import pdfplumber
import pdfplumber.table as t

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from src.pdf_processor import DEFAULT_TABLE_OPTS


def summarize(values: list[float | int]) -> dict[str, float | int]:
    return {
        "min": round(min(values), 6) if values else 0,
        "median": round(statistics.median(values), 6) if values else 0,
        "max": round(max(values), 6) if values else 0,
    }


def timed_extract_text(table: t.Table, **kwargs: Any) -> tuple[list[list[Optional[str]]], dict[str, float | int]]:
    chars_started_at = time.perf_counter()
    chars = table.page.chars
    chars_seconds = time.perf_counter() - chars_started_at

    rows_started_at = time.perf_counter()
    rows = table.rows
    rows_seconds = time.perf_counter() - rows_started_at

    text_call_count = 0
    text_seconds = 0.0
    row_char_filter_seconds = 0.0
    cell_char_filter_seconds = 0.0
    table_arr: list[list[Optional[str]]] = []

    def char_in_bbox(char: dict[str, Any], bbox: tuple[float, float, float, float]) -> bool:
        v_mid = (char["top"] + char["bottom"]) / 2
        h_mid = (char["x0"] + char["x1"]) / 2
        x0, top, x1, bottom = bbox
        return bool((h_mid >= x0) and (h_mid < x1) and (v_mid >= top) and (v_mid < bottom))

    for row in rows:
        row_chars_started_at = time.perf_counter()
        row_chars = [char for char in chars if char_in_bbox(char, row.bbox)]
        row_char_filter_seconds += time.perf_counter() - row_chars_started_at

        arr: list[Optional[str]] = []
        for cell in row.cells:
            if cell is None:
                cell_text = None
            else:
                cell_chars_started_at = time.perf_counter()
                cell_chars = [char for char in row_chars if char_in_bbox(char, cell)]
                cell_char_filter_seconds += time.perf_counter() - cell_chars_started_at

                if cell_chars:
                    kwargs["x_shift"] = cell[0]
                    kwargs["y_shift"] = cell[1]
                    if "layout" in kwargs:
                        kwargs["layout_width"] = cell[2] - cell[0]
                        kwargs["layout_height"] = cell[3] - cell[1]
                    text_started_at = time.perf_counter()
                    cell_text = t.utils.extract_text(cell_chars, **kwargs)
                    text_seconds += time.perf_counter() - text_started_at
                    text_call_count += 1
                else:
                    cell_text = ""
            arr.append(cell_text)
        table_arr.append(arr)

    return table_arr, {
        "chars_access_seconds": chars_seconds,
        "rows_build_seconds": rows_seconds,
        "row_char_filter_seconds": row_char_filter_seconds,
        "cell_char_filter_seconds": cell_char_filter_seconds,
        "extract_text_seconds": text_seconds,
        "extract_text_calls": text_call_count,
    }


def profile_page(page: pdfplumber.page.Page, table_settings: dict[str, Any]) -> dict[str, Any]:
    tset = t.TableSettings.resolve(table_settings)
    phase: dict[str, float | int] = {}

    resolve_started_at = time.perf_counter()
    resolved = t.TableSettings.resolve(table_settings)
    phase["resolve_settings_seconds"] = time.perf_counter() - resolve_started_at

    page_edges_started_at = time.perf_counter()
    page_edges = page.edges
    phase["page_edges_access_seconds"] = time.perf_counter() - page_edges_started_at
    phase["page_edges_count"] = len(page_edges)

    v_base_started_at = time.perf_counter()
    v_base = t.utils.filter_edges(page_edges, "v")
    phase["vertical_filter_seconds"] = time.perf_counter() - v_base_started_at

    h_base_started_at = time.perf_counter()
    h_base = t.utils.filter_edges(page_edges, "h")
    phase["horizontal_filter_seconds"] = time.perf_counter() - h_base_started_at

    merge_started_at = time.perf_counter()
    merged_edges = t.merge_edges(
        list(v_base) + list(h_base),
        snap_x_tolerance=resolved.snap_x_tolerance,
        snap_y_tolerance=resolved.snap_y_tolerance,
        join_x_tolerance=resolved.join_x_tolerance,
        join_y_tolerance=resolved.join_y_tolerance,
    )
    phase["merge_edges_seconds"] = time.perf_counter() - merge_started_at
    phase["merged_edges_count"] = len(merged_edges)

    final_filter_started_at = time.perf_counter()
    edges = t.utils.filter_edges(merged_edges, min_length=resolved.edge_min_length)
    phase["final_filter_seconds"] = time.perf_counter() - final_filter_started_at
    phase["final_edges_count"] = len(edges)

    intersections_started_at = time.perf_counter()
    intersections = t.edges_to_intersections(
        edges,
        resolved.intersection_x_tolerance,
        resolved.intersection_y_tolerance,
    )
    phase["edges_to_intersections_seconds"] = time.perf_counter() - intersections_started_at
    phase["intersection_count"] = len(intersections)

    cells_started_at = time.perf_counter()
    cells = t.intersections_to_cells(intersections)
    phase["intersections_to_cells_seconds"] = time.perf_counter() - cells_started_at
    phase["cell_count"] = len(cells)

    tables_started_at = time.perf_counter()
    cell_groups = t.cells_to_tables(cells)
    tables = [t.Table(page, cell_group) for cell_group in cell_groups]
    phase["cells_to_tables_seconds"] = time.perf_counter() - tables_started_at
    phase["table_group_count"] = len(tables)

    select_started_at = time.perf_counter()
    if tables:
        largest = sorted(tables, key=lambda table: (-len(table.cells), table.bbox[1], table.bbox[0]))[0]
    else:
        largest = None
    phase["select_table_seconds"] = time.perf_counter() - select_started_at

    extract_result = None
    if largest is not None:
        extract_started_at = time.perf_counter()
        extract_result, extract_phase = timed_extract_text(largest, **(resolved.text_settings or {}))
        phase["table_extract_seconds"] = time.perf_counter() - extract_started_at
        phase.update(extract_phase)
        phase["selected_table_cell_count"] = len(largest.cells)
        phase["selected_table_row_count"] = len(largest.rows)
    else:
        phase["table_extract_seconds"] = 0.0
        phase["chars_access_seconds"] = 0.0
        phase["rows_build_seconds"] = 0.0
        phase["row_char_filter_seconds"] = 0.0
        phase["cell_char_filter_seconds"] = 0.0
        phase["extract_text_seconds"] = 0.0
        phase["extract_text_calls"] = 0
        phase["selected_table_cell_count"] = 0
        phase["selected_table_row_count"] = 0

    phase["has_table"] = largest is not None
    return {
        "result": extract_result,
        "phase": phase,
    }


def profile_corpus(pdf_root: Path, table_settings: dict[str, Any]) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    for path in sorted(pdf_root.glob("*.pdf")):
        with pdfplumber.open(path) as pdf:
            for index, page in enumerate(pdf.pages, start=1):
                profiled = profile_page(page, table_settings)
                rows.append(
                    {
                        "file": path.name,
                        "page": index,
                        **profiled["phase"],
                    }
                )

    table_pages = [row for row in rows if row["has_table"]]
    non_table_pages = [row for row in rows if not row["has_table"]]
    timed_fields = [
        "page_edges_access_seconds",
        "vertical_filter_seconds",
        "horizontal_filter_seconds",
        "merge_edges_seconds",
        "final_filter_seconds",
        "edges_to_intersections_seconds",
        "intersections_to_cells_seconds",
        "cells_to_tables_seconds",
        "select_table_seconds",
        "table_extract_seconds",
        "chars_access_seconds",
        "rows_build_seconds",
        "row_char_filter_seconds",
        "cell_char_filter_seconds",
        "extract_text_seconds",
    ]
    count_fields = [
        "page_edges_count",
        "merged_edges_count",
        "final_edges_count",
        "intersection_count",
        "cell_count",
        "table_group_count",
        "selected_table_cell_count",
        "selected_table_row_count",
        "extract_text_calls",
    ]

    summary = {
        "page_count": len(rows),
        "table_pages": len(table_pages),
        "non_table_pages": len(non_table_pages),
        "timings": {},
        "counts": {},
        "slowest_table_pages": sorted(
            table_pages,
            key=lambda row: row["page_edges_access_seconds"]
            + row["merge_edges_seconds"]
            + row["edges_to_intersections_seconds"]
            + row["intersections_to_cells_seconds"]
            + row["cells_to_tables_seconds"]
            + row["table_extract_seconds"],
            reverse=True,
        )[:10],
        "slowest_non_table_pages": sorted(
            non_table_pages,
            key=lambda row: row["page_edges_access_seconds"]
            + row["merge_edges_seconds"]
            + row["edges_to_intersections_seconds"]
            + row["intersections_to_cells_seconds"]
            + row["cells_to_tables_seconds"],
            reverse=True,
        )[:10],
    }

    for field in timed_fields:
        summary["timings"][field] = {
            "table": summarize([float(row[field]) for row in table_pages]),
            "non_table": summarize([float(row[field]) for row in non_table_pages]),
            "table_total": round(sum(float(row[field]) for row in table_pages), 6),
            "non_table_total": round(sum(float(row[field]) for row in non_table_pages), 6),
        }

    for field in count_fields:
        summary["counts"][field] = {
            "table": summarize([int(row[field]) for row in table_pages]),
            "non_table": summarize([int(row[field]) for row in non_table_pages]),
        }

    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="Profile pdfplumber table-finding internals.")
    parser.add_argument(
        "pdf_root",
        nargs="?",
        default="testing_files",
        type=Path,
        help="Directory containing representative PDFs",
    )
    args = parser.parse_args()

    result = profile_corpus(args.pdf_root, DEFAULT_TABLE_OPTS.copy())
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
