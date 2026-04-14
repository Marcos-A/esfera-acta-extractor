"""
PDF processing module for extracting and processing tables from Esfer@ grade reports.
"""

from __future__ import annotations

import time

import pdfplumber
import pandas as pd

from .perf_timing import TimingRecorder


DEFAULT_TABLE_OPTS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 3,
}

# The group code always lives near the top of the first page in the sample corpus, so
# extracting text from just the header band avoids parsing the entire page for metadata.
GROUP_CODE_REGION_HEIGHT = 220


def extract_tables(pdf_path: str, table_opts: dict = None) -> list[pd.DataFrame]:
    """
    Extract all tables from a PDF, one DataFrame per table found.

    The default pdfplumber settings are tuned for Esfer@ reports, which use visible
    table lines and benefit from a small snap tolerance.
    """
    if table_opts is None:
        table_opts = DEFAULT_TABLE_OPTS.copy()

    _group_code, tables = extract_group_code_and_tables(pdf_path, table_opts)
    return tables


def extract_group_code(pdf_path: str) -> str:
    """
    Extract the group code that appears under 'Codi del grup' in the first page.
    Returns the code with underscores instead of spaces.
    """
    group_code, _tables = extract_group_code_and_tables(pdf_path)
    return group_code


def extract_group_code_and_tables(
    pdf_path: str,
    table_opts: dict | None = None,
) -> tuple[str, list[pd.DataFrame]]:
    """
    Extract the group code and all tables in one PDF pass.

    The conversion pipeline needs both the first-page metadata and the page tables, so
    opening the PDF only once avoids duplicate parsing work.
    """
    if table_opts is None:
        table_opts = DEFAULT_TABLE_OPTS.copy()

    tables: list[pd.DataFrame] = []
    group_code = "unknown_group"
    page_count = 0
    table_page_count = 0
    extract_table_seconds = 0.0
    dataframe_build_seconds = 0.0
    timings = TimingRecorder("extract_group_code_and_tables")

    with timings.measure("open_pdf"):
        pdf = pdfplumber.open(pdf_path)

    try:
        page_count = len(pdf.pages)
        if pdf.pages:
            with timings.measure("extract_first_page_text"):
                group_code = _extract_group_code_from_page(pdf.pages[0])

        with timings.measure("extract_page_tables"):
            for page in pdf.pages:
                extract_started_at = time.perf_counter()
                raw = page.extract_table(table_opts)
                extract_table_seconds += time.perf_counter() - extract_started_at
                if not raw:
                    continue

                headers, *rows = raw
                dataframe_started_at = time.perf_counter()
                tables.append(pd.DataFrame(rows, columns=headers))
                dataframe_build_seconds += time.perf_counter() - dataframe_started_at
                table_page_count += 1
    finally:
        pdf.close()

    timings.log(
        pdf_path=pdf_path,
        page_count=page_count,
        table_pages=table_page_count,
        empty_pages=page_count - table_page_count,
        raw_table_extract_seconds=round(extract_table_seconds, 6),
        dataframe_build_seconds=round(dataframe_build_seconds, 6),
        group_code_region_height=GROUP_CODE_REGION_HEIGHT,
    )

    return group_code, tables


def _extract_group_code_from_text(text: str) -> str:
    """Extract the group code from the first-page text block."""
    lines = text.split('\n')
    for index, line in enumerate(lines):
        if 'Codi del grup' in line and index + 1 < len(lines):
            group_code = lines[index + 1].strip()
            return group_code.replace(' ', '_')
    return 'unknown_group'


def _extract_group_code_from_page(page: pdfplumber.page.Page) -> str:
    """Read only the top band of the first page, where Esfer@ prints the group code."""
    header_region = page.crop((0, 0, page.width, min(page.height, GROUP_CODE_REGION_HEIGHT)))
    return _extract_group_code_from_text(header_region.extract_text() or "")
