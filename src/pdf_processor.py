"""
PDF processing module for extracting and processing tables from Esfer@ grade reports.
"""

import pdfplumber
import pandas as pd


def extract_tables(pdf_path: str, table_opts: dict = None) -> list[pd.DataFrame]:
    """
    Extract all tables from a PDF, one DataFrame per table found.

    The default pdfplumber settings are tuned for Esfer@ reports, which use visible
    table lines and benefit from a small snap tolerance.
    """
    if table_opts is None:
        table_opts = {
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "snap_tolerance": 3
        }

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
        table_opts = {
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "snap_tolerance": 3
        }

    tables: list[pd.DataFrame] = []
    group_code = 'unknown_group'

    with pdfplumber.open(pdf_path) as pdf:
        if pdf.pages:
            first_page_text = pdf.pages[0].extract_text() or ""
            group_code = _extract_group_code_from_text(first_page_text)

        for page in pdf.pages:
            raw = page.extract_table(table_opts)
            if not raw:
                continue
            headers, *rows = raw
            tables.append(pd.DataFrame(rows, columns=headers))

    return group_code, tables


def _extract_group_code_from_text(text: str) -> str:
    """Extract the group code from the first-page text block."""
    lines = text.split('\n')
    for index, line in enumerate(lines):
        if 'Codi del grup' in line and index + 1 < len(lines):
            group_code = lines[index + 1].strip()
            return group_code.replace(' ', '_')
    return 'unknown_group'
