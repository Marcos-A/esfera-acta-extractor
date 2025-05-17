"""
PDF processing module for extracting and processing tables from Esfer@ grade reports.
"""

import pdfplumber
import pandas as pd


def extract_tables(pdf_path: str, table_opts: dict = None) -> list[pd.DataFrame]:
    """
    Extract all tables from a PDF, one DataFrame per table found.
    """
    if table_opts is None:
        table_opts = {
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines",
            "snap_tolerance": 3
        }

    tables: list[pd.DataFrame] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            raw = page.extract_table(table_opts)
            if not raw:
                continue
            headers, *rows = raw
            tables.append(pd.DataFrame(rows, columns=headers))

    with open('extracted_tables.txt', 'w', encoding='utf-8') as summary:
        for i, df in enumerate(tables, start=1):
            summary.write(f"-- Table {i} --\n")
            summary.write(df.to_string(index=False))
            summary.write("\n\n")

    return tables


def extract_group_code(pdf_path: str) -> str:
    """
    Extract the group code that appears under 'Codi del grup' in the first page.
    Returns the code with underscores instead of spaces.
    """
    with pdfplumber.open(pdf_path) as pdf:
        first_page = pdf.pages[0]
        text = first_page.extract_text()
        
        # Find the line after "Codi del grup"
        lines = text.split('\n')
        for i, line in enumerate(lines):
            if 'Codi del grup' in line and i + 1 < len(lines):
                # Extract the code from the next line and replace spaces with underscores
                group_code = lines[i + 1].strip()
                return group_code.replace(' ', '_')
    
    return 'unknown_group'  # Fallback if code not found 