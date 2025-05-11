"""
PDF Grade Extractor - Educational Grade Report Parser

This script extracts student grades from PDF reports by:
1. Reading tables from a PDF file
2. Processing and cleaning the data
3. Extracting specific grade information
4. Outputting a clean CSV file with student names and grades

Dependencies:
- pdfplumber: For extracting tables from PDFs
- pandas: For data manipulation
- numpy: For numerical operations
- re: For pattern matching using regular expressions
"""

import re
import pdfplumber
import pandas as pd
import numpy as np


def debug_print(
    df: pd.DataFrame,
    label: str,
    export_txt: bool = True,
    print_df: bool = False
) -> None:
    """
    Helper for inspecting DataFrames during development.
    - df: DataFrame to inspect
    - label: identifier for the output file/name
    - export_txt: if True, save DataFrame as a .txt snapshot
    - print_df: if True, print DataFrame in markdown format
    """
    if export_txt:
        txt_file = f"{label}.txt"
        with open(txt_file, 'w', encoding='utf-8') as f:
            f.write(f"=== {label} ===\n")
            f.write(df.to_string(index=False))
            f.write("\n")
    if print_df:
        print(df.to_markdown(index=False))


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


def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Collapse multiline headers into single-line, trimmed names.
    """
    df = df.copy()
    df.columns = (
        df.columns
          .str.replace(r"\s*\n\s*", ' ', regex=True)
          .str.strip()
    )
    return df


def drop_irrelevant_columns(
    df: pd.DataFrame,
    irrelevant_columns: list[str]
) -> pd.DataFrame:
    """
    Drop columns named in irrelevant_columns (with optional .n suffix).
    """
    df = df.copy()
    pattern = re.compile(
        r"^(?:" + "|".join(map(re.escape, irrelevant_columns)) + r")(?:\.\d+)?$",
        flags=re.IGNORECASE
    )
    to_drop = [col for col in df.columns if pattern.match(col)]
    return df.drop(columns=to_drop, errors='ignore')


def forward_fill_names(
    df: pd.DataFrame,
    name_keyword: str
) -> tuple[pd.DataFrame, str]:
    """
    Locate the name column by keyword, blank->NaN, then forward-fill.
    Returns updated df and column name.
    """
    df = df.copy()
    name_col = next(
        col for col in df.columns if name_keyword.lower() in col.lower()
    )
    df[name_col] = (
        df[name_col]
          .replace(r'^\s*$', np.nan, regex=True)
          .ffill()
    )
    return df, name_col


def join_nonempty(series: pd.Series) -> str:
    """
    Join non-empty values in a group with newline separators.
    """
    return "\n".join(str(v).strip() for v in series.dropna() if str(v).strip())


def select_melt_code_conv_grades(
    df: pd.DataFrame,
    name_col: str,
    code_pattern: re.Pattern
) -> pd.DataFrame:
    """
    Melt only columns matching code_pattern into long form.
    """
    cols = [c for c in df.columns if code_pattern.match(c)]
    melted = df.melt(
        id_vars=[name_col],
        value_vars=cols,
        var_name='column_header',
        value_name='entry'
    ).dropna(subset=['entry'])
    return melted


def clean_entries(series: pd.Series) -> pd.Series:
    """
    Normalize whitespace and reattach RA/EM suffixes.
    """
    s = series.astype(str).str.replace(r"\s+", ' ', regex=True).str.strip()
    return s.str.replace(
        r'([A-Za-z0-9_]+)\s+(RA|EM)(?=\s*\(\d+\))',
        r'\1\2',
        regex=True
    )


def extract_records(
    melted: pd.DataFrame,
    name_col: str,
    entry_pattern: re.Pattern
) -> pd.DataFrame:
    """
    Extract code & grade pairs from each entry via regex.
    """
    rows = []
    for _, row in melted.iterrows():
        for code, grade in entry_pattern.findall(row['entry']):
            rows.append({
                'student': row[name_col],
                'ra_code': code,
                'grade': grade
            })
    df = pd.DataFrame(rows)
    df['ra_code'] = df['ra_code'].str.replace(r"\s+", '', regex=True)
    return df


def sort_and_save(df: pd.DataFrame, output_path: str) -> None:
    """
    Clean names, sort, and save to semicolon-delimited CSV.
    """
    df['student'] = df['student'].str.replace(r"\s*\n\s*", ' ', regex=True).str.strip()
    df = df.sort_values(
        by=['student', 'ra_code'],
        key=lambda c: c.str.lower()
    ).reset_index(drop=True)
    df.to_csv(output_path, sep=';', encoding='utf-8', index=False)
    print(f"✅ Extracted {len(df)} entries to {output_path}")


def main(pdf_path: str, output_csv: str) -> None:
    # 1) Extract tables
    tables = extract_tables(pdf_path)
    # 2) Combine
    combined = pd.concat(tables, ignore_index=True)
    # 3) Normalize headers
    combined = normalize_headers(combined)
    # 4) Drop irrelevant
    combined = drop_irrelevant_columns(combined, ['MC', 'H', 'Pas de curs', 'MH'])
    # 5) Forward-fill names
    filled, name_col = forward_fill_names(combined, 'nom i cognoms')
    # 6) Merge split rows
    merged = (
        filled.set_index(name_col)
              .groupby(level=0, sort=False)
              .agg(join_nonempty)
              .reset_index()
    )
    # 7) Melt only Codi (Conv) - Qual columns (include optional .n suffix):
    code_pattern = re.compile(r'^Codi \(Conv\) - Qual(?:\.\d+)?$', re.IGNORECASE)
    melted = select_melt_code_conv_grades(merged, name_col, code_pattern)
    # 8) Clean entry text
    melted['entry'] = clean_entries(melted['entry'])
    # 9) Extract records using original entry_pattern
    entry_pattern = re.compile(
        r"""
        (?P<code>[A-Za-z0-9]{4,}                # MP code format
        _               
        [A-Za-z0-9]{4,5}                        # CF code format
        _                       
        \d(?:\s*\d)RA)                          # RA
        \s\(\d\)\s*-\s*                         # round (convocatòria)
        (?P<grade>A\d{1,2}|PDT|EP)              # grade options: A#, PDT, EP
        """,
        flags=re.IGNORECASE | re.VERBOSE
    )
    records = extract_records(melted, name_col, entry_pattern)
    # 10) Sort and save
    sort_and_save(records, output_csv)


if __name__ == '__main__':
    # Example usage
    PDF_FILE = 'ActaAvaluacioFlexible_Gestió Administrativa_1_1 GA-A ( CFPM AG10 )_2_263702.pdf'
    OUTPUT_CSV = 'report_ra_marks.csv'
    main(PDF_FILE, OUTPUT_CSV)
