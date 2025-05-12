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


def extract_ra_records(
    melted: pd.DataFrame,
    name_col: str,
    entry_pattern: re.Pattern
) -> pd.DataFrame:
    """
    Extract RA code & grade pairs from each entry via regex.
    """
    rows = []
    for _, row in melted.iterrows():
        for code, grade in entry_pattern.findall(row['entry']):
            rows.append({
                'estudiant': row[name_col],
                'ra_code': code,
                'grade': grade
            })
    df = pd.DataFrame(rows)
    df['ra_code'] = df['ra_code'].str.replace(r"\s+", '', regex=True)
    return df


def extract_mp_codes(records: pd.DataFrame) -> list[str]:
    """
    Extract unique MP codes from RA codes.
    """
    mp_pattern = re.compile(r'^([A-Za-z0-9]+)_')
    mp_codes = records['ra_code'].str.extract(mp_pattern, expand=False)
    return sorted(mp_codes.unique().tolist())


def find_mp_codes_with_em(melted: pd.DataFrame, mp_codes: list[str]) -> list[str]:
    """
    Find which MP codes have associated EM entries
    (stops searching once all MP codes have been checked).
    """
    em_entry_pattern = re.compile(
        r"""
        (?P<code>[A-Za-z0-9]{4,}                # MP code format
        _               
        [A-Za-z0-9]{4,5}                        # CF code format
        _                       
        \d(?:\s*\d)EM)                          # EM
        \s\(\d\)                                # round (convocatòria)
        """,
        flags=re.IGNORECASE | re.VERBOSE
    )
    
    mp_with_em = set()
    mp_pattern = re.compile(r'^([A-Za-z0-9]+)_')
    
    # Process entries until we've found all possible MPs with EM
    for entry in melted['entry']:
        # Check if entry contains any EM pattern
        if 'EM' not in str(entry):
            continue
            
        # Extract MP codes from EM entries
        for code in em_entry_pattern.findall(str(entry)):
            mp_match = mp_pattern.match(code)
            if mp_match:
                mp_code = mp_match.group(1)
                if mp_code in mp_codes:
                    mp_with_em.add(mp_code)
                    
        # If we've found EM entries for all MP codes, we can stop
        if len(mp_with_em) == len(mp_codes):
            break
            
    return sorted(list(mp_with_em))


def sort_records(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean names, sort and save to semicolon-delimited CSV.
    """
    df['estudiant'] = df['estudiant'].str.replace(r"\s*\n\s*", ' ', regex=True).str.strip()
    df = df.sort_values(
        by=['estudiant', 'ra_code'],
        key=lambda c: c.str.lower()
    ).reset_index(drop=True)
    return df


def export_csv(df: pd.DataFrame, output_path: str) -> None:
    """
    Save DataFrame to CSV with semicolon delimiter.
    """
    df.to_csv(output_path, sep=';', encoding='utf-8', index=False)
    print(f"\t- Extracted {len(df)} entries to {output_path}")


def export_excel_with_spacing(
    df: pd.DataFrame,
    output_path: str,
    mp_codes_with_em: list[str],
    mp_codes: list[str]
) -> None:
    """
    Export DataFrame to Excel with specific column spacing after each MP's RAs.
    - 3 empty columns after MPs with EM (named: MP CENTRE, MP EMPRESA, MP)
    - 1 empty column after other MPs (named: MP)
    Uses pre-computed mp_codes list for efficient grouping.
    """
    # Get the current column order (excluding 'estudiant')
    ra_codes = [col for col in df.columns if col != 'estudiant']
    
    # Group RA codes by their MP using existing mp_codes
    mp_groups = {mp: [] for mp in mp_codes}  # Initialize with known MPs
    for ra_code in ra_codes:
        # Find which MP this RA belongs to
        mp_code = next(mp for mp in mp_codes if ra_code.startswith(mp + '_'))
        mp_groups[mp_code].append(ra_code)
    
    # Create new column order with spacing
    new_columns = ['estudiant']
    for mp_code in mp_codes:  # Use mp_codes to maintain consistent order
        # Add all RA codes for this MP
        new_columns.extend(mp_groups[mp_code])
        # Add empty columns with specific names based on whether MP has EM
        if mp_code in mp_codes_with_em:
            new_columns.extend([
                f'{mp_code} CENTRE',
                f'{mp_code} EMPRESA',
                f'{mp_code}'
            ])
        else:
            new_columns.append(f'{mp_code}')
    
    # Create new DataFrame with the desired column order
    export_df = pd.DataFrame(index=df.index)
    export_df['estudiant'] = df['estudiant']
    
    # Add RA columns with their values
    for col in ra_codes:
        export_df[col] = df[col]
    
    # Add empty spacing columns
    for col in new_columns:
        if col not in export_df.columns:
            export_df[col] = ''
    
    # Reorder columns
    export_df = export_df[new_columns]
    
    # Export to Excel
    output_path = output_path.replace('.csv', '.xlsx')
    export_df.to_excel(output_path, index=False)
    print(f"\t- Exported {len(export_df)} entries to {output_path}")


def main(pdf_path: str) -> None:
    # 1) Extract group code for filename
    group_code = extract_group_code(pdf_path)
    output_xlsx = f'{group_code}.xlsx'
    # 2) Extract tables
    tables = extract_tables(pdf_path)
    # 3) Combine
    combined = pd.concat(tables, ignore_index=True)
    # 4) Normalize headers
    combined = normalize_headers(combined)
    # 5) Drop irrelevant
    combined = drop_irrelevant_columns(combined, ['MC', 'H', 'Pas de curs', 'MH'])
    # 6) Forward-fill names
    filled, name_col = forward_fill_names(combined, 'nom i cognoms')
    # 7) Merge split rows
    merged = (
        filled.set_index(name_col)
              .groupby(level=0, sort=False)
              .agg(join_nonempty)
              .reset_index()
    )
    # 8) Melt only Codi (Conv) - Qual columns (include optional .n suffix):
    code_pattern = re.compile(r'^Codi \(Conv\) - Qual(?:\.\d+)?$', re.IGNORECASE)
    melted = select_melt_code_conv_grades(merged, name_col, code_pattern)
    # 9) Clean entry text
    melted['entry'] = clean_entries(melted['entry'])
    # 10) Extract RA records using original entry_pattern
    ra_entry_pattern = re.compile(
        r"""
        (?P<code>[A-Za-z0-9]{4,}                # MP code format
        _               
        [A-Za-z0-9]{4,5}                        # CF code format
        _                       
        \d(?:\s*\d)RA)                          # RA
        \s\(\d\)\s*-\s*                         # round (convocatòria)
        (?P<grade>A\d{1,2}|PDT|EP|NA)           # grade options: A#, PDT, EP, NA
        """,
        flags=re.IGNORECASE | re.VERBOSE
    )
    records = extract_ra_records(melted, name_col, ra_entry_pattern)
    # 11) Get unique MP codes and identify those with EM entries (qualificació de pràctiques en empresa)
    mp_codes = extract_mp_codes(records)
    mp_codes_with_em = find_mp_codes_with_em(melted, mp_codes)
    # 12) Sort records by student and RA code
    records = sort_records(records)
    # 13) Pivot to wide format: students × RA codes
    wide = records.pivot(index='estudiant', columns='ra_code', values='grade')
    wide = wide.fillna('')                      # optional: blank instead of NaN
    wide = wide.reset_index()                   # make 'estudiant' a column again
    # 14) Export to Excel with proper spacing between MP groups
    export_excel_with_spacing(wide, output_xlsx, mp_codes_with_em, mp_codes)


if __name__ == '__main__':
    # Example usage
    PDF_FILE = 'ActaAvaluacioFlexible_Gestió Administrativa_1_1 GA-A ( CFPM AG10 )_2_263702.pdf'
    main(PDF_FILE)
