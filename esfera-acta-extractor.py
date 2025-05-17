"""
PDF Grade Extractor - Educational Grade Report Parser

This script extracts student grades from PDF reports by:
1. Reading tables from a PDF file
2. Processing and cleaning the data
3. Extracting specific grade information
4. Outputting a clean Excel file with student names and grades

Dependencies:
- pdfplumber: For extracting tables from PDFs
- pandas: For data manipulation
- numpy: For numerical operations
- re: For pattern matching using regular expressions
- openpyxl: For Excel file generation
"""

import re
import pandas as pd
from src import (
    extract_tables,
    extract_group_code,
    normalize_headers,
    drop_irrelevant_columns,
    forward_fill_names,
    join_nonempty,
    select_melt_code_conv_grades,
    clean_entries,
    extract_ra_records,
    extract_mp_codes,
    find_mp_codes_with_em,
    sort_records,
    export_excel_with_spacing
)


def main(pdf_path: str) -> None:
    # Extract group code for filename
    group_code = extract_group_code(pdf_path)
    output_xlsx = f'{group_code}.xlsx'
    
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
    # 9) Extract RA records using original entry_pattern
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
    # 10) Get unique MP codes and identify those with EM entries
    mp_codes = extract_mp_codes(records)
    mp_codes_with_em = find_mp_codes_with_em(melted, mp_codes)
    # 11) Sort records by student and RA code
    records = sort_records(records)
    # 12) Pivot to wide format: students × RA codes
    wide = records.pivot(index='estudiant', columns='ra_code', values='grade')
    wide = wide.fillna('')                      # optional: blank instead of NaN
    wide = wide.reset_index()                   # make 'estudiant' a column again
    # 13) Export to Excel with proper spacing between MP groups
    export_excel_with_spacing(wide, output_xlsx, mp_codes_with_em, mp_codes)


if __name__ == '__main__':
    # Example usage
    PDF_FILE = 'ActaAvaluacioFlexible_Gestió Administrativa_1_1 GA-A ( CFPM AG10 )_2_263702.pdf'
    main(PDF_FILE)
