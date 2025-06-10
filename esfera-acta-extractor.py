"""
PDF Grade Extractor - Educational Grade Report Parser

This script extracts student grades from multiple PDF reports by:
1. Scanning input directory for PDF files
2. Processing each PDF file individually
3. Exporting results to separate Excel files in output directory

Dependencies:
- pdfplumber: For extracting tables from PDFs
- pandas: For data manipulation
- numpy: For numerical operations
- re: For pattern matching using regular expressions
- openpyxl: For Excel file generation
"""

import os
import re
import glob
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
    extract_records,
    extract_mp_codes,
    find_mp_codes_with_em,
    sort_records,
    export_excel_with_spacing
)
from src.excel_processor import export_excel_with_spacing
import glob


def process_pdf(pdf_path: str) -> None:
    """
    Process a single PDF file and generate the corresponding Excel output.
    """
    # Create output directory if it doesn't exist
    os.makedirs('02_extracted_data', exist_ok=True)
    
    # Extract group code for filename
    group_code = extract_group_code(pdf_path)
    output_xlsx = os.path.join('02_extracted_data', f'{group_code}.xlsx')
    
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
        (?P<code>[A-Za-z0-9]{3,5}                # MP code format
        _               
        [A-Za-z0-9]{4,5}                         # CF code format
        _                       
        \d(?:\s*\d)RA)                           # RA
        \s\(\d\)\s*-\s*                          # round (convocatòria)
        (?P<grade>A\d{1,2}|PDT|EP|NA)            # grade options: A#, PDT, EP, NA
        """,
        flags=re.IGNORECASE | re.VERBOSE
    )
    ra_records = extract_records(melted, name_col, ra_entry_pattern)

    # 10) Extract EM records using original entry_pattern
    em_entry_pattern = re.compile(
        r"""
        (?P<code>[A-Za-z0-9]{3,5}                # MP code format
        _               
        [A-Za-z0-9]{4,5}                         # CF code format
        _                       
        \d(?:\s*\d)EM)                           # EM
        \s\(\d\)\s*-\s*                          # round (convocatòria)
        (?P<grade>A\d{1,2}|PDT|EP|NA)            # grade options: A#, PDT, EP, NA
        """,
        flags=re.IGNORECASE | re.VERBOSE
    )
    em_records = extract_records(melted, name_col, em_entry_pattern)
    
    # 11) Extract MP records using entry pattern
    mp_entry_pattern = re.compile(
        r"""
        (?<!\S)                                   # must start at whitespace or BOF
        (?P<code>[A-Za-z0-9]{3,5}                 # MP code format
        _                                         # exactly one underscore
        [A-Za-z0-9]{4,5})                         # CF code format
        \s\(\d\)\s*-\s*                           # round (convocatòria)
        (?P<grade>A?\d{1,2}|PDT|EP|NA|PQ)         # grade options: A#, #, PDT, EP, NA, PQ
        (?!\S)                                    # must end at whitespace or EOF
        """,
        flags=re.VERBOSE | re.IGNORECASE
    )
    
    mp_records = extract_records(melted, name_col, mp_entry_pattern)
    # print(mp_records)
    # 12) Combine RA, EM and MP records
    combined_records = pd.concat([ra_records, em_records, mp_records], ignore_index=True)

    # 13) Get unique MP codes and identify those with EM entries
    mp_codes = extract_mp_codes(ra_records)
    mp_codes_with_em = find_mp_codes_with_em(melted, mp_codes)
    # 14) Sort records by student and code
    combined_records = sort_records(combined_records)
    # 15) Pivot to wide format: students × codes
    wide = combined_records.pivot(index='estudiant', columns='code', values='grade')
    # print(wide)
    # Convert numeric grades to float where possible, keep non-numeric grades as strings
    for col in wide.columns:
        if col != 'estudiant':
            # Convert to numeric where possible, keeping non-numeric values as strings
            # First, convert to string to handle all values consistently
            str_series = wide[col].astype(str)
            # Try to convert to numeric, keeping original strings where conversion fails
            numeric_series = pd.to_numeric(wide[col], errors='coerce')
            # Combine numeric and string values
            wide[col] = numeric_series.combine(str_series, 
                lambda x, y: x if pd.notna(x) else (y if y != 'nan' else ''))
    
    wide = wide.reset_index()

    # 16) Export to Excel with proper spacing between MP groups
    export_excel_with_spacing(wide, output_xlsx, mp_codes_with_em, mp_codes)


def main() -> None:
    """
    Main function to process all PDF files in the input directory.
    """    
    # Create output directory if it doesn't exist
    if not os.path.exists('02_extracted_data'):
        os.makedirs('02_extracted_data')
        
    pdf_files = glob.glob(os.path.join('01_source_pdfs', '*.pdf'))
    
    if not pdf_files:
        print("No PDF files found in '01_source_pdfs' directory.")
        return

    for pdf_file in pdf_files:
        print(f"\Extracting data from {pdf_file}...")
        try:
            process_pdf(pdf_file)
            print(f"\t- Successfully extracted data from {pdf_file}")
        except Exception as e:
            print(f"ERROR processing {pdf_file}: {str(e)}")


    # After processing all PDFs, generate summary reports
    summary_output_dir = '03_final_grade_summaries'
    if not os.path.exists(summary_output_dir):
        os.makedirs(summary_output_dir)

    # Get all xlsx files from the '02_extracted_data' directory
    source_files_pattern = os.path.join('02_extracted_data', '*.xlsx')
    all_potential_source_files = glob.glob(source_files_pattern)
    
    # Filter out temporary/owner files (starting with ~$)
    actual_source_xlsx_files = [f for f in all_potential_source_files if not os.path.basename(f).startswith('~$')]

    if not actual_source_xlsx_files:
        print("No valid processed XLSX files found in '02_extracted_data' to summarize.")
    else:
        print("\nGenerating summary reports...")
        # Import here to avoid circular dependency if summary_generator grows
        from src.summary_generator import generate_summary_report 
        for source_xlsx_file in actual_source_xlsx_files:
            summary_file_name = f"qualificacions_MP-{os.path.basename(source_xlsx_file)}"
            # summary_output_dir is '03_final_grade_summaries', defined above
            output_summary_path = os.path.join(summary_output_dir, summary_file_name)
            try:
                generate_summary_report(source_xlsx_file, output_summary_path)
                print(f"\t- Successfully generated summary: {output_summary_path}")
            except Exception as e:
                print(f"ERROR generating summary for {source_xlsx_file}: {str(e)}")

if __name__ == '__main__':
    main()