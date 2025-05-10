import re
import pdfplumber
import pandas as pd
import numpy as np

def debug_print(df, output_file):
    # Assign a path for the text file
    txt_path = output_file + '.txt'
    # Write each DataFrame to the text file
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write("Forward-filled DataFrame:\n")
        # to_string() gives a nicely formatted plain-text table
        f.write(df.to_string(index=False))
        f.write('\n')


def extract_tables(pdf_path, table_opts=None):
    """
    Open a PDF and extract tables from each page into a list of DataFrames.
    """
    table_opts = table_opts or {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "snap_tolerance": 3
    }
    dfs = []
    # Open the PDF file
    with pdfplumber.open(pdf_path) as pdf:
        # Iterate over pages and extract tables
        for page in pdf.pages:
            table = page.extract_table(table_opts)
            # Convert non-empty tables to DataFrames
            if table:
                df = pd.DataFrame(table[1:], columns=table[0])
                dfs.append(df)
    # Assign a path for the text file
    txt_path = 'extract_tables.txt'
    # Write each DataFrame to the text file
    with open(txt_path, 'w', encoding='utf-8') as f:
        for i, df in enumerate(dfs, start=1):
            f.write(f"Table {i}\n")
            # to_string() gives a nicely formatted plain-text table
            f.write(df.to_string(index=False))
            f.write('\n\n')
    return dfs

def normalize_headers(df):
    """
    Clean up column names by collapsing newlines into spaces.
    """
    df = df.copy()
    df.columns = (
        df.columns
          .str.replace(r"\s*\n\s*", ' ', regex=True)
          .str.strip()
    )
    return df


def drop_irrelevant_columns(df, irrelevant_columns):
    """
    Remove any columns whose base name appears in `irrelevant_columns`, 
    ignoring case and allowing for a .2, .3 suffix.
    """
    df = df.copy()
    # Obtain the list of columns
    cols = df.columns.tolist()
    # Escape each entry in the irrelevant columns list and join with |
    alts = "|".join(map(re.escape, irrelevant_columns))
    # Build the full pattern
    pattern = rf"^({alts})(?:\.\d+)?$"
    # Build a regex matching MC, H, Pas de curs or MH (with optional .2, .3 suffix)
    drop_re = re.compile(pattern, flags=re.IGNORECASE)
    # Find every column that matches
    to_drop = [c for c in cols if drop_re.match(c)]
    # Drop them all at once
    df = df.drop(columns=to_drop, errors='ignore')
    return df


def forward_fill_names(df, name_keyword):
    """
    Identify the column containing student names and forward-fill missing entries.
    """
    df = df.copy()
    # Find the name column by keyword
    name_col = next(c for c in df.columns if name_keyword.lower() in c.lower())
    # Replace empty strings with NaN and forward-fill
    df[name_col] = (
        df[name_col]
          .replace(r'^\s*$', np.nan, regex=True)
          .ffill()
    )
    debug_print(df, 'forward_fill_names')
    return df, name_col


def select_melt_code_conv_grades(df, name_col, code_pattern):
    """
    Keep only columns matching the code pattern and melt into long format.
    """
    # Identify columns to melt
    codi_conv_qual_cols = [col for col in df.columns if code_pattern.match(col)]
    # Melt into records
    melted = (
        df.melt(
            id_vars=[name_col],
            value_vars=codi_conv_qual_cols,
            var_name='column_header',
            value_name='entry'
        )
        .dropna(subset=['entry'])
    )
    debug_print(df, 'select_melt_codes')
    return melted


def remove_stray_marks(df):
    """
    Remove sequences of 2+ spaces + 2–3 digits + newline,
    e.g. '   250\n' or '  66\n', from each string in the series.
    """
    df = df.copy()
    pattern = r"\s{2,}\d{2,3}\n"
    for col in df.columns:
        if pd.api.types.is_string_dtype(df[col]):
            # cast to str (to handle NaN), then str.replace on the Series
            df[col] = (
                df[col]
                  .fillna("")       # avoid None / nan
                  .astype(str)
                  .str.replace(pattern, " ", regex=True)
            )
    return df


def clean_entries(series):
    """
    Normalize whitespace and glue RA/EM suffixes back onto codes.
    """
    # Collapse whitespace to single spaces
    s = (
        series.astype(str)
              .str.replace(r"\s+", ' ', regex=True)
              .str.strip()
    )
    # Glue RA/EM suffix back onto codes (e.g., "01 RA" -> "01RA")
    s = s.str.replace(
        r'([A-Za-z0-9_]+)\s+(RA|EM)(?=\s*\(\d+\)\s*-)',
        r'\1\2',
        regex=True
    )

    return s


def extract_records(melted, name_col, entry_pattern):
    """
    From each entry, extract code and grade pairs using regex.
    """
    rows = []
    for _, row in melted.iterrows():
        student = row[name_col]
        entry = row['entry']
        for code, grade in entry_pattern.findall(entry):
            rows.append({'student': student, 'ra_code': code, 'grade': grade})
    return pd.DataFrame(rows)


def sort_and_save(df, output_path):
    """
    Clean student names, sort records, reset index, and save to CSV.
    """
    # Normalize student names
    df['student'] = (
        df['student']
          .str.replace(r"\s*\n\s*", ' ', regex=True)
          .str.strip()
    )
    # Sort by student name and RA code (case-insensitive)
    df = df.sort_values(
        by=['student', 'ra_code'],
        key=lambda col: col.str.lower()
    )
    df = df.reset_index(drop=True)
    # Save to file
    df.to_csv(output_path, encoding='UTF-8', sep=';', index=False)
    print(f"✅ Extracted {len(df)} entries to {output_path}")


def main(pdf_path, output_csv):
    # 1. Extract tables from PDF
    raw_tables = extract_tables(pdf_path)

    # 2. Concatenate all DataFrames
    big_df = pd.concat(raw_tables, ignore_index=True)

    # 3. Normalize headers
    norm_df = normalize_headers(big_df)
    debug_print(norm_df, 'normalize_headers')

    # 4. Drop irrelevant columns
    irrelevant_columns = ['MC','H', 'Pas de curs', 'MH']
    relevant_df = drop_irrelevant_columns(norm_df, irrelevant_columns)

    # 7) Remove stray MC/H marks like "  66\n" or "   250\n"
    relevant_df = remove_stray_marks(relevant_df)

    # 5. Forward-fill student names
    df_filled, name_col = forward_fill_names(relevant_df, 'nom i cognoms')

    # 6. Melt only the Codi (Conv) - Qual columns
    code_pattern = re.compile(r'^Codi \(Conv\) - Qual$', re.IGNORECASE)
    melted = select_melt_code_conv_grades(df_filled, name_col, code_pattern)

    # 7. Clean up the entries
    melted['entry'] = clean_entries(melted['entry'])
    debug_print(melted, 'clean_melted')

    # 8. Extract code-grade records
    # entry_pattern = re.compile(
    # r"(?P<code>[A-Za-z0-9_]+)\s*\(\d+\)\s*-\s*(?P<grade>A\d{1,2}|PDT|EP|NA|PQ)",
    # flags=re.IGNORECASE
    # )
    # entry_pattern = re.compile(
    #     r"""
    #     (?P<code>                   # capture the RA code
    #     [A-Za-z0-9]{4,}               #   subject code (≥4 alnum chars)
    #     _[A-Za-z0-9]{4,5}             #   underscore + studies code (4–5 alnum chars)
    #     _\d{2}RA                      #   underscore + two digits + "RA"
    #     )
    #     \s                           # exactly one space
    #     \(\d\)                       # a single digit in parentheses
    #     \s*-\s*                      # space(s), dash, space(s)
    #     (?P<grade>                 # capture the grade
    #     A\d{1,2}                     #   "A" + 1–2 digits (e.g. A9 or A10)
    #     |PDT                         #   or "PDT"
    #     )
    #     """,
    #     flags=re.IGNORECASE | re.VERBOSE
    # )
    entry_pattern = re.compile(r"""
        (?P<code>
        [A-Za-z0-9]{4,}            # subject
        _[A-Za-z0-9]{4,5}          # studies
        _\d(?:\s*\d)RA             # allow space between the two RA digits
        )
        \s\(\d\)\s*-\s*
        (?P<grade>A\d{1,2}|PDT)
    """, flags=re.VERBOSE|re.IGNORECASE)
    result = extract_records(melted, name_col, entry_pattern)
    # Collapse any stray spaces out of the code
    result["ra_code"] = result["ra_code"].str.replace(r'\s+', '', regex=True)
    debug_print(result, 'extract_records')

    # 9. Sort results and save
    sort_and_save(result, output_csv)


if __name__ == '__main__':
    # Example usage
    PDF_FILE = 'ActaAvaluacioFlexible_Gestió Administrativa_1_1 GA-A ( CFPM AG10 )_2_263702.pdf'
    OUTPUT_CSV = 'report_ra_marks.csv'
    main(PDF_FILE, OUTPUT_CSV)
