"""
Grade processing module for handling RA and EM records.
"""

import re
import pandas as pd


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