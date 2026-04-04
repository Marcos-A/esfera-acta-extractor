"""
Data processing module for handling and transforming grade data.
"""

import re
import pandas as pd
import numpy as np


def normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Collapse multiline headers into single-line, trimmed names.

    PDF table extraction often keeps line breaks from the original layout, which makes
    later column matching brittle unless headers are normalized first.
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

    Some exports duplicate headers with suffixes such as ".1", so the regex removes
    both the original and duplicated versions.
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
    Locate the student-name column and forward-fill blank cells.

    Esfer@ tables often print the student name once and leave the following rows blank
    for related grades, so later grouping depends on copying the name downward.
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

    Keeping the original fragments together lets the regex stage extract multiple
    RA/EM/MP entries from one reconstructed text block.
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

    PDF extraction sometimes inserts spaces before "RA" or "EM", which would stop the
    grade-matching regexes from recognizing otherwise valid codes.
    """
    s = series.astype(str).str.replace(r"\s+", ' ', regex=True).str.strip()
    return s.str.replace(
        r'([A-Za-z0-9_]+)\s+(RA|EM)(?=\s*\(\d+\))',
        r'\1\2',
        regex=True
    ) 
