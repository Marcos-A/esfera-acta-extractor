from __future__ import annotations

import re

import pandas as pd

from src.data_processor import (
    clean_entries,
    drop_irrelevant_columns,
    forward_fill_names,
    join_nonempty,
    normalize_headers,
    select_melt_code_conv_grades,
)


def test_normalize_headers_collapses_line_breaks() -> None:
    dataframe = pd.DataFrame(columns=["Nom\ni cognoms ", " Codi (Conv)\n- Qual "])

    normalized = normalize_headers(dataframe)

    assert list(normalized.columns) == ["Nom i cognoms", "Codi (Conv) - Qual"]


def test_drop_irrelevant_columns_removes_base_and_duplicated_headers() -> None:
    dataframe = pd.DataFrame(columns=["Nom", "MC", "MC.1", "Pas de curs", "Qualificació"])

    filtered = drop_irrelevant_columns(dataframe, ["MC", "Pas de curs"])

    assert list(filtered.columns) == ["Nom", "Qualificació"]


def test_forward_fill_names_and_join_nonempty_rebuild_split_student_rows() -> None:
    dataframe = pd.DataFrame(
        {
            "Nom i cognoms": ["Alice Example", "", "Bob Example"],
            "Codi (Conv) - Qual": ["MP01_CF01_1RA (1) - A7", "MP01_CF01_2RA (1) - PDT", "MP02_CF02_1RA (1) - 8"],
        }
    )

    filled, name_column = forward_fill_names(dataframe, "nom i cognoms")
    merged = (
        filled.set_index(name_column)
        .groupby(level=0, sort=False)
        .agg(join_nonempty)
        .reset_index()
    )

    assert name_column == "Nom i cognoms"
    assert merged.loc[0, "Nom i cognoms"] == "Alice Example"
    assert merged.loc[0, "Codi (Conv) - Qual"] == "MP01_CF01_1RA (1) - A7\nMP01_CF01_2RA (1) - PDT"


def test_select_and_clean_entries_prepare_long_form_values_for_regex_extraction() -> None:
    dataframe = pd.DataFrame(
        {
            "Nom i cognoms": ["Alice Example"],
            "Codi (Conv) - Qual": ["MP01_CF01_1 RA   (1) - A7"],
            "No match": ["ignored"],
        }
    )

    melted = select_melt_code_conv_grades(
        dataframe,
        "Nom i cognoms",
        re.compile(r"^Codi \(Conv\) - Qual$", re.IGNORECASE),
    )
    melted["entry"] = clean_entries(melted["entry"])

    assert len(melted) == 1
    assert melted.iloc[0]["entry"] == "MP01_CF01_1RA (1) - A7"
