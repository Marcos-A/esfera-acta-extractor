from __future__ import annotations

import re

import pandas as pd

from src.grade_processor import extract_mp_codes, extract_records, find_mp_codes_with_em, sort_records


def test_extract_records_converts_numeric_grades_but_keeps_status_codes() -> None:
    melted = pd.DataFrame(
        {
            "Nom i cognoms": ["Alice Example", "Bob Example"],
            "entry": [
                "MP01_CF01_1RA (1) - A7\nMP01_CF01_2RA (1) - PDT",
                "MP02_CF02_1RA (1) - 8",
            ],
        }
    )
    pattern = re.compile(
        r"(?P<code>[A-Za-z0-9]{3,5}_[A-Za-z0-9 ]{4,5}_\dRA)\s\(\d\)\s*-\s*(?P<grade>A\d{1,2}|PDT|EP|NA|\d+)",
        flags=re.IGNORECASE,
    )

    records = extract_records(melted, "Nom i cognoms", pattern)

    assert records.to_dict("records") == [
        {"estudiant": "Alice Example", "code": "MP01_CF01_1RA", "grade": 7},
        {"estudiant": "Alice Example", "code": "MP01_CF01_2RA", "grade": "PDT"},
        {"estudiant": "Bob Example", "code": "MP02_CF02_1RA", "grade": 8},
    ]


def test_extract_mp_codes_and_find_mp_codes_with_em_identify_module_structure() -> None:
    ra_records = pd.DataFrame(
        [
            {"estudiant": "Alice Example", "code": "MP01_CF01_1RA", "grade": 7},
            {"estudiant": "Alice Example", "code": "MP02_CF02_1RA", "grade": 8},
        ]
    )
    melted = pd.DataFrame(
        {
            "entry": [
                "MP01_CF01_12EM (1) - A8",
                "no EM here",
                "MP01_CF01_12EM (1) - A7",
            ]
        }
    )

    mp_codes = extract_mp_codes(ra_records)
    mp_codes_with_em = find_mp_codes_with_em(melted, mp_codes)

    assert mp_codes == ["MP01", "MP02"]
    assert mp_codes_with_em == ["MP01"]


def test_sort_records_normalizes_wrapped_names_and_sorts_case_insensitively() -> None:
    records = pd.DataFrame(
        [
            {"estudiant": "bob\nexample", "code": "mp02", "grade": 5},
            {"estudiant": "Alice Example", "code": "MP01", "grade": 7},
            {"estudiant": "bob\nexample", "code": "MP01", "grade": 6},
        ]
    )

    sorted_records = sort_records(records)

    assert sorted_records["estudiant"].tolist() == ["Alice Example", "bob example", "bob example"]
    assert sorted_records["code"].tolist() == ["MP01", "MP01", "mp02"]
