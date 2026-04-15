from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from src.excel_processor import export_excel_with_spacing


def test_export_excel_with_spacing_creates_expected_public_safe_workbook(tmp_path: Path, export_case) -> None:
    dataframe, mp_codes_with_em, mp_codes = export_case
    output_path = tmp_path / "grades.xlsx"

    export_excel_with_spacing(dataframe, str(output_path), mp_codes_with_em, mp_codes)

    workbook = load_workbook(output_path)
    worksheet = workbook.active

    assert worksheet.title == "Sheet1"
    assert worksheet.freeze_panes == "C2"
    assert worksheet["A1"].value == "#"
    assert worksheet["B1"].value == "ESTUDIANT"
    assert worksheet["C1"].value == "MP01_CF01_\n1RA"
    assert worksheet["D1"].value == "MP01_CF01_\n2RA"
    assert worksheet["E1"].value == "MP01 CENTRE"
    assert worksheet["F1"].value == "MP01 EMPRESA"
    assert worksheet["G1"].value == "MP01"
    assert worksheet["H1"].value == "MP02_CF02_\n1RA"
    assert worksheet["I1"].value == "MP02"

    assert worksheet["A2"].value == 1
    assert worksheet["B2"].value == "Alice Example"
    assert worksheet["F2"].value == 8.5
    assert worksheet["G2"].value == 8
    assert worksheet["H2"].value in ("", None)
    assert worksheet["I2"].value == 4

    assert worksheet["B5"].value == "MP amb estada a l'empresa"
    assert worksheet["B6"].value == "MP sense estada a l'empresa"
    assert len(worksheet.conditional_formatting) >= 7


def test_export_excel_output_can_be_reopened_after_write(tmp_path: Path, export_case) -> None:
    dataframe, mp_codes_with_em, mp_codes = export_case
    output_path = tmp_path / "roundtrip.xlsx"

    export_excel_with_spacing(dataframe, str(output_path), mp_codes_with_em, mp_codes)

    reopened = load_workbook(output_path, data_only=False)
    assert reopened.active.max_column == 9
    assert reopened.active.max_row >= 6
