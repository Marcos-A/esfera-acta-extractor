"""
Public package exports for the Esfer@ conversion pipeline.

Keeping the main helpers here makes both the CLI script and any future integrations
import from a single stable location.
"""

from .pdf_processor import extract_tables, extract_group_code, extract_group_code_and_tables
from .data_processor import (
    normalize_headers,
    drop_irrelevant_columns,
    forward_fill_names,
    join_nonempty,
    select_melt_code_conv_grades,
    clean_entries
)
from .grade_processor import (
    extract_records,
    extract_mp_codes,
    find_mp_codes_with_em,
    sort_records
)
from .excel_processor import export_excel_with_spacing
from .conversion_service import (
    build_zip_from_artifacts,
    convert_input_directory,
    convert_pdf_collection,
    convert_pdf_to_excel,
    extract_zip_to_temp,
)

__all__ = [
    'extract_tables',
    'extract_group_code',
    'extract_group_code_and_tables',
    'normalize_headers',
    'drop_irrelevant_columns',
    'forward_fill_names',
    'join_nonempty',
    'select_melt_code_conv_grades',
    'clean_entries',
    'extract_records',
    'extract_mp_codes',
    'find_mp_codes_with_em',
    'sort_records',
    'export_excel_with_spacing',
    'convert_pdf_to_excel',
    'convert_pdf_collection',
    'convert_input_directory',
    'extract_zip_to_temp',
    'build_zip_from_artifacts',
]
