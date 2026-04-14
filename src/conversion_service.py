"""
Reusable conversion workflow for CLI and web entrypoints.
"""

from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .data_processor import (
    clean_entries,
    drop_irrelevant_columns,
    forward_fill_names,
    join_nonempty,
    normalize_headers,
    select_melt_code_conv_grades,
)
from .excel_processor import export_excel_with_spacing
from .perf_timing import TimingRecorder
from .grade_processor import (
    extract_mp_codes,
    extract_records,
    find_mp_codes_with_em,
    sort_records,
)
from .pdf_processor import extract_group_code, extract_tables


@dataclass
class ConversionArtifact:
    """Describes one generated file so callers can package or report it consistently."""
    source_name: str
    output_name: str
    output_path: Path


@dataclass
class ConversionResult:
    """Collects all files produced from a single request."""
    artifacts: list[ConversionArtifact]

def convert_pdf_to_excel(
    pdf_path: str | Path,
    output_dir: str | Path,
) -> Path:
    """
    Process one Esfer@ acta PDF and write its Excel workbook.

    This function owns the full parsing pipeline: table extraction, cleanup, record
    detection, reshaping, and final workbook formatting.
    """
    pdf_path = Path(pdf_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    timings = TimingRecorder("convert_pdf_to_excel")

    with timings.measure("extract_group_code"):
        group_code = extract_group_code(str(pdf_path))
        output_xlsx = output_dir / f"{group_code}.xlsx"

    with timings.measure("extract_tables"):
        tables = extract_tables(str(pdf_path))

    with timings.measure("shape_pdf_tables"):
        combined = pd.concat(tables, ignore_index=True)
        combined = normalize_headers(combined)
        combined = drop_irrelevant_columns(combined, ["MC", "H", "Pas de curs", "MH"])
        filled, name_col = forward_fill_names(combined, "nom i cognoms")
        merged = (
            # Some PDFs split a single student across multiple table rows. Grouping by the
            # filled name column reconstructs the original logical row before regex parsing.
            filled.set_index(name_col)
            .groupby(level=0, sort=False)
            .agg(join_nonempty)
            .reset_index()
        )

    code_pattern = re.compile(r"^Codi \(Conv\) - Qual(?:\.\d+)?$", re.IGNORECASE)
    melted = select_melt_code_conv_grades(merged, name_col, code_pattern)
    melted["entry"] = clean_entries(melted["entry"])

    ra_entry_pattern = re.compile(
        r"""
        (?P<code>[A-Za-z0-9]{3,5}
        _
        [A-Za-z0-9 ]{4,5}
        \s*_
        \d(?:\s*\d)RA)
        \s\(\d\)\s*-\s*
        (?P<grade>A\d{1,2}|PDT|EP|NA)?
        """,
        flags=re.IGNORECASE | re.VERBOSE,
    )
    em_entry_pattern = re.compile(
        r"""
        (?P<code>[A-Za-z0-9]{3,5}
        _
        [A-Za-z0-9 ]{4,5}
        _
        \d(?:\s*\d)EM)
        \s\(\d\)\s*-\s*
        (?P<grade>A\d{1,2}|PDT|EP|NA)?
        """,
        flags=re.IGNORECASE | re.VERBOSE,
    )
    mp_entry_pattern = re.compile(
        r"""
        (?<!\S)
        (?P<code>[A-Za-z0-9]{3,5}_
        [A-Za-z0-9 ]{4,5})
        \s*\(\d\)\s*-\s*
        (?P<grade>A?\d{1,2}|PDT|EP|NA|PQ)?
        (?!\S)
        """,
        flags=re.VERBOSE | re.IGNORECASE,
    )

    with timings.measure("extract_records"):
        ra_records = extract_records(melted, name_col, ra_entry_pattern)
        em_records = extract_records(melted, name_col, em_entry_pattern)
        mp_records = extract_records(melted, name_col, mp_entry_pattern)

    with timings.measure("shape_records"):
        combined_records = pd.concat([ra_records, em_records, mp_records], ignore_index=True)
        mp_codes = extract_mp_codes(ra_records)
        mp_codes_with_em = find_mp_codes_with_em(melted, mp_codes)
        combined_records = sort_records(combined_records)

        wide = combined_records.pivot(index="estudiant", columns="code", values="grade")
        for col in wide.columns:
            if col != "estudiant":
                # Excel should receive numbers as numbers so conditional formatting works,
                # but status markers like PDT or EP must remain as text.
                str_series = wide[col].astype(str)
                numeric_series = pd.to_numeric(wide[col], errors="coerce")
                wide[col] = numeric_series.combine(
                    str_series,
                    lambda x, y: x if pd.notna(x) else (y if y != "nan" else ""),
                )

        wide = wide.reset_index()

    with timings.measure("export_excel"):
        export_excel_with_spacing(
            wide,
            str(output_xlsx),
            mp_codes_with_em,
            mp_codes,
        )
    timings.log(
        pdf_path=pdf_path.name,
        output_path=output_xlsx.name,
        page_tables=len(tables),
        student_rows=len(wide),
        output_columns=len(wide.columns),
    )
    return output_xlsx


def convert_pdf_collection(
    pdf_paths: list[Path],
    output_dir: str | Path,
) -> ConversionResult:
    """Convert several PDFs and return metadata for every generated workbook."""
    timings = TimingRecorder("convert_pdf_collection")
    artifacts: list[ConversionArtifact] = []
    with timings.measure("convert_each_pdf"):
        for pdf_path in pdf_paths:
            output_path = convert_pdf_to_excel(pdf_path, output_dir)
            artifacts.append(
                ConversionArtifact(
                    source_name=pdf_path.name,
                    output_name=output_path.name,
                    output_path=output_path,
                )
            )
    timings.log(pdf_count=len(pdf_paths), artifact_count=len(artifacts))
    return ConversionResult(artifacts=artifacts)


def convert_input_directory(
    input_dir: str | Path,
    output_dir: str | Path,
) -> ConversionResult:
    """Batch-convert every PDF found in an input directory."""
    pdf_paths = sorted(Path(input_dir).glob("*.pdf"))
    if not pdf_paths:
        raise FileNotFoundError("No PDF files found in the input directory.")
    return convert_pdf_collection(pdf_paths, output_dir)


def extract_zip_to_temp(zip_path: str | Path) -> tuple[Path, list[Path]]:
    """
    Extract only PDFs from a ZIP into a temp directory.

    Hidden files and non-PDF members are ignored because users often upload ZIP files
    created by desktop tools that include extra metadata entries.
    """
    work_dir = Path(tempfile.mkdtemp(prefix="esfera-upload-"))
    pdf_paths: list[Path] = []
    with zipfile.ZipFile(zip_path) as archive:
        for member in archive.infolist():
            if member.is_dir():
                continue
            member_name = Path(member.filename)
            if member_name.name.startswith("."):
                continue
            if member_name.suffix.lower() != ".pdf":
                continue
            target_name = _dedupe_name(member_name.name, {path.name for path in pdf_paths})
            target_path = work_dir / target_name
            with archive.open(member) as source, target_path.open("wb") as destination:
                shutil.copyfileobj(source, destination)
            pdf_paths.append(target_path)
    if not pdf_paths:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise ValueError("The ZIP file does not contain any PDF files.")
    return work_dir, pdf_paths


def build_zip_from_artifacts(artifacts: list[ConversionArtifact], destination: str | Path) -> Path:
    """Package generated artifacts into a download ZIP, deduplicating file names if needed."""
    destination = Path(destination)
    archive_names: set[str] = set()
    timings = TimingRecorder("build_zip_from_artifacts")
    with timings.measure("zip_artifacts"):
        with zipfile.ZipFile(destination, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for artifact in artifacts:
                archive_name = _dedupe_name(artifact.output_name, archive_names)
                archive_names.add(archive_name)
                archive.write(artifact.output_path, arcname=archive_name)
    timings.log(artifact_count=len(artifacts), destination=destination.name)
    return destination


def cleanup_path(path: str | Path) -> None:
    """Delete a file or directory if it still exists."""
    target = Path(path)
    if target.is_dir():
        shutil.rmtree(target, ignore_errors=True)
    elif target.exists():
        target.unlink(missing_ok=True)


def _dedupe_name(name: str, existing_names: set[str]) -> str:
    """Keep archive members unique when different inputs would generate the same name."""
    candidate = name
    stem = Path(name).stem
    suffix = Path(name).suffix
    counter = 1
    while candidate in existing_names:
        candidate = f"{stem}-{counter}{suffix}"
        counter += 1
    return candidate
