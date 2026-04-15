from __future__ import annotations

import zipfile
from pathlib import Path

from src.conversion_service import ConversionArtifact, build_zip_from_artifacts, cleanup_path, extract_zip_to_temp


def test_extract_zip_to_temp_ignores_non_pdfs_and_dedupes_names(tmp_path: Path) -> None:
    archive_path = tmp_path / "input.zip"
    with zipfile.ZipFile(archive_path, "w") as archive:
        archive.writestr(".DS_Store", "ignored")
        archive.writestr("grades.pdf", b"%PDF-1")
        archive.writestr("nested/grades.pdf", b"%PDF-2")
        archive.writestr("notes.txt", "ignored")

    work_dir, pdf_paths = extract_zip_to_temp(archive_path)
    try:
        names = sorted(path.name for path in pdf_paths)
        assert names == ["grades-1.pdf", "grades.pdf"]
        assert all(path.suffix.lower() == ".pdf" for path in pdf_paths)
    finally:
        cleanup_path(work_dir)


def test_build_zip_from_artifacts_keeps_duplicate_output_names_unique(tmp_path: Path) -> None:
    first = tmp_path / "first.xlsx"
    second = tmp_path / "second.xlsx"
    first.write_bytes(b"first")
    second.write_bytes(b"second")
    artifacts = [
        ConversionArtifact(source_name="a.pdf", output_name="report.xlsx", output_path=first),
        ConversionArtifact(source_name="b.pdf", output_name="report.xlsx", output_path=second),
    ]
    destination = tmp_path / "reports.zip"

    build_zip_from_artifacts(artifacts, destination)

    with zipfile.ZipFile(destination) as archive:
        assert sorted(archive.namelist()) == ["report-1.xlsx", "report.xlsx"]
        assert archive.getinfo("report.xlsx").compress_type == zipfile.ZIP_STORED


def test_cleanup_path_removes_files_and_directories(tmp_path: Path) -> None:
    file_path = tmp_path / "file.txt"
    dir_path = tmp_path / "dir"
    nested_file = dir_path / "nested.txt"
    file_path.write_text("hello", encoding="utf-8")
    dir_path.mkdir()
    nested_file.write_text("world", encoding="utf-8")

    cleanup_path(file_path)
    cleanup_path(dir_path)

    assert not file_path.exists()
    assert not dir_path.exists()
