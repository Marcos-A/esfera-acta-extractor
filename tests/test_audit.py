from __future__ import annotations

from pathlib import Path

from src.audit import AuditStore


def test_audit_store_tracks_jobs_progress_and_summary(tmp_path: Path) -> None:
    store = AuditStore(tmp_path / "audit.sqlite3")

    store.create_job(
        job_id="job-1",
        request_type="pdf",
        source_name="sample.pdf",
        source_size_bytes=123,
        source_file_count=1,
        remote_addr="127.0.0.1",
    )
    store.mark_started("job-1")
    store.update_progress(
        "job-1",
        stage="processing",
        message="Working",
        progress_current=0,
        progress_total=1,
        metadata={"step": "unit-test"},
    )
    store.record_file_result(
        job_id="job-1",
        source_name="sample.pdf",
        output_name="sample.xlsx",
        status="success",
    )
    store.mark_success("job-1", returned_name="sample.xlsx", returned_file_count=1)

    job = store.get_job("job-1")
    summary = store.get_summary()
    recent_files = store.list_recent_files(limit=10)

    assert job is not None
    assert job["status"] == "success"
    assert job["stage"] == "processing"
    assert summary["total_jobs"] == 1
    assert summary["success_jobs"] == 1
    assert summary["total_files"] == 1
    assert recent_files[0]["output_name"] == "sample.xlsx"


def test_audit_store_lists_error_jobs_oldest_first(tmp_path: Path) -> None:
    store = AuditStore(tmp_path / "audit.sqlite3")
    for job_id in ("job-1", "job-2"):
        store.create_job(
            job_id=job_id,
            request_type="pdf",
            source_name=f"{job_id}.pdf",
            source_size_bytes=10,
            source_file_count=1,
            remote_addr=None,
        )
        store.mark_error(job_id, error_message="boom", debug_path=f"/tmp/{job_id}", metadata={})

    error_jobs = store.list_error_jobs()

    assert [job["job_id"] for job in error_jobs] == ["job-1", "job-2"]
