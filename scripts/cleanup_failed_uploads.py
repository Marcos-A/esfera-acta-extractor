#!/usr/bin/env python3
"""
Delete retained failed-upload artifacts by age and optional storage limit.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import shutil
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]
AUDIT_MODULE_PATH = REPO_ROOT / "src" / "audit.py"
AUDIT_SPEC = importlib.util.spec_from_file_location("cleanup_audit", AUDIT_MODULE_PATH)
if AUDIT_SPEC is None or AUDIT_SPEC.loader is None:
    raise RuntimeError(f"Unable to load audit module from {AUDIT_MODULE_PATH}")
AUDIT_MODULE = importlib.util.module_from_spec(AUDIT_SPEC)
# Load AuditStore directly from the repository so this maintenance script can run
# without requiring the project to be installed as a Python package first.
AUDIT_SPEC.loader.exec_module(AUDIT_MODULE)
AuditStore = AUDIT_MODULE.AuditStore


def main() -> int:
    """Delete retained failed-upload artifacts based on age and optional size limits."""
    args = _parse_args()
    audit_store = AuditStore(args.audit_db)
    now = datetime.now(timezone.utc)
    jobs = audit_store.list_error_jobs()

    cleaned = 0
    for job in jobs:
        if _should_delete_by_age(job, now, args.retention_days):
            if _cleanup_job_artifacts(audit_store, job, reason=f"retention>{args.retention_days}d", dry_run=args.dry_run):
                cleaned += 1

    if args.max_size_mb is not None:
        max_bytes = args.max_size_mb * 1024 * 1024
        jobs = audit_store.list_error_jobs()
        while _folder_size_bytes(args.failure_root) > max_bytes:
            # Once the folder grows past the configured cap, remove the oldest retained
            # failures first so the newest debugging material stays available longer.
            next_job = _oldest_job_with_artifacts(jobs)
            if next_job is None:
                break
            if _cleanup_job_artifacts(audit_store, next_job, reason=f"size-limit>{args.max_size_mb}MB", dry_run=args.dry_run):
                cleaned += 1
            jobs = audit_store.list_error_jobs()

    print(f"Cleanup complete. Jobs cleaned: {cleaned}")
    print(f"Retained folder size: {_folder_size_bytes(args.failure_root)} bytes")
    return 0


def _parse_args() -> argparse.Namespace:
    """Parse command-line options and environment-based defaults."""
    parser = argparse.ArgumentParser(description="Clean retained failed uploads.")
    parser.add_argument(
        "--failure-root",
        default=os.getenv("FAILURE_ROOT", "./failed_uploads"),
        help="Root directory containing retained failed uploads.",
    )
    parser.add_argument(
        "--audit-db",
        default=os.getenv("AUDIT_DB_PATH", "./data/conversion_audit.sqlite3"),
        help="Path to the audit SQLite database.",
    )
    parser.add_argument(
        "--retention-days",
        type=int,
        default=int(os.getenv("FAILURE_RETENTION_DAYS", "30")),
        help="Delete retained artifacts older than this many days.",
    )
    parser.add_argument(
        "--max-size-mb",
        type=int,
        default=int(os.getenv("FAILURE_MAX_SIZE_MB")) if os.getenv("FAILURE_MAX_SIZE_MB") else None,
        help="Optional maximum size for retained artifacts. Oldest jobs are deleted until under this limit.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would be deleted without removing anything.",
    )
    args = parser.parse_args()
    args.failure_root = Path(args.failure_root)
    return args


def _should_delete_by_age(job: dict[str, Any], now: datetime, retention_days: int) -> bool:
    """Return True when a failed job is older than the configured retention window."""
    if retention_days < 0 or not _job_has_artifacts(job):
        return False
    reference = job.get("completed_at") or job.get("created_at")
    if not reference:
        return False
    return now - _parse_utc(reference) > timedelta(days=retention_days)


def _oldest_job_with_artifacts(jobs: list[dict[str, Any]]) -> dict[str, Any] | None:
    for job in jobs:
        if _job_has_artifacts(job):
            return job
    return None


def _job_has_artifacts(job: dict[str, Any]) -> bool:
    """Check both the main debug path and the per-file metadata paths stored in JSON."""
    metadata = _load_metadata(job.get("metadata_json"))
    debug_path = job.get("debug_path")
    if debug_path and Path(str(debug_path)).exists():
        return True
    for key in ("failed_source_path", "failure_log_path"):
        path_value = metadata.get(key)
        if path_value and Path(str(path_value)).exists():
            return True
    return False


def _cleanup_job_artifacts(
    audit_store: AuditStore,
    job: dict[str, Any],
    *,
    reason: str,
    dry_run: bool,
) -> bool:
    """Delete one job's retained artifacts and mirror that change back into the audit DB."""
    metadata = _load_metadata(job.get("metadata_json"))
    debug_path_value = job.get("debug_path")
    if not debug_path_value:
        debug_path_value = metadata.get("debug_path")

    if not debug_path_value:
        return False

    debug_path = Path(str(debug_path_value))
    if not debug_path.exists():
        metadata["failed_source_path"] = None
        metadata["failure_log_path"] = None
        metadata["artifacts_deleted_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        metadata["artifacts_deleted_reason"] = reason
        if not dry_run:
            audit_store.update_job_artifact_metadata(str(job["job_id"]), debug_path=None, metadata=metadata)
        return True

    print(f"{'Would delete' if dry_run else 'Deleting'} retained artifacts for job {job['job_id']} at {debug_path}")
    if not dry_run:
        if debug_path.is_dir():
            shutil.rmtree(debug_path, ignore_errors=False)
        else:
            debug_path.unlink(missing_ok=True)
        metadata["failed_source_path"] = None
        metadata["failure_log_path"] = None
        metadata["artifacts_deleted_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
        metadata["artifacts_deleted_reason"] = reason
        audit_store.update_job_artifact_metadata(str(job["job_id"]), debug_path=None, metadata=metadata)
    return True


def _load_metadata(raw_metadata: Any) -> dict[str, Any]:
    """Safely decode metadata blobs that may be empty or malformed."""
    if not raw_metadata:
        return {}
    try:
        parsed = json.loads(raw_metadata)
    except json.JSONDecodeError:
        return {}
    return parsed if isinstance(parsed, dict) else {}


def _parse_utc(value: str) -> datetime:
    parsed = datetime.fromisoformat(value)
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def _folder_size_bytes(path: Path) -> int:
    """Measure disk usage recursively for the retained-failures folder."""
    if not path.exists():
        return 0
    total = 0
    for entry in path.rglob("*"):
        if entry.is_file():
            total += entry.stat().st_size
    return total


if __name__ == "__main__":
    raise SystemExit(main())
