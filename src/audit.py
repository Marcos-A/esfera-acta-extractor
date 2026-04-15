"""
SQLite-backed audit logging for conversion jobs.
"""

from __future__ import annotations

import json
import os
import sqlite3
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


class AuditStore:
    """Persist job history so background conversions can be tracked outside request memory."""
    def __init__(self, db_path: str | os.PathLike[str]) -> None:
        self.db_path = str(db_path)
        Path(self.db_path).parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def create_job(
        self,
        *,
        job_id: str,
        request_type: str,
        source_name: str,
        source_size_bytes: int,
        source_file_count: int,
        remote_addr: str | None,
    ) -> None:
        self._execute(
            """
            INSERT INTO conversion_jobs (
                job_id,
                request_type,
                source_name,
                source_size_bytes,
                source_file_count,
                remote_addr,
                status,
                created_at,
                updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                job_id,
                request_type,
                source_name,
                source_size_bytes,
                source_file_count,
                remote_addr,
                "received",
                _utc_now(),
                _utc_now(),
            ),
        )

    def mark_started(self, job_id: str) -> None:
        self._update_status(job_id, "processing", started_at=_utc_now())

    def update_progress(
        self,
        job_id: str,
        *,
        stage: str,
        message: str,
        progress_current: int,
        progress_total: int,
        metadata: dict[str, Any] | None = None,
    ) -> None:
        self._execute(
            """
            UPDATE conversion_jobs
            SET stage = ?,
                progress_message = ?,
                progress_current = ?,
                progress_total = ?,
                metadata_json = ?,
                updated_at = ?
            WHERE job_id = ?
            """,
            (
                stage,
                message,
                progress_current,
                progress_total,
                json.dumps(metadata or {}, ensure_ascii=True),
                _utc_now(),
                job_id,
            ),
        )

    def update_source_file_count(self, job_id: str, source_file_count: int) -> None:
        self._execute(
            """
            UPDATE conversion_jobs
            SET source_file_count = ?,
                updated_at = ?
            WHERE job_id = ?
            """,
            (source_file_count, _utc_now(), job_id),
        )

    def mark_success(
        self,
        job_id: str,
        *,
        returned_name: str,
        returned_file_count: int,
    ) -> None:
        timestamp = _utc_now()
        self._execute(
            """
            UPDATE conversion_jobs
            SET status = ?,
                returned_name = ?,
                returned_file_count = ?,
                completed_at = ?,
                returned_at = ?,
                updated_at = ?
            WHERE job_id = ?
            """,
            ("success", returned_name, returned_file_count, timestamp, timestamp, timestamp, job_id),
        )

    def record_file_result(
        self,
        *,
        job_id: str,
        source_name: str,
        output_name: str | None,
        status: str,
        error_message: str | None = None,
    ) -> None:
        self._execute(
            """
            INSERT INTO conversion_files (
                job_id,
                source_name,
                output_name,
                status,
                error_message,
                created_at
            ) VALUES (?, ?, ?, ?, ?, ?)
            """,
            (job_id, source_name, output_name, status, error_message, _utc_now()),
        )

    def mark_error(
        self,
        job_id: str,
        *,
        error_message: str,
        debug_path: str | None,
        metadata: dict[str, Any] | None = None,
    ) -> None:
        self._execute(
            """
            UPDATE conversion_jobs
            SET status = ?,
                error_message = ?,
                debug_path = ?,
                metadata_json = ?,
                completed_at = ?,
                updated_at = ?
            WHERE job_id = ?
            """,
            (
                "error",
                error_message,
                debug_path,
                json.dumps(metadata or {}, ensure_ascii=True),
                _utc_now(),
                _utc_now(),
                job_id,
            ),
        )

    def list_recent_jobs(self, limit: int = 100) -> list[dict[str, Any]]:
        return self._query_all(
            """
            SELECT
                job_id,
                request_type,
                source_name,
                source_size_bytes,
                source_file_count,
                remote_addr,
                status,
                returned_name,
                returned_file_count,
                error_message,
                debug_path,
                metadata_json,
                stage,
                progress_message,
                progress_current,
                progress_total,
                created_at,
                started_at,
                completed_at,
                returned_at,
                updated_at
            FROM conversion_jobs
            ORDER BY created_at DESC
            LIMIT ?
            """,
            (limit,),
        )

    def update_job_artifact_metadata(
        self,
        job_id: str,
        *,
        debug_path: str | None,
        metadata: dict[str, Any],
    ) -> None:
        self._execute(
            """
            UPDATE conversion_jobs
            SET debug_path = ?,
                metadata_json = ?,
                updated_at = ?
            WHERE job_id = ?
            """,
            (debug_path, json.dumps(metadata, ensure_ascii=True), _utc_now(), job_id),
        )

    def get_job(self, job_id: str) -> dict[str, Any] | None:
        return self._query_one(
            """
            SELECT
                job_id,
                request_type,
                source_name,
                source_size_bytes,
                source_file_count,
                remote_addr,
                status,
                returned_name,
                returned_file_count,
                error_message,
                debug_path,
                metadata_json,
                stage,
                progress_message,
                progress_current,
                progress_total,
                created_at,
                started_at,
                completed_at,
                returned_at,
                updated_at
            FROM conversion_jobs
            WHERE job_id = ?
            """,
            (job_id,),
        )

    def list_error_jobs(self) -> list[dict[str, Any]]:
        """Return failed jobs oldest-first so retention cleanup can delete safely by age."""
        return self._query_all(
            """
            SELECT
                job_id,
                request_type,
                source_name,
                source_size_bytes,
                source_file_count,
                remote_addr,
                status,
                returned_name,
                returned_file_count,
                error_message,
                debug_path,
                metadata_json,
                stage,
                progress_message,
                progress_current,
                progress_total,
                created_at,
                started_at,
                completed_at,
                returned_at,
                updated_at
            FROM conversion_jobs
            WHERE status = 'error'
            ORDER BY COALESCE(completed_at, created_at) ASC
            """
        )

    def list_recent_files(self, limit: int = 200) -> list[dict[str, Any]]:
        """Return the per-file conversion log used by the admin dashboard."""
        return self._query_all(
            """
            SELECT
                cf.job_id,
                cf.source_name,
                cf.output_name,
                cf.status,
                cf.error_message,
                cf.created_at
            FROM conversion_files AS cf
            ORDER BY cf.created_at DESC
            LIMIT ?
            """,
            (limit,),
        )

    def get_summary(self) -> dict[str, Any]:
        """Return compact counts for the admin dashboard cards."""
        jobs = self._query_one(
            """
            SELECT
                COUNT(*) AS total_jobs,
                SUM(CASE WHEN status = 'success' THEN 1 ELSE 0 END) AS success_jobs,
                SUM(CASE WHEN status = 'error' THEN 1 ELSE 0 END) AS error_jobs
            FROM conversion_jobs
            """
        ) or {}
        files = self._query_one(
            """
            SELECT
                COUNT(*) AS total_files,
                SUM(CASE WHEN status = 'success' THEN 1 ELSE 0 END) AS success_files,
                SUM(CASE WHEN status = 'error' THEN 1 ELSE 0 END) AS error_files
            FROM conversion_files
            """
        ) or {}
        return {**jobs, **files}

    def _update_status(self, job_id: str, status: str, started_at: str | None = None) -> None:
        self._execute(
            """
            UPDATE conversion_jobs
            SET status = ?,
                started_at = COALESCE(?, started_at),
                updated_at = ?
            WHERE job_id = ?
            """,
            (status, started_at, _utc_now(), job_id),
        )

    def _init_db(self) -> None:
        """Create the audit tables and add newer columns when upgrading older databases."""
        self._execute(
            """
            CREATE TABLE IF NOT EXISTS conversion_jobs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id TEXT NOT NULL UNIQUE,
                request_type TEXT NOT NULL,
                source_name TEXT NOT NULL,
                source_size_bytes INTEGER NOT NULL,
                source_file_count INTEGER NOT NULL,
                remote_addr TEXT,
                status TEXT NOT NULL,
                returned_name TEXT,
                returned_file_count INTEGER,
                error_message TEXT,
                debug_path TEXT,
                metadata_json TEXT,
                stage TEXT,
                progress_message TEXT,
                progress_current INTEGER DEFAULT 0,
                progress_total INTEGER DEFAULT 1,
                created_at TEXT NOT NULL,
                started_at TEXT,
                completed_at TEXT,
                returned_at TEXT,
                updated_at TEXT NOT NULL
            )
            """
        )
        for column_name, column_type, default_sql in [
            ("stage", "TEXT", ""),
            ("progress_message", "TEXT", ""),
            ("progress_current", "INTEGER", " DEFAULT 0"),
            ("progress_total", "INTEGER", " DEFAULT 1"),
        ]:
            self._ensure_column("conversion_jobs", column_name, column_type, default_sql)
        self._execute(
            """
            CREATE TABLE IF NOT EXISTS conversion_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id TEXT NOT NULL,
                source_name TEXT NOT NULL,
                output_name TEXT,
                status TEXT NOT NULL,
                error_message TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(job_id) REFERENCES conversion_jobs(job_id)
            )
            """
        )

    def _execute(self, sql: str, params: tuple[Any, ...] = ()) -> None:
        """Run one write query with a short-lived connection."""
        with closing(sqlite3.connect(self.db_path)) as connection:
            connection.execute(sql, params)
            connection.commit()

    def _query_all(self, sql: str, params: tuple[Any, ...] = ()) -> list[dict[str, Any]]:
        """Return all rows as plain dictionaries to keep callers framework-agnostic."""
        with closing(sqlite3.connect(self.db_path)) as connection:
            connection.row_factory = sqlite3.Row
            cursor = connection.execute(sql, params)
            return [dict(row) for row in cursor.fetchall()]

    def _query_one(self, sql: str, params: tuple[Any, ...] = ()) -> dict[str, Any] | None:
        """Return one row as a dictionary, or None when no record matches."""
        with closing(sqlite3.connect(self.db_path)) as connection:
            connection.row_factory = sqlite3.Row
            cursor = connection.execute(sql, params)
            row = cursor.fetchone()
            return dict(row) if row else None

    def _ensure_column(self, table_name: str, column_name: str, column_type: str, default_sql: str = "") -> None:
        """Perform a lightweight schema migration without requiring a separate tool."""
        with closing(sqlite3.connect(self.db_path)) as connection:
            cursor = connection.execute(f"PRAGMA table_info({table_name})")
            existing_columns = {row[1] for row in cursor.fetchall()}
            if column_name not in existing_columns:
                connection.execute(
                    f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}{default_sql}"
                )
                connection.commit()


def _utc_now() -> str:
    """Store timestamps in UTC ISO format so sorting and retention logic stay consistent."""
    return datetime.now(timezone.utc).isoformat(timespec="seconds")
