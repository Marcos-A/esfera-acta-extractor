from __future__ import annotations

import hmac
import json
import os
import shutil
import tempfile
import threading
import traceback
import uuid
from datetime import datetime, timezone
from functools import wraps
from io import BytesIO
from pathlib import Path

from flask import (
    Flask,
    abort,
    after_this_request,
    flash,
    redirect,
    render_template,
    request,
    jsonify,
    send_file,
    session,
    url_for,
)
from rich.console import Console
from rich.traceback import Traceback
from werkzeug.utils import secure_filename

from src.audit import AuditStore
from src.conversion_service import (
    build_zip_from_artifacts,
    cleanup_path,
    ConversionArtifact,
    convert_pdf_to_excel,
    extract_zip_to_temp,
)
from src.notifier import notify_failure


def create_app() -> Flask:
    """Create the Flask application used by the public upload flow and admin UI."""
    app = Flask(__name__)
    default_data_dir = str(Path(app.root_path) / "data")
    default_failure_dir = str(Path(app.root_path) / "failed_uploads")
    app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "change-me-before-production")
    app.config["MAX_CONTENT_LENGTH"] = int(os.getenv("MAX_UPLOAD_SIZE_MB", "50")) * 1024 * 1024
    app.config["UPLOAD_ROOT"] = os.getenv("UPLOAD_ROOT", "/tmp/esfera-acta-extractor")
    app.config["FAILURE_ROOT"] = os.getenv("FAILURE_ROOT", default_failure_dir)
    app.config["AUDIT_DB_PATH"] = os.getenv("AUDIT_DB_PATH", str(Path(default_data_dir) / "conversion_audit.sqlite3"))
    app.config["ADMIN_USERNAME"] = os.getenv("ADMIN_USERNAME", "admin")
    app.config["ADMIN_PASSWORD"] = os.getenv("ADMIN_PASSWORD", "change-me")

    Path(app.config["UPLOAD_ROOT"]).mkdir(parents=True, exist_ok=True)
    Path(app.config["FAILURE_ROOT"]).mkdir(parents=True, exist_ok=True)
    app.audit_store = AuditStore(app.config["AUDIT_DB_PATH"])
    app.config["PUBLIC_ERROR_MESSAGE"] = (
        "Ep! No hem pogut convertir aquest fitxer. Hi ha hagut un problema durant el procés. "
        "Pot ser un error temporal nostre o bé que el fitxer pujat no sigui una acta d'Esfer@. "
        "Si és el cas, prova-ho amb un altre fitxer."
    )

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.get("/health")
    def health():
        return {"status": "ok"}

    @app.route("/admin/login", methods=["GET", "POST"])
    def admin_login():
        error = None
        if request.method == "POST":
            username = request.form.get("username", "")
            password = request.form.get("password", "")
            if _valid_admin_credentials(app, username, password):
                session["admin_authenticated"] = True
                session["admin_username"] = username
                return redirect(url_for("admin_dashboard"))
            error = "Invalid credentials."
        return render_template("admin_login.html", error=error)

    @app.post("/admin/logout")
    def admin_logout():
        session.clear()
        return redirect(url_for("admin_login"))

    @app.get("/admin")
    @_admin_required
    def admin_dashboard():
        summary = app.audit_store.get_summary()
        recent_jobs = [_present_admin_job(job) for job in app.audit_store.list_recent_jobs(limit=100)]
        recent_files = app.audit_store.list_recent_files(limit=200)
        return render_template(
            "admin_dashboard.html",
            summary=summary,
            recent_jobs=recent_jobs,
            recent_files=recent_files,
            failure_root=app.config["FAILURE_ROOT"],
            admin_username=session.get("admin_username", app.config["ADMIN_USERNAME"]),
        )

    @app.post("/admin/job/<job_id>/delete-artifacts")
    @_admin_required
    def admin_delete_job_artifacts(job_id: str):
        job = app.audit_store.get_job(job_id)
        if job is None:
            abort(404, "Job not found.")
        _delete_job_artifacts(app, job, reason="manual admin cleanup")
        flash("Retained debug artifacts deleted for this job.")
        return redirect(url_for("admin_dashboard"))

    @app.get("/admin/job/<job_id>/download/source")
    @_admin_required
    def admin_download_failed_source(job_id: str):
        job = app.audit_store.get_job(job_id)
        if job is None:
            abort(404, "Job not found.")
        metadata = _load_metadata(job.get("metadata_json"))
        source_path_value = metadata.get("failed_source_path")
        if not source_path_value:
            abort(404, "No retained source file for this job.")
        source_path = Path(source_path_value)
        if not source_path.exists():
            abort(404, "Retained source file is no longer available.")
        return send_file(
            source_path,
            as_attachment=True,
            download_name=source_path.name,
            mimetype=_download_mimetype(source_path),
        )

    @app.get("/admin/job/<job_id>/download/log")
    @_admin_required
    def admin_download_failure_log(job_id: str):
        job = app.audit_store.get_job(job_id)
        if job is None:
            abort(404, "Job not found.")
        metadata = _load_metadata(job.get("metadata_json"))
        log_path_value = metadata.get("failure_log_path")
        if not log_path_value:
            abort(404, "No failure log for this job.")
        log_path = Path(log_path_value)
        if not log_path.exists():
            abort(404, "Failure log is no longer available.")
        return send_file(
            log_path,
            as_attachment=True,
            download_name=log_path.name,
            mimetype="text/plain; charset=utf-8",
        )

    @app.post("/convert")
    def convert():
        uploaded_file = request.files.get("file")
        if uploaded_file is None or uploaded_file.filename is None or not uploaded_file.filename.strip():
            abort(400, "Cal carregar un fitxer PDF o ZIP.")

        job_id = uuid.uuid4().hex
        source_name = secure_filename(uploaded_file.filename) or f"upload-{job_id}"
        source_ext = Path(source_name).suffix.lower()
        if source_ext not in {".pdf", ".zip"}:
            abort(400, "Només s'admeten fitxers PDF i ZIP.")

        request_type = "zip" if source_ext == ".zip" else "pdf"
        work_dir = Path(tempfile.mkdtemp(prefix=f"esfera-job-{job_id}-", dir=app.config["UPLOAD_ROOT"]))
        extracted_dir: Path | None = None
        upload_path = work_dir / source_name
        uploaded_file.save(upload_path)
        source_size_bytes = upload_path.stat().st_size
        source_file_count = 1

        app.audit_store.create_job(
            job_id=job_id,
            request_type=request_type,
            source_name=source_name,
            source_size_bytes=source_size_bytes,
            source_file_count=source_file_count,
            remote_addr=request.headers.get("X-Forwarded-For", request.remote_addr),
        )
        app.audit_store.update_progress(
            job_id,
            stage="received",
            message="Fitxer rebut. Preparant la conversió.",
            progress_current=0,
            progress_total=1,
            metadata={},
        )
        threading.Thread(
            target=_run_conversion_job,
            args=(app, job_id, request_type, source_name, upload_path, work_dir),
            daemon=True,
        ).start()
        return jsonify(
            {
                "job_id": job_id,
                "status_url": url_for("conversion_status", job_id=job_id),
                "download_url": url_for("conversion_download", job_id=job_id),
            }
        ), 202

    @app.get("/convert/status/<job_id>")
    def conversion_status(job_id: str):
        job = app.audit_store.get_job(job_id)
        if job is None:
            abort(404, "No s'ha trobat la conversió.")
        metadata = _load_metadata(job.get("metadata_json"))
        progress_total = max(job.get("progress_total") or 1, 1)
        progress_current = min(job.get("progress_current") or 0, progress_total)
        return jsonify(
            {
                "job_id": job["job_id"],
                "status": job["status"],
                "stage": job.get("stage") or "",
                "message": job.get("progress_message") or "",
                "progress_current": progress_current,
                "progress_total": progress_total,
                "percent": int((progress_current / progress_total) * 100),
                "returned_name": job.get("returned_name"),
                "error_message": job.get("error_message"),
                "public_error_message": app.config["PUBLIC_ERROR_MESSAGE"],
                "download_ready": job["status"] == "success",
                "download_url": url_for("conversion_download", job_id=job_id),
                "source_file_count": job.get("source_file_count"),
                "metadata": metadata,
            }
        )

    @app.get("/convert/download/<job_id>")
    def conversion_download(job_id: str):
        job = app.audit_store.get_job(job_id)
        if job is None:
            abort(404, "No s'ha trobat la conversió.")
        if job["status"] != "success":
            abort(409, "La conversió encara no està disponible.")

        metadata = _load_metadata(job.get("metadata_json"))
        output_path_value = metadata.get("output_path")
        if not output_path_value:
            abort(404, "No s'ha trobat el fitxer convertit.")

        output_path = Path(output_path_value)
        if not output_path.exists():
            abort(404, "El fitxer convertit ja no és disponible.")

        work_dir_value = metadata.get("work_dir")

        @after_this_request
        def _cleanup(response):
            cleanup_path(output_path)
            if work_dir_value:
                cleanup_path(work_dir_value)
            return response

        return send_file(
            BytesIO(output_path.read_bytes()),
            as_attachment=True,
            download_name=job.get("returned_name") or output_path.name,
            mimetype=_download_mimetype(output_path),
        )

    return app


def _admin_required(view_func):
    """Redirect unauthenticated admin requests to the login page."""
    @wraps(view_func)
    def wrapped(*args, **kwargs):
        if not session.get("admin_authenticated"):
            return redirect(url_for("admin_login"))
        return view_func(*args, **kwargs)

    return wrapped


def _valid_admin_credentials(app: Flask, username: str, password: str) -> bool:
    """Compare credentials in constant time to avoid leaking partial matches."""
    return hmac.compare_digest(username, app.config["ADMIN_USERNAME"]) and hmac.compare_digest(
        password,
        app.config["ADMIN_PASSWORD"],
    )


def _retain_failed_job(
    failure_root: str | os.PathLike[str],
    job_id: str,
    source_name: str,
    upload_path: Path,
    work_dir: Path,
) -> Path:
    """Keep failed inputs and partial outputs so administrators can inspect conversion errors later."""
    target_dir = Path(failure_root) / job_id
    target_dir.mkdir(parents=True, exist_ok=True)

    if upload_path.exists():
        shutil.copy2(upload_path, target_dir / source_name)

    output_dir = work_dir / "output"
    if output_dir.exists():
        target_output = target_dir / "output"
        shutil.copytree(output_dir, target_output, dirs_exist_ok=True)

    return target_dir


def _run_conversion_job(
    app: Flask,
    job_id: str,
    request_type: str,
    source_name: str,
    upload_path: Path,
    work_dir: Path,
) -> None:
    """
    Run the conversion in a background thread and keep the audit trail updated.

    The web request returns immediately with a job id, so this worker is responsible for
    recording progress, preserving failed artifacts, and cleaning temporary files when
    processing succeeds.
    """
    extracted_dir: Path | None = None
    source_file_count = 1
    app.audit_store.mark_started(job_id)
    try:
        output_dir = work_dir / "output"
        output_dir.mkdir(parents=True, exist_ok=True)

        if request_type == "pdf":
            app.audit_store.update_progress(
                job_id,
                stage="processing",
                message="Convertint el fitxer 1 de 1",
                progress_current=0,
                progress_total=1,
                metadata={"work_dir": str(work_dir)},
            )
            output_path = convert_pdf_to_excel(upload_path, output_dir)
            app.audit_store.record_file_result(
                job_id=job_id,
                source_name=source_name,
                output_name=output_path.name,
                status="success",
            )
            cleanup_path(upload_path)
            app.audit_store.update_progress(
                job_id,
                stage="ready",
                message="Conversió completada. Preparant la descàrrega.",
                progress_current=1,
                progress_total=1,
                metadata={"work_dir": str(work_dir), "output_path": str(output_path)},
            )
            app.audit_store.mark_success(
                job_id,
                returned_name=output_path.name,
                returned_file_count=1,
            )
            return

        app.audit_store.update_progress(
            job_id,
            stage="extracting",
            message="Descomprimint el fitxer",
            progress_current=0,
            progress_total=1,
            metadata={"work_dir": str(work_dir)},
        )
        extracted_dir, pdf_paths = extract_zip_to_temp(upload_path)
        source_file_count = len(pdf_paths)
        app.audit_store.update_source_file_count(job_id, source_file_count)
        cleanup_path(upload_path)
        artifacts: list[ConversionArtifact] = []

        for index, pdf_path in enumerate(pdf_paths, start=1):
            try:
                # ZIP uploads are processed one PDF at a time so the status endpoint can
                # show meaningful progress instead of a single long-running "busy" state.
                app.audit_store.update_progress(
                    job_id,
                    stage="processing",
                    message=f"Convertint el fitxer {index} de {source_file_count}",
                    progress_current=index - 1,
                    progress_total=source_file_count + 1,
                    metadata={"work_dir": str(work_dir), "current_file": pdf_path.name},
                )
                output_path = convert_pdf_to_excel(pdf_path, output_dir)
                artifact = ConversionArtifact(
                    source_name=pdf_path.name,
                    output_name=output_path.name,
                    output_path=output_path,
                )
                artifacts.append(artifact)
                app.audit_store.record_file_result(
                    job_id=job_id,
                    source_name=artifact.source_name,
                    output_name=artifact.output_name,
                    status="success",
                )
            except Exception as file_exc:
                app.audit_store.record_file_result(
                    job_id=job_id,
                    source_name=pdf_path.name,
                    output_name=None,
                    status="error",
                    error_message=str(file_exc),
                )
                raise RuntimeError(f"Failed to convert {pdf_path.name}: {file_exc}") from file_exc

        app.audit_store.update_progress(
            job_id,
            stage="compressing",
            message="Comprimint els fitxers convertits",
            progress_current=source_file_count,
            progress_total=source_file_count + 1,
            metadata={"work_dir": str(work_dir)},
        )
        zip_path = build_zip_from_artifacts(artifacts, work_dir / f"{Path(source_name).stem}-converted.zip")
        if extracted_dir is not None:
            cleanup_path(extracted_dir)
        app.audit_store.update_progress(
            job_id,
            stage="ready",
            message="Conversió completada. Preparant la descàrrega.",
            progress_current=source_file_count + 1,
            progress_total=source_file_count + 1,
            metadata={"work_dir": str(work_dir), "output_path": str(zip_path)},
        )
        app.audit_store.mark_success(
            job_id,
            returned_name=zip_path.name,
            returned_file_count=len(artifacts),
        )
    except Exception as exc:
        debug_path = _retain_failed_job(app.config["FAILURE_ROOT"], job_id, source_name, upload_path, work_dir)
        current_job = app.audit_store.get_job(job_id)
        existing_metadata = _load_metadata((current_job or {}).get("metadata_json"))
        failure_log_path = _write_failure_log(
            debug_path=debug_path,
            job_id=job_id,
            request_type=request_type,
            source_name=source_name,
            exception=exc,
            traceback_text=traceback.format_exc(),
            job=current_job,
        )
        failed_source_path = _find_failed_source_path(debug_path)
        app.audit_store.record_file_result(
            job_id=job_id,
            source_name=source_name,
            output_name=None,
            status="error",
            error_message=str(exc),
        )
        app.audit_store.update_progress(
            job_id,
            stage="error",
            message="La conversió ha fallat.",
            progress_current=0,
            progress_total=1,
            metadata={"work_dir": str(work_dir)},
        )
        app.audit_store.mark_error(
            job_id,
            error_message=str(exc),
            debug_path=str(debug_path),
            metadata={
                **existing_metadata,
                "source_file_count": source_file_count,
                "work_dir": str(work_dir),
                "failed_source_path": str(failed_source_path) if failed_source_path else None,
                "failure_log_path": str(failure_log_path),
            },
        )
        notify_failure(
            subject=f"[esfera-acta-extractor] Conversion failed for {source_name}",
            body=(
                f"Job ID: {job_id}\n"
                f"Source: {source_name}\n"
                f"Type: {request_type}\n"
                f"Error: {exc}\n"
                f"Debug path: {debug_path}\n"
            ),
        )
        if extracted_dir is not None:
            cleanup_path(extracted_dir)


def _present_admin_job(job: dict[str, object]) -> dict[str, object]:
    """Add derived admin-only flags so the template can stay simple."""
    presented = dict(job)
    metadata = _load_metadata(job.get("metadata_json"))
    failed_source_path = metadata.get("failed_source_path")
    failure_log_path = metadata.get("failure_log_path")
    presented["metadata"] = metadata
    presented["has_failed_source"] = bool(failed_source_path and Path(str(failed_source_path)).exists())
    presented["has_failure_log"] = bool(failure_log_path and Path(str(failure_log_path)).exists())
    presented["artifacts_deleted_at"] = metadata.get("artifacts_deleted_at")
    return presented


def _delete_job_artifacts(app: Flask, job: dict[str, object], *, reason: str) -> None:
    """Delete retained debug files while preserving an audit note about why they were removed."""
    metadata = _load_metadata(job.get("metadata_json"))
    debug_path_value = job.get("debug_path") or metadata.get("debug_path")
    if debug_path_value:
        debug_path = Path(str(debug_path_value))
        if debug_path.exists():
            cleanup_path(debug_path)

    retained_debug_path = None
    if debug_path_value:
        debug_path = Path(str(debug_path_value))
        if debug_path.exists():
            retained_debug_path = str(debug_path)

    metadata["failed_source_path"] = None
    metadata["failure_log_path"] = None
    metadata["artifacts_deleted_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
    metadata["artifacts_deleted_reason"] = reason
    app.audit_store.update_job_artifact_metadata(
        str(job["job_id"]),
        debug_path=retained_debug_path,
        metadata=metadata,
    )


def _load_metadata(raw_metadata: str | None) -> dict[str, object]:
    """Safely decode stored JSON metadata from the audit database."""
    if not raw_metadata:
        return {}
    try:
        parsed = json.loads(raw_metadata)
        return parsed if isinstance(parsed, dict) else {}
    except json.JSONDecodeError:
        return {}


def _download_mimetype(output_path: Path) -> str:
    """Return a download content type that matches the generated artifact."""
    if output_path.suffix.lower() == ".zip":
        return "application/zip"
    if output_path.suffix.lower() == ".pdf":
        return "application/pdf"
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _find_failed_source_path(debug_path: Path) -> Path | None:
    """Return the original uploaded file stored for a failed job, if present."""
    for path in debug_path.iterdir():
        if path.is_file() and path.name != "failure.log":
            return path
    return None


def _write_failure_log(
    *,
    debug_path: Path,
    job_id: str,
    request_type: str,
    source_name: str,
    exception: Exception,
    traceback_text: str,
    job: dict[str, object] | None,
) -> Path:
    """Write a human-readable failure report alongside retained debug artifacts."""
    log_path = debug_path / "failure.log"
    console = Console(record=True, width=120)
    console.print(
        Traceback.from_exception(
            type(exception),
            exception,
            exception.__traceback__,
            show_locals=True,
            max_frames=50,
        )
    )
    detailed_traceback = console.export_text()
    lines = [
        f"job_id: {job_id}",
        f"request_type: {request_type}",
        f"source_name: {source_name}",
        f"created_at_utc: {(job or {}).get('created_at', '')}",
        f"started_at_utc: {(job or {}).get('started_at', '')}",
        f"failed_at_utc: {datetime.now(timezone.utc).isoformat(timespec='seconds')}",
        f"status: {(job or {}).get('status', '')}",
        f"stage: {(job or {}).get('stage', '')}",
        f"progress_message: {(job or {}).get('progress_message', '')}",
        f"progress_current: {(job or {}).get('progress_current', '')}",
        f"progress_total: {(job or {}).get('progress_total', '')}",
        f"error_type: {type(exception).__name__}",
        f"error_message: {exception}",
        "",
        "standard_traceback:",
        traceback_text.rstrip(),
        "",
        "rich_traceback:",
        detailed_traceback.rstrip(),
    ]
    metadata = _load_metadata((job or {}).get("metadata_json")) if job else {}
    if metadata:
        lines.extend(
            [
                "",
                "metadata:",
                json.dumps(metadata, ensure_ascii=True, indent=2, sort_keys=True),
            ]
        )
    log_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return log_path


app = create_app()
