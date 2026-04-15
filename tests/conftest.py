from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import pytest

import app as app_module


@pytest.fixture
def app_instance(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    data_dir = tmp_path / "data"
    failure_root = tmp_path / "failed_uploads"
    upload_root = tmp_path / "upload_tmp"

    monkeypatch.setenv("SECRET_KEY", "test-secret")
    monkeypatch.setenv("ADMIN_USERNAME", "admin")
    monkeypatch.setenv("ADMIN_PASSWORD", "s3cret")
    monkeypatch.setenv("AUDIT_DB_PATH", str(data_dir / "audit.sqlite3"))
    monkeypatch.setenv("FAILURE_ROOT", str(failure_root))
    monkeypatch.setenv("UPLOAD_ROOT", str(upload_root))

    flask_app = app_module.create_app()
    flask_app.config.update(TESTING=True)
    return flask_app


@pytest.fixture
def client(app_instance):
    return app_instance.test_client()


@pytest.fixture
def export_case() -> tuple[pd.DataFrame, list[str], list[str]]:
    fixture_path = Path(__file__).parent / "fixtures" / "export_case.json"
    payload = json.loads(fixture_path.read_text(encoding="utf-8"))
    dataframe = pd.DataFrame(payload["rows"])
    return dataframe, payload["mp_codes_with_em"], payload["mp_codes"]
