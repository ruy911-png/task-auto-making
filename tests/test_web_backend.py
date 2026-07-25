import io

import pandas as pd
import pytest
from fastapi.testclient import TestClient

from web.backend import app


@pytest.fixture(autouse=True)
def isolate_dirs(tmp_path, monkeypatch):
    monkeypatch.setattr("web.backend.CONFIG_DIR", tmp_path / "config" / "transactions")
    monkeypatch.setattr("web.backend.JOBS_DIR", tmp_path / "data" / "jobs")


client = TestClient(app)


def _sample_control_payload():
    return {
        "control_id": "test_control",
        "description": "테스트 통제",
        "transaction": "FB03",
        "fields": {"회사코드": "wnd[0]/usr/ctxtBUKRS"},
        "download": {"menu_path": []},
        "edit_rules": {},
        "validation": {"key_columns": [], "responsible_column": "담당자"},
    }


def _excel_bytes(rows):
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return buf


def test_register_and_list_control():
    res = client.post("/api/controls", json=_sample_control_payload())
    assert res.status_code == 200
    assert res.json()["control_id"] == "test_control"

    res = client.get("/api/controls")
    assert any(c["control_id"] == "test_control" for c in res.json())


def test_register_control_missing_field_returns_400():
    payload = _sample_control_payload()
    del payload["transaction"]
    res = client.post("/api/controls", json=payload)
    assert res.status_code == 400


def test_sap_sessions_unavailable_message():
    res = client.get("/api/sap/sessions")
    data = res.json()
    assert data["available"] is False
    assert data["sessions"] == []


def test_full_run_with_mock_session():
    client.post("/api/controls", json=_sample_control_payload())

    condition_file = _excel_bytes([{"회사코드": "1000"}, {"회사코드": "2000"}])
    hr_file = _excel_bytes(
        [{"사번": "EMP001", "이름": "홍길동", "아이디": "hong1", "부서": "회계팀"}]
    )

    res = client.post(
        "/api/runs",
        data={"control_id": "test_control", "session_id": "__mock__"},
        files={
            "condition_file": ("conditions.xlsx", condition_file, "application/octet-stream"),
            "hr_file": ("hr.xlsx", hr_file, "application/octet-stream"),
        },
    )
    assert res.status_code == 200
    data = res.json()
    job_id = data["job_id"]
    assert data["summary"] == {"total": 2, "success": 2, "failed": 0}

    download_res = client.get(f"/api/runs/{job_id}/download")
    assert download_res.status_code == 200

    retry_res = client.post(f"/api/runs/{job_id}/retry")
    assert retry_res.status_code == 200
    assert retry_res.json()["summary"]["failed"] == 0
