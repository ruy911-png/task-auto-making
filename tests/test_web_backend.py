import io
from pathlib import Path

import openpyxl
import pandas as pd
import pytest
from fastapi.testclient import TestClient

from batch.runner import BatchState, ConditionResult, ConditionStatus
from sap_automation.control_config import ControlConfig, SapQuerySchema
from web.backend import _JOB_META, _finalize_result, app


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


def test_finalize_result_excludes_failed_condition_from_population_and_sample(tmp_path):
    """조회 실패 조건의 오류 플레이스홀더 행이 담당자매칭 컬럼이 없는 통제에서도
    모집단/샘플로 새어 들어가지 않고 예외로만 분류되는지 확인 (QA에서 발견된 회귀)."""
    config = ControlConfig(
        control_id="c1",
        description="",
        sap=SapQuerySchema(transaction="FB03", fields={}),
        validation={},  # responsible_column 미설정 -> match_rows가 호출되지 않는 경로
    )
    state = BatchState(
        job_id="job1",
        items=[
            ConditionResult(
                index=0,
                condition={},
                status=ConditionStatus.SUCCESS,
                result={
                    "raw_rows": [{"a": 1}],
                    "rows": [{"a": 1}],
                    "condition_capture": None,
                    "result_capture": None,
                },
            ),
            ConditionResult(
                index=1,
                condition={},
                status=ConditionStatus.SUCCESS,
                result={
                    "raw_rows": [{"a": 2}],
                    "rows": [{"a": 2}],
                    "condition_capture": None,
                    "result_capture": None,
                },
            ),
            ConditionResult(index=2, condition={}, status=ConditionStatus.FAILED, error="SAP timeout"),
        ],
    )
    job_dir = tmp_path / "job1"
    job_dir.mkdir()

    _finalize_result(config, state, job_dir)

    workbook = openpyxl.load_workbook(job_dir / "result.xlsx")

    def _sheet_text(sheet_name: str) -> str:
        sheet = workbook[sheet_name]
        return " ".join(
            str(cell.value) for row in sheet.iter_rows() for cell in row if cell.value is not None
        )

    assert "조회실패" not in _sheet_text("모집단 data")
    assert "조회실패" not in _sheet_text("샘플 data")
    assert "조회실패" in _sheet_text("예외 data")


def test_full_run_with_reference_upload_matching():
    """담당자 매칭 방식을 "reference_upload"로 등록하면, hr_file 대신 reference_file로
    전표번호 기준 매칭이 되어야 한다 (SAP/인사데이터 엑셀 전혀 접촉하지 않음)."""
    payload = _sample_control_payload()
    payload["matching"] = {
        "method": "reference_upload",
        "match_column": "전표번호",
        "reference_key_column": "전표번호",
        "reference_name_column": "담당자",
        "reference_department_column": "담당부서",
    }
    client.post("/api/controls", json=payload)

    condition_file = _excel_bytes(
        [{"회사코드": "1000", "전표번호": "1900000123"}, {"회사코드": "1000", "전표번호": "9999999999"}]
    )
    reference_file = _excel_bytes(
        [{"전표번호": "1900000123", "담당자": "박민수", "담당부서": "회계팀"}]
    )

    res = client.post(
        "/api/runs",
        data={"control_id": "test_control", "session_id": "__mock__"},
        files={
            "condition_file": ("conditions.xlsx", condition_file, "application/octet-stream"),
            "reference_file": ("reference.xlsx", reference_file, "application/octet-stream"),
        },
    )
    assert res.status_code == 200
    job_id = res.json()["job_id"]

    workbook = openpyxl.load_workbook(_result_path_for(job_id))
    population_text = " ".join(
        str(cell.value)
        for row in workbook["모집단 data"].iter_rows()
        for cell in row
        if cell.value is not None
    )
    assert "박민수" in population_text


def _result_path_for(job_id: str) -> Path:
    return Path(_JOB_META[job_id]["job_dir"]) / "result.xlsx"


def test_run_without_reference_file_for_reference_upload_control_returns_400():
    payload = _sample_control_payload()
    payload["matching"] = {"method": "reference_upload", "match_column": "전표번호"}
    client.post("/api/controls", json=payload)

    condition_file = _excel_bytes([{"회사코드": "1000", "전표번호": "1900000123"}])

    res = client.post(
        "/api/runs",
        data={"control_id": "test_control", "session_id": "__mock__"},
        files={"condition_file": ("conditions.xlsx", condition_file, "application/octet-stream")},
    )
    assert res.status_code == 400


def test_run_with_sap_lookup_method_returns_400():
    payload = _sample_control_payload()
    payload["matching"] = {"method": "sap_lookup"}
    client.post("/api/controls", json=payload)

    condition_file = _excel_bytes([{"회사코드": "1000"}])
    hr_file = _excel_bytes([{"사번": "EMP001", "이름": "홍길동", "아이디": "hong1", "부서": "회계팀"}])

    res = client.post(
        "/api/runs",
        data={"control_id": "test_control", "session_id": "__mock__"},
        files={
            "condition_file": ("conditions.xlsx", condition_file, "application/octet-stream"),
            "hr_file": ("hr.xlsx", hr_file, "application/octet-stream"),
        },
    )
    assert res.status_code == 400
