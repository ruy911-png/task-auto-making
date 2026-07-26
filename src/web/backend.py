"""로컬 웹 대시보드 백엔드 (FastAPI).

사용자 PC에서 로컬로 실행된다 (SAP GUI Scripting이 로그인된 Windows 세션에서만
동작하기 때문에, 이 서버 자체를 원격/클라우드에 둘 수 없다). 두 화면을 지원한다:

1. 통제 등록 — 조회(필드매핑)+다운로드(절차)+편집(컬럼정리/필터/계산컬럼)+검증 스키마를
   입력받아 config/transactions/<control_id>.yaml로 저장한다.
2. 실행 — 통제 선택 + SAP 세션 선택 + 조건/인사데이터 엑셀 업로드 + 실행, 완료 후
   성공/실패 상태와 결과 다운로드, 실패건만 재처리.
"""
from __future__ import annotations

import shutil
import uuid
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from batch.runner import BatchState, retry_failed, run_batch
from excel_io import reader, writer
from excel_io.edit_rules import apply_edit_rules
from lookup.hr_matcher import match_rows
from sap_automation import control_config, session_picker, transaction_runner
from sap_automation.control_config import (
    ControlConfig,
    DownloadSchema,
    LayoutSchema,
    SamplingSchema,
    SapQuerySchema,
)
from validation.checks import EXCEPTION_REASON_COLUMN, classify_rows
from validation.sampling import determine_sample_size, select_sample

BASE_DIR = Path(__file__).resolve().parent.parent.parent
CONFIG_DIR = BASE_DIR / "config" / "transactions"
DATA_DIR = BASE_DIR / "data"
JOBS_DIR = DATA_DIR / "jobs"

MOCK_SESSION_ID = "__mock__"
"""SAP GUI가 없는 환경에서 파이프라인 전체(편집규칙/담당자매칭/검증/결과엑셀)를 확인하기
위한 테스트용 세션 id. 실제 배포 환경에서는 실제 세션 목록만 노출하면 된다."""

app = FastAPI(title="ERP 모집단 자동 추출 도구")

_STATIC_DIR = Path(__file__).parent / "static"
app.mount("/static", StaticFiles(directory=_STATIC_DIR), name="static")

_JOB_META: dict[str, dict[str, Any]] = {}


@app.get("/")
def index() -> FileResponse:
    return FileResponse(_STATIC_DIR / "index.html")


# ---------- 통제 등록 ----------


@app.get("/api/controls")
def list_controls() -> list[dict[str, Any]]:
    return [c.to_dict() for c in control_config.load_all(CONFIG_DIR)]


@app.post("/api/controls")
def create_control(payload: dict[str, Any]) -> dict[str, Any]:
    try:
        config = ControlConfig(
            control_id=payload["control_id"],
            description=payload.get("description", ""),
            sap=SapQuerySchema(
                transaction=payload["transaction"],
                fields=payload.get("fields", {}),
                execute_action=payload.get("execute_action", "enter"),
            ),
            download=DownloadSchema(**payload.get("download", {})),
            layout=LayoutSchema(**payload.get("layout", {})),
            edit_rules=payload.get("edit_rules", {}),
            validation=payload.get("validation", {}),
            sampling=SamplingSchema(**payload.get("sampling", {})),
        )
    except KeyError as exc:
        raise HTTPException(status_code=400, detail=f"필수 항목 누락: {exc}") from exc

    path = control_config.save(config, CONFIG_DIR)
    return {"control_id": config.control_id, "path": str(path)}


# ---------- SAP 세션 ----------


@app.get("/api/sap/sessions")
def list_sap_sessions() -> dict[str, Any]:
    try:
        sessions = session_picker.list_sessions()
        return {
            "available": True,
            "sessions": [
                {
                    "session_id": s.session_id,
                    "connection_name": s.connection_name,
                    "transaction": s.transaction,
                }
                for s in sessions
            ],
        }
    except RuntimeError as exc:
        return {"available": False, "sessions": [], "message": str(exc)}


# ---------- 실행 ----------


@app.post("/api/runs")
async def start_run(
    control_id: str = Form(...),
    session_id: str = Form(...),
    condition_file: UploadFile = File(...),
    hr_file: UploadFile = File(...),
) -> dict[str, Any]:
    config = control_config.load(CONFIG_DIR / f"{control_id}.yaml")

    job_id = uuid.uuid4().hex[:12]
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    condition_path = job_dir / "conditions.xlsx"
    hr_path = job_dir / "hr_data.xlsx"
    _save_upload(condition_file, condition_path)
    _save_upload(hr_file, hr_path)

    conditions = reader.read_conditions(condition_path)

    execute = _make_execute_fn(config, session_id, job_dir)
    state_path = job_dir / "job_state.json"
    state = run_batch(job_id, conditions, execute, state_path)

    _JOB_META[job_id] = {
        "control_id": control_id,
        "session_id": session_id,
        "job_dir": str(job_dir),
    }
    _finalize_result(config, state, job_dir)

    return {"job_id": job_id, "summary": state.summary()}


@app.post("/api/runs/{job_id}/retry")
def retry_run(job_id: str) -> dict[str, Any]:
    meta = _get_job_meta(job_id)
    job_dir = Path(meta["job_dir"])
    config = control_config.load(CONFIG_DIR / f"{meta['control_id']}.yaml")
    execute = _make_execute_fn(config, meta["session_id"], job_dir)

    state = retry_failed(execute, job_dir / "job_state.json")
    _finalize_result(config, state, job_dir)

    return {"job_id": job_id, "summary": state.summary()}


@app.get("/api/runs/{job_id}")
def get_run(job_id: str) -> dict[str, Any]:
    meta = _get_job_meta(job_id)
    state = BatchState.load(Path(meta["job_dir"]) / "job_state.json")
    return {"job_id": job_id, "summary": state.summary()}


@app.get("/api/runs/{job_id}/download")
def download_run(job_id: str) -> FileResponse:
    meta = _get_job_meta(job_id)
    result_path = Path(meta["job_dir"]) / "result.xlsx"
    if not result_path.exists():
        raise HTTPException(status_code=404, detail="아직 결과 파일이 없습니다.")
    return FileResponse(result_path, filename=f"{job_id}_result.xlsx")


# ---------- 내부 헬퍼 ----------


def _get_job_meta(job_id: str) -> dict[str, Any]:
    meta = _JOB_META.get(job_id)
    if not meta:
        raise HTTPException(status_code=404, detail="존재하지 않는 작업입니다.")
    return meta


def _save_upload(upload: UploadFile, dest: Path) -> None:
    with dest.open("wb") as f:
        shutil.copyfileobj(upload.file, f)


def _make_execute_fn(config: ControlConfig, session_id: str, job_dir: Path):
    if session_id == MOCK_SESSION_ID:
        return _make_mock_execute_fn(config)

    def execute(condition: dict[str, Any]) -> dict[str, Any]:
        session = session_picker.attach(session_id)
        download_dir = job_dir / "downloads"
        result = transaction_runner.run_condition(session, config, condition, download_dir)
        raw_rows = reader.read_conditions(result["downloaded_path"])
        result["raw_rows"] = raw_rows
        result["rows"] = apply_edit_rules(raw_rows, config.edit_rules)
        return result

    return execute


def _make_mock_execute_fn(config: ControlConfig):
    """SAP GUI가 없는 환경(QA/데모)에서 편집규칙 -> 담당자매칭 -> 검증 -> 결과엑셀까지
    전체 파이프라인을 실제로 확인하기 위한 가짜 실행기. 실제 SAP 세션에는 접근하지 않는다."""
    counter = {"n": 0}

    def execute(condition: dict[str, Any]) -> dict[str, Any]:
        counter["n"] += 1
        row = dict(condition)
        row.setdefault("금액", 1000 * counter["n"])
        row.setdefault("담당자", "EMP001")
        rows = apply_edit_rules([row], config.edit_rules)
        return {
            "downloaded_path": None,
            "condition_capture": None,
            "result_capture": None,
            "raw_rows": [row],
            "rows": rows,
        }

    return execute


_CYCLE_LABELS = {
    "annual": "연간",
    "quarterly": "분기별",
    "monthly": "월별",
    "weekly": "주별",
    "daily": "일별",
}


def _build_sampling_note(sampling: SamplingSchema, sample_size: int, population_count: int) -> str:
    if sampling.control_type == "periodic":
        basis = f"{_CYCLE_LABELS.get(sampling.cycle, sampling.cycle or '')} 통제"
    else:
        basis = f"모집단 {population_count}건 기준"
    return f"{basis} 위험도 최대치 적용 {sample_size}건"


def _finalize_result(config: ControlConfig, state: BatchState, job_dir: Path) -> None:
    raw_all_rows: list[dict[str, Any]] = []
    all_rows: list[dict[str, Any]] = []
    captures: list[dict[str, Any]] = []
    failed_condition_rows: list[dict[str, Any]] = []

    for item in state.items:
        label = f"조건 {item.index + 1}"
        if item.status.value == "success" and item.result:
            raw_all_rows.extend(item.result.get("raw_rows", []))
            all_rows.extend(item.result.get("rows", []))
            captures.append(
                {
                    "label": label,
                    "condition_capture": item.result.get("condition_capture"),
                    "result_capture": item.result.get("result_capture"),
                }
            )
        else:
            # 조회 실패 조건은 정상 데이터가 아니므로 담당자매칭/모집단분류/표본추출
            # 대상에서 제외하고, 바로 예외 목록으로 보낸다 (population/sample에 오류
            # 메시지 행이 섞여 들어가는 것을 방지).
            failed_condition_rows.append(
                {"__조건__": label, EXCEPTION_REASON_COLUMN: "조회실패", "오류": item.error}
            )

    responsible_column = config.validation.get("responsible_column")
    if responsible_column:
        hr_records = reader.read_hr_data(job_dir / "hr_data.xlsx")
        all_rows = match_rows(all_rows, hr_records, responsible_column)

    key_columns = config.validation.get("key_columns", [])
    population_rows, exception_rows = classify_rows(all_rows, key_columns)
    exception_rows = exception_rows + failed_condition_rows

    sample_size = determine_sample_size(
        config.sampling.control_type, config.sampling.cycle, len(population_rows)
    )
    sample_rows = select_sample(population_rows, sample_size)
    sampling_note = _build_sampling_note(config.sampling, sample_size, len(population_rows))

    writer.write_result_excel(
        job_dir / "result.xlsx",
        raw_all_rows,
        population_rows,
        sample_rows,
        exception_rows,
        captures,
        sampling_note=sampling_note,
    )
