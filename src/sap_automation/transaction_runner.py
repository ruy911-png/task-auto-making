"""통제 설정(YAML)을 읽어 SAP GUI Scripting을 범용적으로 구동하는 엔진.

새 통제를 추가할 때 이 파일을 수정하지 않는다 - control_config.ControlConfig에
정의된 필드 매핑/다운로드 절차를 그대로 실행하는 범용 실행기다. 화면 구조가
이 범용 패턴(필드 입력 -> 실행 -> 메뉴로 다운로드)에 맞지 않는 예외적인 경우만
`overrides/`에 개별 코드를 추가한다.
"""
from __future__ import annotations

import time
from pathlib import Path
from typing import Any, Callable, Protocol

from .control_config import ControlConfig, DownloadSchema, LayoutSchema


class SapSession(Protocol):
    """transaction_runner가 필요로 하는 최소 인터페이스.

    실제 구현은 SAP GUI Scripting COM 세션 객체(`find_by_id`가 `session.findById`에
    대응)이고, 테스트에서는 FakeSapSession으로 대체한다.
    """

    def find_by_id(self, component_id: str) -> Any: ...


CaptureFn = Callable[[str], str]


def run_condition(
    session: SapSession,
    config: ControlConfig,
    condition: dict[str, Any],
    download_dir: str | Path,
    capture_fn: CaptureFn | None = None,
) -> dict[str, Any]:
    """조회 조건 하나를 실행한다: 트랜잭션 이동 -> 필드 입력 -> 조회조건 캡처 -> 실행
    -> 결과 캡처 -> SAP 엑셀 다운로드.

    capture_fn: 라벨을 받아 캡처 파일 경로(str)를 반환하는 함수. 실제 환경에서는
    `sap_automation.capture.capture_window`를 감싸서 주입하고, SAP가 없는 테스트/CI
    환경에서는 주입하지 않거나(None) no-op을 준다.
    """
    _start_transaction(session, config.sap.transaction)
    _fill_fields(session, config.sap.fields, condition)

    condition_capture = capture_fn(f"{config.control_id}_condition") if capture_fn else None

    _execute(session, config.sap.execute_action)

    _select_layout(session, config.layout)

    result_capture = capture_fn(f"{config.control_id}_result") if capture_fn else None

    download_dir = Path(download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)
    run_id = str(int(time.time() * 1000))
    file_name = config.download.file_name_pattern.format(
        control_id=config.control_id, run_id=run_id
    )
    downloaded_path = download_dir / file_name
    _download_excel(session, config.download, downloaded_path)

    return {
        "downloaded_path": str(downloaded_path),
        "condition_capture": condition_capture,
        "result_capture": result_capture,
    }


def _start_transaction(session: SapSession, transaction: str) -> None:
    ok_field = session.find_by_id("wnd[0]/tbar[0]/okcd")
    ok_field.text = f"/n{transaction}"
    session.find_by_id("wnd[0]").send_v_key(0)


def _fill_fields(
    session: SapSession, field_mapping: dict[str, str], condition: dict[str, Any]
) -> None:
    for condition_column, sap_field_id in field_mapping.items():
        value = condition.get(condition_column)
        if value is None:
            continue
        component = session.find_by_id(sap_field_id)
        component.text = str(value)


def _execute(session: SapSession, execute_action: str) -> None:
    action = execute_action.lower() if isinstance(execute_action, str) else execute_action
    if action == "enter":
        session.find_by_id("wnd[0]").send_v_key(0)
    elif action == "f8":
        session.find_by_id("wnd[0]").send_v_key(8)
    else:
        session.find_by_id(execute_action).press()


def _select_layout(session: SapSession, layout: LayoutSchema) -> None:
    """조회 실행 후, 엑셀 다운로드 전에 통제별로 지정된 레이아웃을 선택한다.

    select_button_id가 없으면 레이아웃 선택이 필요 없는 통제이므로 아무것도 하지 않는다.
    """
    if not layout.select_button_id:
        return

    session.find_by_id(layout.select_button_id).press()

    if layout.layout_name_field_id and layout.layout_name:
        session.find_by_id(layout.layout_name_field_id).text = layout.layout_name

    if layout.confirm_button_id:
        session.find_by_id(layout.confirm_button_id).press()


def _download_excel(session: SapSession, download_schema: DownloadSchema, output_path: Path) -> None:
    """SAP 자체 엑셀 다운로드(스프레드시트 내보내기) 메뉴를 자동으로 실행한다 (고정 절차 —
    데이터 건수와 무관하게 항상 이 방식을 사용, 그리드를 셀 단위로 읽지 않는다).

    menu_path의 정확한 컴포넌트 id는 환경마다 다르므로, 통제 등록 화면에서
    SAP GUI Scripting 레코딩으로 확인한 값을 입력해야 한다. 이 함수는 그 설정을
    그대로 실행하는 역할만 한다.
    """
    for menu_item_id in download_schema.menu_path:
        session.find_by_id(menu_item_id).select()

    if download_schema.file_path_field_id:
        session.find_by_id(download_schema.file_path_field_id).text = str(output_path)
    if download_schema.confirm_button_id:
        session.find_by_id(download_schema.confirm_button_id).press()
