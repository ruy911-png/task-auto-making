"""SAP GUI Scripting 세션 조회/선택.

이 모듈은 Windows + SAP GUI(SAP Logon)가 설치되고 스크립팅이 활성화된 환경에서만
실제로 동작한다. 개발/테스트 환경에는 SAP GUI가 없을 수 있으므로, 실제 COM 호출은
함수 안에서 지연 import한다 - 모듈 자체는 어디서든 문제없이 import 가능하다.

로그인은 사용자가 SAP Logon에서 직접 수행한다. 이 모듈은 자격증명을 다루지 않고,
이미 로그인된 세션 목록을 조회해서 사용자가 고르게 할 뿐이다.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class SapSessionInfo:
    connection_index: int
    session_index: int
    connection_name: str
    transaction: str

    @property
    def session_id(self) -> str:
        return f"{self.connection_index}:{self.session_index}"


def list_sessions() -> list[SapSessionInfo]:
    """현재 열려있는 SAP GUI 세션 목록을 조회한다 (사용자가 이미 로그인한 것들)."""
    engine = _get_scripting_engine()
    sessions: list[SapSessionInfo] = []
    for c_idx in range(engine.Connections.Count):
        connection = engine.Connections(c_idx)
        for s_idx in range(connection.Sessions.Count):
            session = connection.Sessions(s_idx)
            sessions.append(
                SapSessionInfo(
                    connection_index=c_idx,
                    session_index=s_idx,
                    connection_name=connection.Description,
                    transaction=session.Info.Transaction,
                )
            )
    return sessions


def attach(session_id: str) -> Any:
    """session_id("연결idx:세션idx")로 지정된, 이미 로그인된 세션에 attach한다."""
    engine = _get_scripting_engine()
    c_idx_str, s_idx_str = session_id.split(":")
    connection = engine.Connections(int(c_idx_str))
    return connection.Sessions(int(s_idx_str))


def _get_scripting_engine() -> Any:
    try:
        import win32com.client
    except ImportError as exc:  # pragma: no cover - Windows+SAP GUI 환경 밖에서는 항상 발생
        raise RuntimeError(
            "pywin32가 설치되어 있지 않거나 SAP GUI Scripting을 사용할 수 없는 환경입니다. "
            "Windows + SAP GUI(SAP Logon)가 설치되고 스크립팅이 활성화된 PC에서 실행하세요."
        ) from exc

    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
    except Exception as exc:  # noqa: BLE001 - COM 에러를 사용자에게 명확한 메시지로 변환
        raise RuntimeError(
            "SAP GUI(SAP Logon)에 연결할 수 없습니다. SAP Logon이 실행 중이고 로그인되어 있는지, "
            "GUI 스크립팅이 활성화되어 있는지 확인하세요."
        ) from exc

    return sap_gui.GetScriptingEngine()
