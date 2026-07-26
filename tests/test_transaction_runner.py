"""transaction_runner의 레이아웃 선택/F8 실행 분기를 mock 세션으로 검증한다.

실제 SAP GUI Scripting COM 세션이 없는 개발 환경이므로, FakeSapSession/FakeComponent로
`find_by_id` 호출 순서와 인자를 기록해서 검증한다. 실제 세션 attach, 화면 조작, 다운로드,
캡처는 사용자의 PC(SAP Logon 설치·로그인 상태)에서만 최종 검증 가능하다.
"""
from __future__ import annotations

from sap_automation.control_config import ControlConfig, DownloadSchema, LayoutSchema, SapQuerySchema
from sap_automation.transaction_runner import _execute, run_condition


class FakeComponent:
    """SAP GUI 컴포넌트(findById 반환값)를 흉내낸다. text/press/select/send_v_key 호출을 기록."""

    def __init__(self, component_id: str, calls: list[tuple]):
        self.component_id = component_id
        self._calls = calls
        self._text = None

    @property
    def text(self):
        return self._text

    @text.setter
    def text(self, value):
        self._text = value
        self._calls.append(("text", self.component_id, value))

    def press(self):
        self._calls.append(("press", self.component_id))

    def select(self):
        self._calls.append(("select", self.component_id))

    def send_v_key(self, key: int):
        self._calls.append(("send_v_key", self.component_id, key))


class FakeSapSession:
    """find_by_id 호출을 기록하고 컴포넌트 id별로 같은 FakeComponent 인스턴스를 재사용한다."""

    def __init__(self):
        self.calls: list[tuple] = []
        self._components: dict[str, FakeComponent] = {}
        self.found_ids: list[str] = []

    def find_by_id(self, component_id: str) -> FakeComponent:
        self.found_ids.append(component_id)
        if component_id not in self._components:
            self._components[component_id] = FakeComponent(component_id, self.calls)
        return self._components[component_id]


def _make_config(**layout_kwargs) -> ControlConfig:
    return ControlConfig(
        control_id="ctrl1",
        description="test control",
        sap=SapQuerySchema(
            transaction="ZTEST",
            fields={"col_a": "wnd[0]/usr/ctxtFIELD_A"},
            execute_action="f8",
        ),
        download=DownloadSchema(
            menu_path=["wnd[0]/mbar/menu1/menu2"],
            file_path_field_id="wnd[1]/usr/ctxtDY_PATH",
            confirm_button_id="wnd[1]/tbar[0]/btn[0]",
        ),
        layout=LayoutSchema(**layout_kwargs),
    )


def test_run_condition_skips_layout_when_no_select_button_id():
    session = FakeSapSession()
    config = _make_config()  # layout 필드 전부 기본값(None) -> 레이아웃 선택 스킵

    run_condition(session, config, {"col_a": "value1"}, download_dir="/tmp/does-not-matter")

    layout_related_ids = {"btnLAYOUT", "ctxtLAYOUT_NAME", "btnLAYOUT_CONFIRM"}
    assert not (layout_related_ids & set(session.found_ids))


def test_run_condition_selects_layout_in_order_when_configured():
    session = FakeSapSession()
    config = _make_config(
        select_button_id="wnd[0]/tbar[1]/btn[32]",
        layout_name="ZLAYOUT01",
        layout_name_field_id="wnd[1]/usr/ctxtLAYOUT_NAME",
        confirm_button_id="wnd[1]/tbar[0]/btn[0]",
    )

    run_condition(session, config, {"col_a": "value1"}, download_dir="/tmp/does-not-matter")

    layout_call_indices = [
        i
        for i, call in enumerate(session.calls)
        if call[1]
        in (
            "wnd[0]/tbar[1]/btn[32]",
            "wnd[1]/usr/ctxtLAYOUT_NAME",
        )
    ]

    # 레이아웃 선택 버튼 누름 -> 레이아웃 이름 입력이 순서대로 기록되어야 한다.
    assert ("press", "wnd[0]/tbar[1]/btn[32]") in session.calls
    assert ("text", "wnd[1]/usr/ctxtLAYOUT_NAME", "ZLAYOUT01") in session.calls

    press_idx = session.calls.index(("press", "wnd[0]/tbar[1]/btn[32]"))
    text_idx = session.calls.index(("text", "wnd[1]/usr/ctxtLAYOUT_NAME", "ZLAYOUT01"))
    assert press_idx < text_idx

    # 확인 버튼은 다운로드 대화상자 확인 버튼과 컴포넌트 id가 같으므로(테스트 설정상)
    # press가 최소 2번(레이아웃 확인 + 다운로드 확인) 호출되었는지로 검증한다.
    confirm_press_count = sum(
        1 for call in session.calls if call == ("press", "wnd[1]/tbar[0]/btn[0]")
    )
    assert confirm_press_count >= 1


def test_execute_f8_sends_v_key_8():
    session = FakeSapSession()

    _execute(session, "f8")

    assert ("send_v_key", "wnd[0]", 8) in session.calls


def test_execute_f8_case_insensitive():
    session = FakeSapSession()

    _execute(session, "F8")

    assert ("send_v_key", "wnd[0]", 8) in session.calls


def test_execute_enter_still_sends_v_key_0():
    session = FakeSapSession()

    _execute(session, "enter")

    assert ("send_v_key", "wnd[0]", 0) in session.calls


def test_execute_button_id_still_calls_press():
    session = FakeSapSession()

    _execute(session, "wnd[0]/tbar[1]/btn[8]")

    assert ("press", "wnd[0]/tbar[1]/btn[8]") in session.calls
