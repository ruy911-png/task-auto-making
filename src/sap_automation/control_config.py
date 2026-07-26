"""통제(RPA 대상 업무) 스키마 정의 및 YAML 입출력.

하나의 "통제"는 조회(SAP 필드 매핑) + 다운로드(절차) + 편집(컬럼정리/필터/계산컬럼) +
검증(키컬럼/금액컬럼) 스키마를 하나의 YAML로 표현한다. 신규 통제를 추가할 때는
이 파일 하나만 추가하면 되고, 코드를 수정할 필요가 없다 (웹 등록 화면이 이 YAML을
자동 생성한다).
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any

import yaml


@dataclass
class SapQuerySchema:
    transaction: str
    fields: dict[str, str] = field(default_factory=dict)
    """조건 엑셀의 컬럼명 -> SAP 화면 필드 컴포넌트 id 매핑."""
    execute_action: str = "f8"
    """조회 실행 방식. 실사용 통제가 전부 F8이라 기본값으로 고정 — 예외적인 화면만
    "enter" 또는 특정 버튼의 컴포넌트 id(예: "wnd[0]/tbar[1]/btn[8]")로 override."""


@dataclass
class AdditionalScreenSchema:
    """본 화면 입력 후, 팝업(추가 조회조건 — 예: SAP Multiple Selection의 "계단" 아이콘)을
    열어 필드를 추가로 입력하는 단계 (통제별 선택사항, 예: FB03에서 문서유형 추가 조회).

    open_button_id가 없으면 이 단계 자체를 건너뛴다 (추가화면이 필요 없는 통제도 있음).
    """

    open_button_id: str | None = None
    """팝업을 여는 버튼/아이콘의 컴포넌트 id."""
    fields: dict[str, str] = field(default_factory=dict)
    """조건 엑셀 컬럼 -> 팝업 화면 필드 컴포넌트 id 매핑 (본 화면 fields와 별개 네임스페이스)."""
    confirm_button_id: str | None = None
    """팝업 입력 확인/본 화면 복귀 버튼의 컴포넌트 id."""


@dataclass
class DownloadSchema:
    menu_path: list[str] = field(default_factory=list)
    """SAP 엑셀 다운로드(스프레드시트 내보내기) 메뉴 항목 id를 누르는 순서."""
    file_path_field_id: str | None = None
    """파일 저장 대화상자의 경로 입력 필드 id (있는 경우)."""
    confirm_button_id: str | None = None
    """파일 저장 대화상자의 확인 버튼 id (있는 경우)."""
    file_name_pattern: str = "{control_id}_{run_id}.xlsx"


@dataclass
class LayoutSchema:
    """조회 실행 후, 엑셀 다운로드 전에 특정 레이아웃을 선택하는 단계 (통제별로 명기).

    select_button_id가 없으면 레이아웃 선택 단계 자체를 건너뛴다 (레이아웃 선택이
    필요 없는 통제도 있으므로 선택 사항으로 둔다).
    """

    select_button_id: str | None = None
    """'레이아웃 선택'을 여는 버튼/메뉴의 컴포넌트 id."""
    layout_name: str | None = None
    """선택할 레이아웃의 이름/코드 (통제 등록 시 값으로 명기)."""
    layout_name_field_id: str | None = None
    """레이아웃 이름을 입력/선택하는 필드의 컴포넌트 id."""
    confirm_button_id: str | None = None
    """레이아웃 선택 확인 버튼의 컴포넌트 id."""


@dataclass
class ControlConfig:
    control_id: str
    description: str
    sap: SapQuerySchema
    download: DownloadSchema = field(default_factory=DownloadSchema)
    layout: LayoutSchema = field(default_factory=LayoutSchema)
    additional_screen: AdditionalScreenSchema = field(default_factory=AdditionalScreenSchema)
    edit_rules: dict[str, Any] = field(default_factory=dict)
    """rename_columns / drop_columns / filters / calculated_columns 지원 (excel_io.edit_rules 참고)."""
    validation: dict[str, Any] = field(default_factory=dict)
    """key_columns / amount_column 등 (validation.checks 참고)."""
    # 표본추출은 실행 시점 모집단 건수 하나로만 결정된다(validation.sampling 참고) —
    # 통제별 설정이 필요 없어 여기 별도 스키마를 두지 않는다.

    def to_dict(self) -> dict[str, Any]:
        return {
            "control_id": self.control_id,
            "description": self.description,
            "sap": asdict(self.sap),
            "download": asdict(self.download),
            "layout": asdict(self.layout),
            "additional_screen": asdict(self.additional_screen),
            "edit_rules": self.edit_rules,
            "validation": self.validation,
        }

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> "ControlConfig":
        return cls(
            control_id=d["control_id"],
            description=d.get("description", ""),
            sap=SapQuerySchema(**d["sap"]),
            download=DownloadSchema(**d.get("download", {})),
            layout=LayoutSchema(**d.get("layout", {})),
            additional_screen=AdditionalScreenSchema(**d.get("additional_screen", {})),
            edit_rules=d.get("edit_rules", {}) or {},
            validation=d.get("validation", {}) or {},
        )


def save(config: ControlConfig, directory: str | Path) -> Path:
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"{config.control_id}.yaml"
    path.write_text(
        yaml.safe_dump(config.to_dict(), allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    return path


def load(path: str | Path) -> ControlConfig:
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
    return ControlConfig.from_dict(data)


def load_all(directory: str | Path) -> list[ControlConfig]:
    directory = Path(directory)
    if not directory.exists():
        return []
    return [load(p) for p in sorted(directory.glob("*.yaml"))]
