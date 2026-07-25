# ERP 모집단 자동 추출 도구

SAP GUI Scripting(RPA)으로 조회 조건을 자동 실행하고, 결과를 엑셀로 통합·검증하는
업무 자동화 도구. 자세한 설계 배경은 `.claude/plans/staged-stirring-kite.md` 참고.

## 구성

- `src/sap_automation/` — SAP GUI Scripting 세션 선택, 통제(YAML) 기반 범용 트랜잭션 실행,
  화면 캡처
- `src/batch/` — 여러 조건 순회 실행, 상태 추적, 실패 건만 재처리
- `src/excel_io/` — 조건/인사데이터 엑셀 읽기, 편집 규칙(컬럼정리/필터/계산컬럼) 실행,
  결과 엑셀 쓰기(캡처 이미지 삽입)
- `src/lookup/` — 사번→이름→아이디 순 담당자 매칭
- `src/validation/` — 건수/금액 검증, 중복 체크, 정상/예외 분류
- `src/web/` — 로컬 웹 화면(FastAPI + index.html): 통제 등록, 실행

## 실행 방법

```bash
pip install -r requirements.txt
python src/main.py
```

브라우저에서 http://127.0.0.1:8000 접속.

**중요**: SAP 조회 자동화(sap_automation)는 Windows + SAP GUI(SAP Logon)가 설치되고
스크립팅이 활성화된 PC에서, 사용자가 SAP에 로그인한 세션이 있어야만 동작한다. 이
저장소의 개발 환경에는 SAP GUI가 없으므로, 실행 화면에서 `__mock__` 세션을 선택하면
SAP 없이 나머지 파이프라인(편집규칙 → 담당자매칭 → 검증 → 결과 엑셀)을 확인할 수 있다.

## 새 업무(통제) 추가하기

개발자 코드 수정 없이, 웹 화면의 "통제 등록" 탭에서 거래코드/화면필드/다운로드절차/
편집규칙/검증규칙을 입력하면 `config/transactions/<control_id>.yaml`이 자동 생성된다.
화면 구조가 범용 패턴에 맞지 않는 예외적인 경우에만 `src/sap_automation/overrides/`에
개별 코드를 추가한다.

## 테스트

```bash
pytest
```

SAP GUI Scripting을 직접 호출하는 부분은 이 환경에서 테스트할 수 없어 모킹으로
검증한다. 실제 SAP 연동은 SAP Logon이 설치된 PC에서 재검증이 필요하다.

## 이번 범위 밖 (추후 진행)

- GCP 연동 (결과 아카이빙)
- 내부통제(모집단/샘플링 방법론) 검토
- 보안·테스트 종합 검토
