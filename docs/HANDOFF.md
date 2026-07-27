# 인계 문서 — ERP 모집단 자동 추출 도구 (Gemini CLI / Vertex AI 인계용)

이 문서는 Claude Code(원격 샌드박스)에서 진행한 작업을 사내 환경(Gemini CLI / Vertex AI)에
넘겨 이어서 진행하기 위한 인계서다. 원격 환경에는 SAP GUI가 없어 실제 SAP 검증은 처음부터
불가능했고, PRD·프로토타입·비-SAP 로직(엑셀 가공/표본추출/웹화면)은 여기서 완성했지만
**실제 SAP 환경에서의 최종 확인은 전부 이 문서 하단 "여기서부터 할 일"에 남아있다.**

## 저장소 / 브랜치

- 저장소: `ruy911-png/task-auto-making`
- 브랜치: `claude/agent-team-structure-prd-6b1k93`
- 최신 커밋: `9a1a69a` (2026-07-26 기준)
- 문서: `docs/PRD.md` (제품 요구사항 전체), 이 파일(`docs/HANDOFF.md`)
- 코드 번들 `task-auto-making-source.zip`의 SHA256:
  `12bae887638e69f4411d4b1ae82a6f5366c68b51501660e71c46f18232cc5bc9`
  (받은 파일이 이 값과 다르면 잘못된 파일이니 다시 받을 것)

## 시작하기 (압축 해제 + 버전 확인)

```bash
sha256sum task-auto-making-source.zip   # 위 해시값과 일치하는지 먼저 확인
unzip task-auto-making-source.zip -d task-auto-making
cd task-auto-making
grep -c "data-tab" src/web/static/index.html   # 3이 나와야 최신 버전 (통제등록/담당자매칭/실행)
```

## Gemini CLI 최초 지시문 (그대로 복붙해서 사용)

```
이 프로젝트는 Claude Code로 개발하다가 넘겨받은 것이다. docs/HANDOFF.md와 docs/PRD.md를
먼저 읽고 전체 맥락을 파악해라. 특히 HANDOFF.md의 "여기서부터 할 일" 절을 따른다.

지금 당장 할 일은 A번(실제 SAP 환경 확인)이다:
1. SAP Basis팀에 sapgui/user_scripting 파라미터가 TRUE인지 확인 결과를 나에게 물어봐라.
   확인됐다고 답하면 다음 단계로 넘어가라.
2. 통제 하나를 실제로 등록할 건데, 대상 t-code는 [여기에 실제 t-code 입력, 예: FB03]다.
   이 화면이 Classic Dynpro인지 먼저 확인해라.
3. SAP Logon에서 Alt+F12 → Script Recording and Playback으로 조회조건 입력을 녹화해서
   실제 SAP 필드 id를 알아내고, config/transactions/<control_id>.yaml에 반영해라.
4. src/web/backend.py를 로컬로 띄우고(`PYTHONPATH=src python -m uvicorn web.backend:app
   --port 8000`), 웹 화면(이미 완성되어 있음 — 새로 만들 필요 없음)에서 이 통제를 실제로
   실행해서 결과 엑셀이 정상 생성되는지 end-to-end로 확인해라. 실패하면 원인을 나에게
   보고하고 수정안을 물어봐라.

코드를 마음대로 바꾸기 전에 항상 나한테 먼저 물어봐라 — Claude Code 세션에서도 그렇게
진행했다.
```

`[여기에 실제 t-code 입력]` 부분은 실제로 먼저 등록하고 싶은 업무의 t-code로 바꿀 것.
B번(담당자 보조조회 구현)·C번(알려진 이슈 해결)은 A번이 끝난 뒤 별도로 지시하면 된다.

## 한 줄 요약

감사/회계 부서가 반복하는 "모집단 추출" 업무(SAP 조회 → 엑셀 정리 → 담당자 확인 → 검증
→ 표본추출)를 SAP GUI Scripting 기반 RPA + 파이썬 엑셀 파이프라인으로 자동화하는 로컬
웹 도구. 통제(업무)마다 코드 수정 없이 YAML 설정만으로 확장 가능하게 설계됨.

## 아키텍처 한눈에

```
config/transactions/<control_id>.yaml   # 통제(업무) 하나 = 조회+다운로드+레이아웃+
                                         # 추가화면+편집규칙+검증+매칭 방식을 담은 설정

src/sap_automation/   # SAP GUI Scripting (Windows 전용, 이 환경에선 실행 불가 — 모킹 검증만)
  session_picker.py     - 로그인된 SAP 세션 목록 조회/attach
  control_config.py     - 통제 스키마 정의 + YAML 입출력 (SapQuerySchema/LayoutSchema/
                           AdditionalScreenSchema/MatchingSchema/ControlConfig)
  transaction_runner.py - 설정을 그대로 실행하는 범용 엔진 (본화면입력→추가화면팝업→
                           F8실행→레이아웃선택→결과캡처→엑셀다운로드)
  capture.py             - 화면 캡처 (pywin32+Pillow 필요, Windows 전용)

src/batch/runner.py    # 조건 엑셀 여러 행 순회 실행, 조건별 성공/실패 추적, 실패건 재처리

src/excel_io/          # 엑셀 입출력 (SAP 의존성 없음 — 이 환경에서 완전히 검증됨)
  reader.py, writer.py, edit_rules.py, schema.py

src/lookup/             # 담당자 매칭 (3가지 방식 중 2가지 구현됨)
  hr_matcher.py          - ① 인사데이터 매칭 (사번→이름→아이디)
  reference_matcher.py   - ② 별도 참조 엑셀 업로드 매칭 (SAP 미접촉)
                          - ③ SAP 보조조회는 미구현 (아래 "할 일" 참고)

src/validation/
  checks.py    - 건수·금액 검증, 중복체크, 정상/예외 분류
  sampling.py  - 표본 수 산정 (모집단 건수 기준 표 하나, determine_sample_size/select_sample)

src/web/
  backend.py        - FastAPI: 통제 등록/조회, SAP 세션 목록, 배치 실행, 재처리, 다운로드
  static/index.html - 통제등록 / 담당자매칭 / 실행 3탭 (실제 화면, 프로토타입 아님)
  static/app.js
  static/style.css
```

## 지금까지 확정된 핵심 설계 결정 (재논의 없이 그대로 따를 것)

1. **SAP 자동화 방식**: SAP GUI Scripting(pywin32) 확정. RFC/HANA 직접접속은 보안정책상
   불가. OData API·iPaaS·ABAP 등 대안은 검토했으나 채택 안 함 — GUI Scripting이 이미
   사내에 유사 사례(로봇 자동화)가 있어 실현 가능성 확인됨.
2. **실행은 항상 F8 고정** — 실사용 통제가 전부 F8이라 통제별 실행방식 선택 UI 자체를 없앰.
3. **담당자 매칭 순서**: 표본 뽑기 전에 **모집단 전체**를 먼저 매칭한다 (사람이 수작업할 때
   쓰는 "표본 먼저 뽑고 매칭" 관행은 자동화에는 해당 없음 — 전수 매칭에 인력비용이 안 들기
   때문).
4. **표본 수 산정**: 통제의 명목상 주기(월별/분기별 등)를 따로 등록하지 않고, **실행마다
   나온 실제 모집단 건수 하나만**으로 산정한다 (평가회차별 기간분할·연간표본배분 로직은
   설계 검토 후 불필요 판단, 제외). 근거: 한국회계기준원 내부회계관리제도 모범규준 적용
   FAQ 34, 위험도 항상 최대치 적용. 표는 `docs/PRD.md`의 "표본 수 산정 기준" 절 참고.
5. **결과 엑셀 5탭**: 캡처 / Raw data / 모집단 data / 샘플 data / 예외 data.
6. **조회 실패 건**은 데이터가 아니므로 매칭/모집단/표본추출 대상에서 제외하고 곧바로
   예외(사유: 조회실패)로 분류한다 — QA 재검증 중 발견된 결함을 고친 것.

## 여기서부터 할 일 (Gemini CLI / Vertex AI + 실제 SAP 환경에서)

### A. 반드시 실제 SAP 환경에서만 확인 가능 (이 원격 환경은 SAP GUI가 없어 불가능했음)
1. **SAP GUI Scripting 활성화 여부 확인** — Basis팀에 `sapgui/user_scripting=TRUE`
   (RZ11 또는 프로파일 파일) 설정 여부 확인, 클라이언트 SAP Logon의 Scripting 체크박스도
   켜져 있어야 함.
2. **통제별 실제 SAP 필드 id 확인** — 등록화면에 입력하는 `wnd[0]/usr/...` 같은 값은
   전부 예시(placeholder)다. 실제 값은 SAP GUI Scripting 레코딩(Alt+F12 → Script
   Recording and Playback)으로 통제마다 직접 확인해야 한다.
3. **대상 t-code가 Classic Dynpro인지 확인** — Fiori/WebDynpro 화면은 이 방식 자체가
   안 먹는다. 통제 등록 전 개별 확인 필요.
4. 위 확인이 끝나면 실제 통제 1개를 등록해서 조건 엑셀 실행 → 결과 엑셀 정상 생성까지
   end-to-end 검증 (PRD "성공 기준" 참고).

### B. 설계는 끝났고 코드만 남은 것
- **담당자 매칭 ③ SAP 보조조회 구현** — `ControlConfig.matching.method == "sap_lookup"`을
  선택하면 지금은 `/api/runs`가 400을 반환하도록만 막아뒀다(`src/web/backend.py`). 설계는
  확정됨: 전표번호 등을 키로 `transaction_runner` 엔진을 재사용해 두 번째 SAP 트랜잭션을
  조회하면 됨 — 구조는 `AdditionalScreenSchema`/`LayoutSchema`와 같은 패턴(옵션 필드,
  없으면 스킵)으로 넣는 게 일관적이다.

### C. 기존에 알려졌지만 아직 미해결 (PRD "알려진 이슈" 표 참고)
- `/api/runs`가 배치 실행을 블로킹 (동기 실행을 async 핸들러에서 직접 호출 — 백그라운드
  작업 + job_id 폴링 방식으로 바꿔야 함)
- `validate_totals`(건수·금액 기대값 대사)가 구현은 됐는데 파이프라인에서 호출 안 됨
- 편집규칙 필터가 문자열 비교라 숫자 필터가 틀릴 수 있음
- 통제 등록 시 동일 ID 덮어쓰기 경고 없음
- 다중 사용자 동시 사용 미고려 (작업 이력이 서버 메모리에만 존재)
- 보안 검토 전무 (인증, 업로드 크기 제한 등) — 이번 범위 밖, 별도 보안 검토 필요

### D. 이번 범위에서 의도적으로 제외한 것
- GCP 연동(결과 아카이빙) — 바로 이 인계 작업의 다음 단계가 될 수 있음
- 표본추출 기준표 자체의 방법론 타당성 검증 (사용자가 확정한 기준을 그대로 구현했을 뿐)

## 검증 상태

- `python -m pytest tests/` — 53개 전체 통과 (SAP·GCP 의존성 없는 부분은 전부 모킹으로
  검증됨. `src/sap_automation/*`의 실제 SAP GUI 조작, `src/sap_automation/capture.py`의
  화면 캡처는 Windows+SAP 환경에서만 최종 검증 가능).
- 웹 화면은 로컬에서 서버를 직접 띄우고 Playwright로 등록→담당자매칭 방식 전환→실행까지
  실제 클릭해서 확인함 (콘솔 에러 없음, mock 세션으로 1/1 성공 확인).

## 실행 방법 (사내 환경에서 그대로 사용)

```bash
pip install -r requirements.txt
# Windows + SAP GUI(SAP Logon) 설치·로그인 상태에서:
PYTHONPATH=src python -m uvicorn web.backend:app --port 8000
# 브라우저에서 http://localhost:8000 접속
```

SAP GUI Scripting 관련 실제 동작(`session_picker`, `transaction_runner`, `capture`)은
Windows + SAP Logon 환경에서만 동작한다. 이 원격 환경에서는 `__mock__` 세션으로 나머지
파이프라인(편집규칙→담당자매칭→검증→표본추출→결과엑셀)만 검증했다.
