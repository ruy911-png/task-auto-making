# 인계 문서 — ERP 모집단 자동 추출 도구 (Gemini CLI / Vertex AI 인계용)

이 문서는 Claude Code(원격 샌드박스)에서 진행한 작업을 사내 환경(Gemini CLI / Vertex AI)에
넘겨 이어서 진행하기 위한 인계서다. 원격 환경에는 SAP GUI가 없어 실제 SAP 검증은 처음부터
불가능했고, PRD·프로토타입·비-SAP 로직(엑셀 가공/표본추출/웹화면)은 여기서 완성했지만
**실제 SAP 환경에서의 최종 확인은 전부 이 문서 하단 "여기서부터 할 일"에 남아있다.**

## ⚠️ 먼저 읽을 것 — GCP/Vertex AI에서 되는 것 vs 안 되는 것

이 프로젝트는 성격이 다른 두 부분으로 이루어져 있다. **"GCP에서 실행한다"는 말이 "SAP
자동화까지 GCP에서 다 된다"는 뜻이 절대 아니다.** 아래 표로 명확히 구분한다.

| 구분 | 내용 | GCP/Vertex AI(클라우드)에서 | Windows + SAP Logon 로그인된 PC에서 |
|---|---|---|---|
| **SAP 자동화** | `src/sap_automation/*` — SAP 화면에 직접 값을 입력하고 실행 버튼을 누르는 부분(SAP GUI Scripting) | ❌ 불가능. SAP GUI 프로그램 자체가 클라우드 서버에 설치되어 있지 않고, 설치해도 로그인 세션이 없어 동작할 수 없다 | ✅ 유일하게 여기서만 동작 |
| **웹 화면 + 엑셀 처리** | `src/web`, `src/excel_io`, `src/validation`, `src/lookup` — 화면(통제등록/담당자매칭/실행), 엑셀 읽기/쓰기/편집규칙, 담당자 매칭, 건수·금액 검증, 표본추출 | ✅ 가능. `__mock__`(가짜 SAP 세션)으로 SAP 없이도 전체 파이프라인을 끝까지 눈으로 확인할 수 있다 | ✅ 가능 (SAP 세션도 실제로 붙일 수 있음) |

**결론**: GCP/Vertex AI 쪽 작업은 "웹 화면이 정상적으로 뜨는지, 엑셀 처리 로직이 맞는지"까지만
확인할 수 있다. 실제 SAP 조회 자동화는 **반드시 Windows + SAP Logon이 켜져 있는 회사 PC**에서
따로 확인해야 한다 (아래 "여기서부터 할 일 → A" 참고). 이 순서를 착각해서 GCP에서 SAP 연동을
기대하며 시간을 쓰지 말 것.

## 저장소 / 브랜치

- 저장소: `ruy911-png/task-auto-making`
- 브랜치: `claude/agent-team-structure-prd-6b1k93`
- 이 문서를 작성/정리한 커밋: `<COMMIT_HASH>` — **주의**: 이 zip 파일은 위 커밋에서
  `git archive`로 만들었다. 그런데 이 문서(HANDOFF.md) 자신이 그 커밋에 포함되어 있어서,
  "이 문서에 적힌 SHA256 값"과 "이 문서가 실제로 포함된 최종 커밋의 해시"는 정의상 완벽히
  동시에 일치시킬 수 없다(문서에 값을 적는 행위 자체가 그다음 커밋이 되기 때문). 이후에 이
  안내문 자체를 다듬는 사소한 커밋이 하나 더 있을 수 있는데, **코드는 동일하다.** 진짜 최신
  버전인지는 아래 "버전 확인" 절의 커맨드로 판단할 것 — 커밋 해시 자체보다 그게 더 정확하다.
- 문서: `docs/PRD.md` (제품 요구사항 전체), 이 파일(`docs/HANDOFF.md`)
- 코드 번들 `task-auto-making-source.zip`의 SHA256:
  `<ZIP_SHA256>`
  (받은 파일이 이 값과 다르면 전송 중 손상됐거나 잘못된 파일이니 다시 받을 것)

## 시작하기 — 여기부터 그대로 따라 하면 됨 (비개발자도 가능)

### 1단계. 압축 풀기 + 버전 확인

터미널(명령 프롬프트/Terminal)을 열고 zip 파일이 있는 위치에서 아래를 순서대로 입력한다.

```bash
sha256sum task-auto-making-source.zip   # 위 "SHA256" 값과 똑같은 값이 나와야 정상
unzip task-auto-making-source.zip -d task-auto-making
cd task-auto-making
grep -c "data-tab" src/web/static/index.html   # 3이 나와야 최신 버전 (통제등록/담당자매칭/실행)
ls run.sh                                       # run.sh 파일이 보여야 최신 버전
```

두 값이 문서와 다르면 최신 버전이 아니니 다시 받아야 한다. **이 두 커맨드가 진짜 검증
수단이다 — "예전 버전 같다"는 느낌만으로 판단하지 말고 반드시 이 커맨드를 돌려서 확인할 것.**

### 2단계. 서버 실행 — `run.sh` 스크립트 하나만 실행하면 됨

"서버를 실행한다"는 것은, 이 프로그램을 컴퓨터 안에서 계속 켜놓고 웹 브라우저로 접속할 수
있게 대기시켜 놓는다는 뜻이다. 아래 한 줄만 터미널에 그대로 붙여넣고 Enter를 누르면 된다.
(패키지 설치까지 이 스크립트 하나가 전부 처리한다 — 별도로 `pip install`을 따로 칠 필요 없음.)

```bash
bash run.sh
```

실행하면 화면에 메시지가 몇 줄 지나가다가 마지막에

```
Uvicorn running on http://0.0.0.0:8000 (Press CTRL+C to quit)
```

같은 줄이 뜨고 커서가 깜빡이며 "멈춘 것처럼" 보인다. **이건 오류가 아니라 정상이다.**
프로그램이 계속 켜진 채로 대기 중이라는 뜻이다. 이 터미널 창을 닫지 말고 그대로 둔다.

### 3단계. 브라우저로 접속

이 상태에서 웹 브라우저(크롬 등)를 새로 열고, 주소창에 아래를 입력한다.

```
http://localhost:8000
```

**주의**: `index.html` 파일을 탐색기/Finder에서 더블클릭해서 여는 것과는 다르다. 파일을
직접 열면 서버를 거치지 않아서 디자인(스타일)이 하나도 안 입혀진 화면만 보이고, 버튼을
눌러도 아무 반응이 없다. 반드시 위 2단계로 서버를 먼저 켠 다음, 브라우저 주소창에
`http://localhost:8000`을 입력해서 접속해야 정상 화면이 나온다.

GCP Vertex AI Workbench 같은 원격 서버에서 `run.sh`를 실행하는 경우, `localhost`가 아니라
그 서버로 접속해야 하므로 8000번 포트를 브라우저로 전달하는 포트포워딩/웹 프리뷰 기능이
필요할 수 있다 (Vertex AI Workbench의 "Open JupyterLab" 프록시, `gcloud compute ssh -L
8000:localhost:8000 ...` 같은 방식). 이 부분은 실제 GCP 계정/네트워크 설정에 따라 달라지므로
Gemini CLI가 해당 환경 문서를 참고해 안내해야 한다 (아래 "여기서부터 할 일" 참고).

### 4단계. 끝내기

다 확인했으면 서버를 켜둔 터미널 창에서 `Ctrl+C`를 누르면 종료된다.

### (참고) `run.sh`가 내부에서 실제로 하는 일

비개발자는 몰라도 되지만, 개발자가 궁금하면: `run.sh`는
`pip install -r requirements.txt`로 필요한 패키지를 설치한 뒤,
`PYTHONPATH=src python -m uvicorn web.backend:app --host 0.0.0.0 --port 8000`을 실행하는
셸 스크립트다. 저장소 루트의 `run.sh` 파일을 열어보면 각 줄에 한글 주석이 달려 있다.

## Gemini CLI 최초 지시문 (그대로 복붙해서 사용)

```
이 프로젝트는 Claude Code로 개발하다가 넘겨받은 것이다. docs/HANDOFF.md와 docs/PRD.md를
먼저 읽고 전체 맥락을 파악해라. 특히 HANDOFF.md 맨 위 "GCP/Vertex AI에서 되는 것 vs 안
되는 것" 표를 반드시 먼저 이해해라 — SAP 자동화(src/sap_automation)는 Windows + SAP Logon
PC에서만 동작하고, 이 GCP/Gemini CLI 환경에서는 절대 SAP 연동을 시도하지 마라. 여기서 할 수
있는 건 웹 화면 + 엑셀 파이프라인을 __mock__ 세션으로 띄워서 확인하는 것까지다.

가장 먼저 할 일:
1. 저장소 루트에서 `bash run.sh` 를 실행해라. 패키지 설치와 서버 기동을 이 스크립트 하나가
   전부 처리한다. "Uvicorn running on http://0.0.0.0:8000" 메시지가 뜨면 정상이다.
2. 사용자에게 "서버가 켜졌습니다. 웹 브라우저에서 http://localhost:8000 (원격 환경이면
   해당 환경의 포트포워딩/웹 프리뷰 방식으로 8000번 포트에)를 접속하시면 됩니다"라고
   안내해라. index.html 파일을 직접 더블클릭해서 열면 스타일이 깨져 보이니 그렇게 하지
   말라고 반드시 같이 안내해라.
3. 그다음 HANDOFF.md의 "여기서부터 할 일" 절을 따른다. 지금 당장 할 일은 A번(실제 SAP
   환경 확인)인데, 이건 이 GCP 환경이 아니라 별도의 Windows + SAP Logon PC에서 진행해야
   한다:
   a. SAP Basis팀에 sapgui/user_scripting 파라미터가 TRUE인지 확인 결과를 사용자에게
      물어봐라. 확인됐다고 답하면 다음 단계로 넘어가라.
   b. 통제 하나를 실제로 등록할 건데, 대상 t-code는 [여기에 실제 t-code 입력, 예: FB03]다.
      이 화면이 Classic Dynpro인지 먼저 확인해라.
   c. SAP Logon에서 Alt+F12 → Script Recording and Playback으로 조회조건 입력을 녹화해서
      실제 SAP 필드 id를 알아내고, config/transactions/<control_id>.yaml에 반영해라.
   d. Windows PC에서 `run.sh`(또는 동일한 명령)로 서버를 띄우고, 웹 화면(이미 완성되어
      있음 — 새로 만들 필요 없음)에서 이 통제를 실제 SAP 세션으로 실행해서 결과 엑셀이
      정상 생성되는지 end-to-end로 확인해라. 실패하면 원인을 사용자에게 보고하고 수정안을
      물어봐라.

협업 원칙 (이번 세션에서 겪은 것 그대로 지켜라):
- 코드를 마음대로 바꾸기 전에 항상 사용자에게 먼저 물어봐라.
- 사용자는 비개발자에 가깝다. 터미널 명령을 안내할 때는 "이 명령이 무엇을 하는지"를 명령어
  앞에 한 줄로 먼저 설명하고, "실행하면 화면이 이렇게 보일 것이다/이건 정상이다" 같은 결과
  예측까지 같이 알려줘라. "서버가 켜졌다"는 것을 당연히 아는 개념으로 취급하지 마라.
- "이전 버전 아니냐"는 질문이 나오면 감으로 답하지 말고 위 "버전 확인" 절의 grep/sha256sum
  커맨드를 실제로 실행시켜서 숫자로 확인시켜줘라.
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

run.sh                # 서버 실행 스크립트 (패키지 설치 + uvicorn 기동을 한 번에 처리)

src/sap_automation/   # SAP GUI Scripting (Windows + SAP Logon 전용, GCP/이 환경에선
                       # 실행 불가 — __mock__ 세션으로 나머지 파이프라인만 모킹 검증)
  session_picker.py     - 로그인된 SAP 세션 목록 조회/attach
  control_config.py     - 통제 스키마 정의 + YAML 입출력 (SapQuerySchema/LayoutSchema/
                           AdditionalScreenSchema/MatchingSchema/ControlConfig)
  transaction_runner.py - 설정을 그대로 실행하는 범용 엔진 (본화면입력→추가화면팝업→
                           F8실행→레이아웃선택→결과캡처→엑셀다운로드)
  capture.py             - 화면 캡처 (pywin32+Pillow 필요, Windows 전용)

src/batch/runner.py    # 조건 엑셀 여러 행 순회 실행, 조건별 성공/실패 추적, 실패건 재처리

src/excel_io/          # 엑셀 입출력 (SAP 의존성 없음 — GCP 포함 어디서나 완전히 검증됨)
  reader.py, writer.py, edit_rules.py, schema.py

src/lookup/             # 담당자 매칭 (3가지 방식 중 2가지 구현됨, SAP 의존성 없음)
  hr_matcher.py          - ① 인사데이터 매칭 (사번→이름→아이디)
  reference_matcher.py   - ② 별도 참조 엑셀 업로드 매칭 (SAP 미접촉)
                          - ③ SAP 보조조회는 미구현 (아래 "할 일" 참고)

src/validation/         # SAP 의존성 없음
  checks.py    - 건수·금액 검증, 중복체크, 정상/예외 분류
  sampling.py  - 표본 수 산정 (모집단 건수 기준 표 하나, determine_sample_size/select_sample)

src/web/                # SAP 의존성 없음(단, "SAP 세션 선택" 기능만 실제 SAP 필요)
  backend.py        - FastAPI: 통제 등록/조회, SAP 세션 목록, 배치 실행, 재처리, 다운로드
  static/index.html - 통제등록 / 담당자매칭 / 실행 3탭 (실제 화면, 프로토타입 아님)
  static/app.js
  static/style.css
```

## 지금까지 확정된 핵심 설계 결정 (재논의 없이 그대로 따를 것)

1. **SAP 자동화 방식**: SAP GUI Scripting(pywin32) 확정. RFC/HANA 직접접속은 보안정책상
   불가. OData API·iPaaS·ABAP 등 대안은 검토했으나 채택 안 함 — GUI Scripting이 이미
   사내에 유사 사례(로봇 자동화)가 있어 실현 가능성 확인됨. **이 방식의 근본적 한계로,
   서버가 클라우드(GCP 포함)에 있으면 절대 동작하지 않는다 — 반드시 SAP Logon이 로그인된
   PC 안에서 서버가 돌아야 한다.**
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
7. **실행 스크립트는 `run.sh` 하나로 단순화**: 사용자가 비개발자라 `pip install`,
   `PYTHONPATH` 설정, `uvicorn` 명령을 각각 이해할 필요가 없게, 이 세 단계를 스크립트
   하나로 묶었다. 대신 개발자를 위해 각 단계의 원래 명령은 이 문서 "실행 방법" 절에도
   그대로 남겨둔다.

## 여기서부터 할 일 (Gemini CLI / Vertex AI + 실제 SAP 환경에서)

### A. 반드시 실제 SAP 환경(Windows + SAP Logon PC)에서만 확인 가능 — GCP에서는 절대 불가능
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
- **GCP Vertex AI Workbench 등 원격 환경에서의 포트 접속 방법 확정** — `run.sh`는
  `--host 0.0.0.0`으로 띄우도록 이미 맞춰뒀지만, 실제로 어떤 방식(SSH 터널/웹 프리뷰
  프록시 등)으로 8000번 포트에 접속할지는 사내 GCP 네트워크 정책에 따라 다르므로, 여기서는
  Windows/로컬 기준의 `http://localhost:8000` 접속만 검증됐다. Gemini CLI가 실제 GCP
  환경에서 이 부분을 확인하고 사용자에게 정확한 접속 URL을 안내해야 한다.

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
- `run.sh`는 이 원격 샌드박스(리눅스, SAP 없음)에서 직접 실행해 `Uvicorn running on
  http://0.0.0.0:8000` 기동 및 `curl http://localhost:8000/`(200), `curl
  http://localhost:8000/static/style.css`(200) 응답을 확인했다. SAP 세션 목록/조회 관련
  기능은 이 환경에 SAP가 없어 확인 대상에서 제외했다(Windows+SAP 환경에서 별도 확인 필요).

## 실행 방법 (요약 — 자세한 단계별 설명은 위 "시작하기" 절 참고)

### 방법 1 — 스크립트 하나로 (권장, 비개발자용)

```bash
bash run.sh
```

### 방법 2 — 수동 (개발자가 각 단계를 직접 제어하고 싶을 때)

```bash
pip install -r requirements.txt
# Windows + SAP GUI(SAP Logon) 설치·로그인 상태에서:
PYTHONPATH=src python -m uvicorn web.backend:app --port 8000
# 브라우저에서 http://localhost:8000 접속
```

어느 방법이든 SAP GUI Scripting 관련 실제 동작(`session_picker`, `transaction_runner`,
`capture`)은 Windows + SAP Logon 환경에서만 동작한다. SAP가 없는 환경(GCP 포함)에서는
`__mock__` 세션으로 나머지 파이프라인(편집규칙→담당자매칭→검증→표본추출→결과엑셀)만
검증할 수 있다.
