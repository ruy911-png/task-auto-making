# v24 자동 시험

합성 데이터만 씀(개인·회사 자료 없음). 이 폴더에서 실행하면 두 폴더 위 `excel_automation_v24.py`를 시험한다.

```
python3 test_v24_round11.py      # review ~ round11 누적 98건 (round11이 앞 회차를 상속)
python3 test_v24_independent.py  # 독립 에이전트 합격 시험 44건
TOOL=excel_automation_v23 TOOL_DIR=../.. python3 test_v24_round11.py   # 다른 판과 비교
```

| 파일 | 내용 |
|---|---|
| test_v24_review.py | PR #2 검토·이슈 #5 지적 (빈 칸 누락, 건너뛴 시트, 긴 번호, 시트 분할, 재검산 칸 수, cp949) |
| test_v24_round3~11.py | 독립 검토 2~10차 지적 (구분 제목, 위쪽 소계형, 같은 계정 두 번, 층 판정, 중복 위치 등) |
| test_v24_independent.py | 별도 에이전트가 먼저 쓴 합격 시험 |
