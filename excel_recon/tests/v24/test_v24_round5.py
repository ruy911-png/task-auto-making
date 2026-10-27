# 5차 검토(독립 리뷰어 4차) 지적 고정. 실행: TOOL_DIR=<폴더> python3 test_v24_round5.py
import os, sys, unittest
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_v24_review import m, wx, quiet, H
from test_v24_round4 import Round4, RAW
RAWB = RAW + [[a, b, c, "B", v * 2] for a, b, c, _, v in RAW[1:]]
PA = [i for i in range(100) if i % 5 < 3]; PB = [i for i in range(100) if i % 5 >= 3]


class Round5(Round4):
    def test_r5_a1_column_top_subtotal_blank_is_missing(self):
        r, b1, u, mm = self.agg(RAW, [["구분", "유동자산", "현금", "예금", "상품", "비유동자산", "토지"], ["A", 350, 100, 200, 50, None, 500]], "r5a1")
        self.assertEqual(u.get("보고서!F2"), "누락", u)
        r, b1, u, mm = self.agg(RAWB, [["구분", "유동자산", "현금", "예금", "상품", "비유동자산", "토지"], ["A", 350, 100, 200, 50, None, 500], ["B", 700, 200, 400, 100, None, 1000]], "r5a1c")
        self.assertEqual((u.get("보고서!F2"), u.get("보고서!F3")), ("누락", "누락"), u)

    def test_r5_a2_second_block_blank_is_missing(self):
        r, b1, u, mm = self.agg(RAW, [["계정과목", "금액"], ["비유동자산", 500], ["토지", 500], [None, None], ["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["상품", 50]], "r5a2")
        self.assertEqual(u.get("보고서!B6"), "누락", u)

    def test_r5_a3_zero_accounts_omitted_ok(self):
        rep = [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["상품", 50], ["비유동자산", None], ["토지", 500], ["합계", 850]]
        self.ok(RAW[:3] + [["유동자산", "당좌자산", "단기대여금", "A", 0]] + RAW[3:], rep, "r5a3")
        self.ok(RAW[:3] + [["유동자산", "당좌자산", "가지급금", "A", 30], ["유동자산", "당좌자산", "가지급금", "A", -30]] + RAW[3:], rep, "r5a3b")

    def test_r5_f_formula_heading_without_cache(self):
        r, b1, u, mm = self.agg(RAW, [["계정과목", "금액"], ["유동자산", "=SUM(B3:B5)"], ["현금", 100], ["예금", 200], ["상품", 50], ["비유동자산", "=B7"], ["토지", 500]], "r5f1")
        self.assertEqual((u.get("보고서!B2"), u.get("보고서!B6")), ("계산값 없음", "계산값 없음"), u)
        r, b1, u, mm = self.agg(RAW, [["구분", "유동자산", "현금", "예금", "상품", "비유동자산", "토지"], ["A", "=SUM(C2:E2)", 100, 200, 50, "=G2", 500]], "r5f2")
        self.assertEqual((u.get("보고서!B2"), u.get("보고서!F2")), ("계산값 없음", "계산값 없음"), u)

    def test_r5_b9_two_partitions(self):
        sheets = lambda a, b: [("1월", range(50)), ("2월", range(50, 100)), ("A법인", a), ("B법인", b)]
        rep, res, c = self.groups(100, sheets(PA, PB))
        self.assertEqual(self.final(rep), "전부 일치", c)
        rep, res, c = self.groups(100, sheets([i for i in PA if i != 7], PB))
        miss = [x for x in res if x["판정"] == "보고서에 없음"]
        self.assertEqual([x["원본 행"] for x in miss], [9], c); self.assertFalse([x for x in res if "중복" in x["이유"]], c)

    def test_r5_b2_split_missing_with_bigger_extract(self):
        ex = [i for i in range(99) if i % 5 in (0, 2)][:40]
        rep, res, c = self.groups(99, [("1월", range(33)), ("2월", [i for i in range(33, 66) if i != 35]), ("3월", range(66, 99)), ("주요거래", ex)])
        self.assertEqual([x["원본 행"] for x in res if x["판정"] == "보고서에 없음"], [37], c)
        self.assertFalse([x for x in res if "중복" in x["이유"]], c)

    def test_r5_b6_two_views_each_split_with_dup(self):
        rep, res, c = self.groups(100, [("원화", range(100)), ("달러", range(100)), ("원화상", range(50)), ("원화하", range(50, 100)),
                                        ("달러상", range(50)), ("달러하", [10] + list(range(50, 100)))])
        d = [x for x in res if "중복" in x["이유"]]
        self.assertEqual([(x["시트"], x["원본 행"]) for x in d], [("달러하", 12)], c)

    def test_r5_b12_overlapping_split_plus_copy(self):
        rep, res, c = self.groups(20, [("A", range(15)), ("B", range(5, 20)), ("A사본", range(15))])
        d = [x for x in res if "중복" in x["이유"]]
        self.assertEqual(len(d), 10, c)

    def test_r5_b4_filter_view_is_note(self):
        rep, res, c = self.groups(99, [("전체", range(99)), ("수익거래", [i for i in range(99) if i % 5 < 3])])
        self.assertEqual((c["보고서에 없음"], c["확인 필요"]), (0, 1), c)

    def test_r5_m_memo_rows_block(self):
        from test_v24_review import ROWS
        for nm, extra in (("m1", [["검토 메모"], ["3월분", "삼성전자 매수 누락", "추가 확인 필요"]]), ("m2", [["확인사항", "계정3 2건 누락", "재작성 요청"]]),
                          ("m3", [["수기 조정"], ["2026-03"], ["계정1", "추가 반영", "5만원", "김대리"]])):
            rep = wx(os.path.join(self.d, f"p{nm}.xlsx"), [("본표", [H] + ROWS), ("메모", extra)])
            self.assertNotEqual(self.final(rep), "전부 일치", nm)


if __name__ == "__main__":
    unittest.main(defaultTest="Round5")
