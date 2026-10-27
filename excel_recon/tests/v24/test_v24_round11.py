# 11차 검토(독립 리뷰어 10차) 지적 고정. 실행: TOOL_DIR=<폴더> python3 test_v24_round11.py
import os, sys, unittest, datetime, random
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from test_v24_review import m, wx, quiet
from test_v24_round10 import Round10
from test_v24_round8 import RAWG
D = datetime.date
ACC = [("매출", "제품매출", 1000), ("매출", "상품매출", 400), ("판관비", "급여", 300), ("판관비", "복리후생비", 80), ("판관비", "임차료", 120)]
BS = [("유동자산", "현금", 100), ("유동자산", "예금", 200), ("비유동자산", "토지", 500), ("유동부채", "매입채무", 150), ("비유동부채", "차입금", 250), ("자본", "자본금", 300), ("자본", "이익잉여금", 100)]


class Round11(Round10):
    def test_r11_b1_cross_section_dup_not_silent(self):
        r, b1, u, mm = self.agg(RAWG, [["계정과목", "금액"], ["현금", 100], ["예금", 200], ["소계", 300], ["현금", 100], ["토지", 500], ["소계", 600], ["합계", 800]], "r11b1")
        self.assertNotEqual(b1, "전부 일치", (u, mm))
        r, b1, u, mm = self.agg(RAWG, [["계정과목", "금액"], ["현금", 100], ["예금", 200], ["소계", 300], ["현금", 100], ["토지", 500], ["소계", 600], ["합계", 900]], "r11b1c")
        self.assertEqual(u.get("보고서!B5"), "확인 필요", (u, mm))      # 소계·합계 산수는 맞아도 다른 구간의 같은 계정은 확인 필요
        r, b1, u, mm = self.agg(RAWG, [["구분", "현금", "예금", "소계", "현금", "토지", "소계", "합계"], ["금액", 100, 200, 300, 100, 500, 600, 800]], "r11b1h")
        self.assertNotEqual(b1, "전부 일치", (u, mm))

    def test_r11_f_stacked_tables_with_totals(self):
        raw = [["대분류", "계정과목", "금액"]] + [list(x) for x in ACC]
        self.ok(raw, [["계정과목", "금액"], ["제품매출", 1000], ["급여", 300], ["합계", 1300], [None, None], ["제품매출", 1000], ["상품매출", 400], ["급여", 300], ["복리후생비", 80], ["임차료", 120], ["합계", 1900]], "r11f1")
        self.ok(raw, [["계정과목", "금액"], ["제품매출", 1000], ["상품매출", 400], ["매출 합계", 1400], ["제품매출", 1000], ["상품매출", 400], ["급여", 300], ["복리후생비", 80], ["임차료", 120], ["총계", 1900]], "r11f2")

    def test_r11_c_balance_sheet_forms(self):
        raw = [["대분류", "계정과목", "금액"]] + [list(x) for x in BS]
        self.ok(raw, [["계정과목", "금액"], ["유동자산", None], ["현금", 100], ["예금", 200], ["소계", 300], ["비유동자산", None], ["토지", 500], ["소계", 500], ["자산총계", 800],
                      ["부채", None], ["매입채무", 150], ["차입금", 250], ["부채총계", 400], ["자본", None], ["자본금", 300], ["이익잉여금", 100], ["자본총계", 400]], "r11c2")

    def test_r11_c5_column_group_totals(self):
        raw = [["부서", "계정과목", "금액"]] + [[dp, ac, v] for dp, k in (("영업1팀", 1), ("영업2팀", 2), ("관리팀", 3), ("인사팀", 4)) for ac, v in (("급여", 100 * k), ("복리", 10 * k))]
        self.ok(raw, [["계정과목", "영업1팀", "영업2팀", "영업 합계", "관리팀", "인사팀", "관리 합계"], ["급여", 100, 200, 300, 300, 400, 700], ["복리", 10, 20, 30, 30, 40, 70]], "r11c5")
        r, b1, u, mm = self.agg(raw, [["계정과목", "영업1팀", "영업2팀", "영업 합계", "관리팀", "인사팀", "관리 합계"], ["급여", 100, 200, 300, 300, 400, 1000], ["복리", 10, 20, 30, 30, 40, 70]], "r11c5b")
        self.assertIn("보고서!G2", mm)

    def test_r11_e_small_department_copy_location(self):
        H = ["일자", "부서", "계정명", "금액", "적요"]
        small = [5, 20, 35, 50, 65, 80]
        rnd = random.Random(3)
        R = [(D(2026, 1, 1) + datetime.timedelta(days=i), "관리부" if i in small else "영업부", f"계정{rnd.randint(0, 4)}", 10 * rnd.randint(1, 10 ** 5), f"적요{i}") for i in range(100)]
        self.raw = wx(os.path.join(self.d, "raws.xlsx"), [("원본", [H] + R)])
        big = [i for i in range(100) if i not in small]
        wrong = []
        for k in (5, 80, 35):
            for ys in (sorted(big + [k]), big + [k]):
                rep, res, c = self.rows_of([("영업부", [H] + [R[i] for i in ys]), ("관리부", [H] + [R[i] for i in small])])
                if [x["시트"] for x in res if "중복" in x["이유"]] != ["영업부"]:
                    wrong.append(k)
        self.assertEqual(wrong, [])


if __name__ == "__main__":
    unittest.main(defaultTest="Round11")
