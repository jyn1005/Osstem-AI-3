"""
GR/IR 매칭 엔진 단위 테스트

실행:
    pytest tests/test_gr_ir_matcher.py -v
"""

import pytest
import pandas as pd
from datetime import date, timedelta

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from src.modules.analytics.gr_ir_matcher import (
    match_gr_ir,
    add_aging,
    _classify,
    CLS_MATCHED,
    CLS_GOODS_RECV,
    CLS_ACCRUED,
    CLS_DIFF_AMT,
    CLS_DIFF_QTY,
    AGING_LABELS,
)


# ════════════════════════════════════════════════════════
#  픽스처: 샘플 GR / IR DataFrame
# ════════════════════════════════════════════════════════

def make_gr(rows: list) -> pd.DataFrame:
    """테스트용 GR DataFrame 생성."""
    return pd.DataFrame(rows, columns=[
        "po_no", "po_date", "mat_code", "gr_qty", "gr_amount", "currency"
    ])


def make_ir(rows: list) -> pd.DataFrame:
    """테스트용 IR DataFrame 생성."""
    return pd.DataFrame(rows, columns=[
        "po_no", "po_date", "mat_code", "ir_qty", "ir_amount", "currency"
    ])


TODAY = date.today()
PO_DATE = pd.Timestamp(TODAY - timedelta(days=10))


# ════════════════════════════════════════════════════════
#  match_gr_ir 테스트
# ════════════════════════════════════════════════════════

class TestMatchGrIr:

    def test_완전매칭(self):
        """GR금액 == IR금액 → 완전매칭"""
        gr = make_gr([("PO001", PO_DATE, "MAT001", 10.0, 100_000.0, "KRW")])
        ir = make_ir([("PO001", PO_DATE, "MAT001", 10.0, 100_000.0, "KRW")])
        result = match_gr_ir(gr, ir)
        assert result.loc[0, "분류"] == CLS_MATCHED

    def test_완전매칭_허용오차_이내(self):
        """잔액이 허용오차(1원) 이내 → 완전매칭"""
        gr = make_gr([("PO001", PO_DATE, "MAT001", 10.0, 100_000.0, "KRW")])
        ir = make_ir([("PO001", PO_DATE, "MAT001", 10.0, 100_000.5, "KRW")])
        result = match_gr_ir(gr, ir, tolerance=1.0)
        assert result.loc[0, "분류"] == CLS_MATCHED

    def test_미착품_GR만_있음(self):
        """GR은 있고 IR이 없는 경우 → 미착품"""
        gr = make_gr([("PO002", PO_DATE, "MAT002", 5.0, 50_000.0, "KRW")])
        ir = make_ir([])
        result = match_gr_ir(gr, ir)
        assert result.loc[0, "분류"] == CLS_GOODS_RECV

    def test_미확정채무_IR만_있음(self):
        """IR은 있고 GR이 없는 경우 → 미확정채무"""
        gr = make_gr([])
        ir = make_ir([("PO003", PO_DATE, "MAT003", 3.0, 30_000.0, "KRW")])
        result = match_gr_ir(gr, ir)
        assert result.loc[0, "분류"] == CLS_ACCRUED

    def test_예외_금액차이(self):
        """GR·IR 모두 있지만 금액이 다른 경우 → 예외_금액차이"""
        gr = make_gr([("PO004", PO_DATE, "MAT004", 10.0, 100_000.0, "KRW")])
        ir = make_ir([("PO004", PO_DATE, "MAT004", 10.0,  80_000.0, "KRW")])
        result = match_gr_ir(gr, ir)
        assert result.loc[0, "분류"] == CLS_DIFF_AMT

    def test_예외_수량차이(self):
        """GR·IR 수량이 다른 경우 → 예외_수량차이"""
        gr = make_gr([("PO005", PO_DATE, "MAT005", 10.0, 100_000.0, "KRW")])
        ir = make_ir([("PO005", PO_DATE, "MAT005",  7.0,  70_000.0, "KRW")])
        result = match_gr_ir(gr, ir)
        assert result.loc[0, "분류"] == CLS_DIFF_QTY

    def test_여러_PO_동시_처리(self):
        """여러 PO가 섞인 경우 각각 올바르게 분류"""
        gr = make_gr([
            ("PO001", PO_DATE, "MAT001", 10.0, 100_000.0, "KRW"),
            ("PO002", PO_DATE, "MAT002",  5.0,  50_000.0, "KRW"),
        ])
        ir = make_ir([
            ("PO001", PO_DATE, "MAT001", 10.0, 100_000.0, "KRW"),
            # PO002는 IR 없음 → 미착품
        ])
        result = match_gr_ir(gr, ir).set_index("PO번호")
        assert result.loc["PO001", "분류"] == CLS_MATCHED
        assert result.loc["PO002", "분류"] == CLS_GOODS_RECV

    def test_동일_PO_복수행_집계(self):
        """같은 PO 번호가 여러 행일 때 합산 후 매칭"""
        gr = make_gr([
            ("PO006", PO_DATE, "MAT006", 3.0, 30_000.0, "KRW"),
            ("PO006", PO_DATE, "MAT006", 7.0, 70_000.0, "KRW"),
        ])
        ir = make_ir([
            ("PO006", PO_DATE, "MAT006", 10.0, 100_000.0, "KRW"),
        ])
        result = match_gr_ir(gr, ir)
        assert len(result) == 1
        assert result.loc[0, "GR금액"] == 100_000.0
        assert result.loc[0, "분류"] == CLS_MATCHED

    def test_잔액_계산(self):
        """잔액 = GR금액 - IR금액"""
        gr = make_gr([("PO007", PO_DATE, "MAT007", 10.0, 100_000.0, "KRW")])
        ir = make_ir([("PO007", PO_DATE, "MAT007", 10.0,  60_000.0, "KRW")])
        result = match_gr_ir(gr, ir)
        assert result.loc[0, "잔액"] == 40_000.0


# ════════════════════════════════════════════════════════
#  add_aging 테스트
# ════════════════════════════════════════════════════════

class TestAddAging:

    def _base_df(self, days_ago: int) -> pd.DataFrame:
        po_date = pd.Timestamp(TODAY - timedelta(days=days_ago))
        gr = make_gr([("PO001", po_date, "MAT001", 1.0, 1000.0, "KRW")])
        ir = make_ir([])
        return match_gr_ir(gr, ir)

    def test_30일이하(self):
        df = add_aging(self._base_df(15), base_date=TODAY)
        assert df.loc[0, "Aging구간"] == "30일이하"

    def test_31_60일(self):
        df = add_aging(self._base_df(45), base_date=TODAY)
        assert df.loc[0, "Aging구간"] == "31~60일"

    def test_61_90일(self):
        df = add_aging(self._base_df(75), base_date=TODAY)
        assert df.loc[0, "Aging구간"] == "61~90일"

    def test_90일초과(self):
        df = add_aging(self._base_df(120), base_date=TODAY)
        assert df.loc[0, "Aging구간"] == "90일초과"

    def test_경과일수_계산(self):
        df = add_aging(self._base_df(30), base_date=TODAY)
        assert df.loc[0, "경과일수"] == 30

    def test_날짜없음(self):
        """PO일자가 NaT인 경우 '날짜없음'"""
        gr = make_gr([("PO001", pd.NaT, "MAT001", 1.0, 1000.0, "KRW")])
        ir = make_ir([])
        matched = match_gr_ir(gr, ir)
        df = add_aging(matched, base_date=TODAY)
        assert df.loc[0, "Aging구간"] == "날짜없음"


# ════════════════════════════════════════════════════════
#  _classify 함수 단위 테스트
# ════════════════════════════════════════════════════════

class TestClassify:

    def _row(self, gr_amt, ir_amt, gr_qty=10, ir_qty=10):
        return pd.Series({
            "gr_amount": gr_amt,
            "ir_amount": ir_amt,
            "gr_qty":    gr_qty,
            "ir_qty":    ir_qty,
        })

    def test_완전매칭_잔액0(self):
        assert _classify(self._row(100, 100), 1.0) == CLS_MATCHED

    def test_완전매칭_허용오차(self):
        assert _classify(self._row(100, 100.5), 1.0) == CLS_MATCHED

    def test_허용오차_초과_예외(self):
        assert _classify(self._row(100, 98), 1.0) != CLS_MATCHED

    def test_미착품(self):
        assert _classify(self._row(100, 0), 1.0) == CLS_GOODS_RECV

    def test_미확정채무(self):
        assert _classify(self._row(0, 100), 1.0) == CLS_ACCRUED

    def test_수량차이_우선(self):
        assert _classify(self._row(100, 70, gr_qty=10, ir_qty=7), 1.0) == CLS_DIFF_QTY

    def test_금액차이(self):
        assert _classify(self._row(100, 80, gr_qty=10, ir_qty=10), 1.0) == CLS_DIFF_AMT
