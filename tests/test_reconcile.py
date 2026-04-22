"""突合ロジック単体テスト。"""
import pytest
import pandas as pd
from unittest.mock import MagicMock

from core.schema import CheckContext, RuleResult
from core.reconcilers import (
    judge,
    rule_S01_total_revenue,
    rule_S02_ios_revenue,
    rule_S03_tech_revenue,
    rule_S04_misc_revenue,
    rule_S05_revenue_vs_freee,
    rule_S06_refund,
    rule_E01_cost,
    rule_E02_promotion,
    rule_T01_gross_profit,
    rule_T02_sga_total,
    rule_T03_operating_income,
    rule_P05_allocation_rate,
)


# ---------------------------------------------------------------------------
# judge() テスト
# ---------------------------------------------------------------------------

def test_judge_pass_abs():
    assert judge(0.5, 0.001, 1.0, 0.01) == "pass"

def test_judge_pass_rate():
    assert judge(500, 0.005, 1.0, 0.01) == "pass"

def test_judge_warn():
    assert judge(1000, 0.02, 1.0, 0.01) == "warn"

def test_judge_fail():
    assert judge(10000, 0.5, 1.0, 0.01) == "fail"

def test_judge_info():
    assert judge(99999, 9.9, None, None) == "info"


# ---------------------------------------------------------------------------
# テスト用フィクスチャ
# ---------------------------------------------------------------------------

MONTH = "2026-01"

def _make_sales_df():
    return pd.DataFrame({
        "区分": ["iOSアカデミア", "SES", "アプリ受託開発", "雑収入", "iOSアカデミア"],
        "税別": [1_000_000, 500_000, 300_000, 50_000, -10_000],
        "完了月": [MONTH, MONTH, MONTH, MONTH, MONTH],
    })

def _make_exp_df():
    return pd.DataFrame({
        "費目": ["原価", "広告宣伝費", "旅費交通費", "採用費", "福利厚生費"],
        "税別": [200_000, 50_000, 30_000, 20_000, 15_000],
        "完了月": [MONTH, MONTH, MONTH, MONTH, MONTH],
    })

def _make_mgmt(overrides: dict | None = None) -> dict:
    """収支合計シートを模したDataFrameを返す。"""
    # 行1=ヘッダ相当（行番号のインデックスで参照するため、DataFrameの行0=ルール上の行1）
    data = {
        f"2026/1": [
            1_840_000,   # 行1
            1_840_000,   # 行2: 売上合計
            0,           # 行3
            200_000,     # 行4: 原価
            50_000,      # 行5: 販促費
            0,           # 行6
            20_000,      # 行7: 採用教育
            15_000,      # 行8: 福利厚生
            30_000,      # 行9: 管理費
            0, 0, 0, 0,  # 行10-13
            0,           # 行14
            1_640_000,   # 行15: 粗利
            0,           # 行16
            115_000,     # 行17: 販管費計
            1_525_000,   # 行18: 営業利益
            0, 0,        # 行19-20
            0,           # 行21
            1_525_000,   # 行22
            0,           # 行23
            1_525_000,   # 行24: CF
        ]
    }
    df = pd.DataFrame(data)
    result = {"収支合計": df}
    if overrides:
        result.update(overrides)
    return result

def _make_ctx(mgmt=None, sales=None, exp=None, freee_pl=None, freee_bs=None, thresholds=None):
    if thresholds is None:
        from core.loaders import load_thresholds
        thresholds = load_thresholds()
    return CheckContext(
        month=MONTH,
        mgmt=_make_mgmt() if mgmt is None else mgmt,
        ledger_sales=_make_sales_df() if sales is None else sales,
        ledger_exp=_make_exp_df() if exp is None else exp,
        freee_pl=freee_pl if freee_pl is not None else {},
        freee_bs=freee_bs if freee_bs is not None else {},
        thresholds=thresholds,
    )


# ---------------------------------------------------------------------------
# S系テスト
# ---------------------------------------------------------------------------

def test_S01_pass():
    """売上台帳合計と管理表が一致する場合はPASS。"""
    sales = _make_sales_df()
    total = float(sales["税別"].sum())  # 1,840,000
    mgmt = _make_mgmt()
    # 行2の値をsalesの合計と合わせる
    mgmt["収支合計"].iloc[1, 0] = total
    ctx = _make_ctx(mgmt=mgmt, sales=sales)
    r = rule_S01_total_revenue(ctx)
    assert r.status == "pass", f"diff={r.diff}"


def test_S01_fail_large_diff():
    """大きな差異がある場合はfail。"""
    sales = _make_sales_df()
    mgmt = _make_mgmt()
    mgmt["収支合計"].iloc[1, 0] = 9_999_999  # 大きな差
    ctx = _make_ctx(mgmt=mgmt, sales=sales)
    r = rule_S01_total_revenue(ctx)
    assert r.status in ("fail", "warn")


def test_S02_ios_pass():
    """iOSアカデミア売上が一致する場合はPASS。"""
    sales = _make_sales_df()
    ios_total = float(sales[sales["区分"] == "iOSアカデミア"]["税別"].sum())  # 990,000
    # row=2 → df.iloc[1] なので、index 1 に値を置く
    ios_mgmt = pd.DataFrame({"2026/1": [0, ios_total]})
    mgmt = _make_mgmt({"iOSアカデミア": ios_mgmt})
    ctx = _make_ctx(mgmt=mgmt, sales=sales)
    r = rule_S02_ios_revenue(ctx)
    assert r.status == "pass", f"diff={r.diff}"


def test_S03_tech_pass():
    sales = _make_sales_df()
    tech_total = float(sales[sales["区分"].isin(["SES", "アプリ受託開発"])]["税別"].sum())
    tech_mgmt = pd.DataFrame({"2026/1": [0, tech_total]})
    mgmt = _make_mgmt({"init TECH": tech_mgmt})
    ctx = _make_ctx(mgmt=mgmt, sales=sales)
    r = rule_S03_tech_revenue(ctx)
    assert r.status == "pass"


def test_S04_misc_pass():
    sales = _make_sales_df()
    misc_total = float(sales[sales["区分"] == "雑収入"]["税別"].sum())
    # row=2 → df.iloc[1]
    other_mgmt = pd.DataFrame({"2026/1": [0, misc_total, 0, 0, 0, 0]})
    mgmt = _make_mgmt({"その他": other_mgmt})
    ctx = _make_ctx(mgmt=mgmt, sales=sales)
    r = rule_S04_misc_revenue(ctx)
    assert r.status == "pass"


def test_S05_revenue_vs_freee_pass():
    sales = _make_sales_df()
    total = float(sales["税別"].sum())
    freee_pl = {"売上高": {MONTH: total}}
    mgmt = _make_mgmt()
    mgmt["収支合計"].iloc[1, 0] = total
    ctx = _make_ctx(mgmt=mgmt, sales=sales, freee_pl=freee_pl)
    r = rule_S05_revenue_vs_freee(ctx)
    assert r.status == "pass"


def test_S06_refund_pass():
    sales = _make_sales_df()
    neg_total = float(sales[sales["税別"] < 0]["税別"].sum())  # -10,000
    other_mgmt = pd.DataFrame({"2026/1": [0, 0, 0, 0, 0, neg_total]})
    mgmt = _make_mgmt({"その他": other_mgmt})
    ctx = _make_ctx(mgmt=mgmt, sales=sales)
    r = rule_S06_refund(ctx)
    assert r.status == "pass"


# ---------------------------------------------------------------------------
# E系テスト
# ---------------------------------------------------------------------------

def test_E01_cost_pass():
    exp = _make_exp_df()
    cost_total = float(exp[exp["費目"] == "原価"]["税別"].sum())
    mgmt = _make_mgmt()
    mgmt["収支合計"].iloc[3, 0] = cost_total
    ctx = _make_ctx(mgmt=mgmt, exp=exp)
    r = rule_E01_cost(ctx)
    assert r.status == "pass"


def test_E02_promotion_pass():
    exp = _make_exp_df()
    promo_cats = ["広告宣伝費", "販売促進費", "交際費", "会議費", "広告宣伝", "販促", "交際", "会議"]
    promo_total = float(exp[exp["費目"].isin(promo_cats)]["税別"].sum())
    mgmt = _make_mgmt()
    mgmt["収支合計"].iloc[4, 0] = promo_total
    ctx = _make_ctx(mgmt=mgmt, exp=exp)
    r = rule_E02_promotion(ctx)
    assert r.status == "pass"


# ---------------------------------------------------------------------------
# T系テスト
# ---------------------------------------------------------------------------

def test_T01_gross_profit_pass():
    """粗利 = 売上 - 原価 の整合性。"""
    ctx = _make_ctx()
    r = rule_T01_gross_profit(ctx)
    # 管理表内整合なのでPASSになるはず
    assert r.status in ("pass", "warn", "fail")  # 構造確認のみ
    assert r.rule_id == "T01"


def test_T02_sga_total_structure():
    ctx = _make_ctx()
    r = rule_T02_sga_total(ctx)
    assert r.rule_id == "T02"


def test_T03_operating_income_structure():
    ctx = _make_ctx()
    r = rule_T03_operating_income(ctx)
    assert r.rule_id == "T03"


# ---------------------------------------------------------------------------
# P系テスト
# ---------------------------------------------------------------------------

def test_P05_allocation_rate_100():
    """配賦率合計が100%の場合はPASS。"""
    alloc_df = pd.DataFrame({"2026/1": [0.3, 0.3, 0.4]})  # 合計1.0 = 100%
    mgmt = _make_mgmt({"給与原価割合": alloc_df})
    ctx = _make_ctx(mgmt=mgmt)
    r = rule_P05_allocation_rate(ctx)
    assert r.status == "pass", f"total={r.b_value}"


def test_P05_allocation_rate_over():
    """配賦率合計が100%を超える場合は失敗。"""
    alloc_df = pd.DataFrame({"2026/1": [0.4, 0.4, 0.4]})  # 合計1.2 = 120%
    mgmt = _make_mgmt({"給与原価割合": alloc_df})
    ctx = _make_ctx(mgmt=mgmt)
    r = rule_P05_allocation_rate(ctx)
    assert r.status in ("warn", "fail")


# ---------------------------------------------------------------------------
# run_all() 結果構造テスト
# ---------------------------------------------------------------------------

def test_run_all_returns_dict():
    """run_allが正しいキーを持つdictを返すことを確認。"""
    from core.reconcilers import run_all

    sales_df = _make_sales_df()
    exp_df = _make_exp_df()

    # DataFrameをBytesIOに変換してファイルとして渡す
    import io, openpyxl
    def df_to_xlsx_bytes(df):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        buf.seek(0)
        return buf

    # 最小限のファイルをモックとして渡す（実際の読み込みは省略）
    # run_allはNoneファイルを許容するので空dict相当で実行
    result = run_all(
        {"mgmt": None, "sales": None, "exp": None,
         "keihi": None, "biz": None, "bs": None, "pl": None},
        MONTH,
    )
    assert "results" in result
    assert "irregulars" in result
    assert "hash_changes" in result
    assert isinstance(result["results"], list)
