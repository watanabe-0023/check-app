"""月次集計ユーティリティ。"""
import pandas as pd


def sum_ledger_by_month_category(df: pd.DataFrame, month: str, category_col: str, category_val: str | list, amount_col: str = "税別") -> float:
    """台帳DataFrameを月・カテゴリでフィルタしてSUMする。"""
    if df is None or df.empty:
        return 0.0
    mask = df["完了月"] == month
    if isinstance(category_val, list):
        mask &= df[category_col].isin(category_val)
    else:
        mask &= df[category_col] == category_val
    return float(df.loc[mask, amount_col].sum())


def sum_ledger_by_month(df: pd.DataFrame, month: str, amount_col: str = "税別") -> float:
    """台帳DataFrameを月でフィルタしてSUMする。"""
    if df is None or df.empty:
        return 0.0
    mask = df["完了月"] == month
    return float(df.loc[mask, amount_col].sum())


def sum_negative_by_month(df: pd.DataFrame, month: str, amount_col: str = "税別") -> float:
    """返金（金額<0）の合計を返す。"""
    if df is None or df.empty:
        return 0.0
    mask = (df["完了月"] == month) & (df[amount_col] < 0)
    return float(df.loc[mask, amount_col].sum())


def get_freee_account(freee_dict: dict, account: str, month: str) -> float:
    """freee辞書から特定科目・月の金額を取得する。"""
    return float(freee_dict.get(account, {}).get(month, 0.0))


def sum_freee_accounts(freee_dict: dict, accounts: list[str], month: str) -> float:
    return sum(get_freee_account(freee_dict, a, month) for a in accounts)
