"""勘定科目 → 管理表カテゴリ変換。"""
from .loaders import load_account_map


_MAP: dict | None = None


def get_map() -> dict:
    global _MAP
    if _MAP is None:
        _MAP = load_account_map()
    return _MAP


def account_to_category(account: str) -> str | None:
    m = get_map()
    entry = m.get(account)
    if entry is None:
        return None
    return entry.get("category")


def filter_by_category(freee_dict: dict, category: str, month: str) -> float:
    """freee PL/BS の辞書から特定カテゴリの金額を集計する。"""
    total = 0.0
    for account, months in freee_dict.items():
        cat = account_to_category(account)
        if cat == category:
            total += months.get(month, 0.0)
    return total
