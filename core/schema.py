from dataclasses import dataclass, field
from typing import Literal, Optional

Status = Literal["pass", "warn", "fail", "info"]


@dataclass
class RuleResult:
    rule_id: str
    label: str
    a_value: Optional[float]
    b_value: Optional[float]
    diff: Optional[float]
    diff_rate: Optional[float]
    status: Status
    note: str
    a_source: str = ""
    b_source: str = ""


@dataclass
class CheckContext:
    month: str  # "YYYY-MM"
    mgmt: dict = field(default_factory=dict)       # 収益管理表データ
    ledger_sales: object = None                     # 売上管理台帳 DataFrame
    ledger_exp: object = None                       # 支出管理表 DataFrame
    ledger_keihi: object = None                     # 経費台帳 DataFrame
    ledger_biz: object = None                       # 事業別支出管理 DataFrame
    freee_bs: dict = field(default_factory=dict)    # freee BS {科目: {月: 金額}}
    freee_pl: dict = field(default_factory=dict)    # freee PL {科目: {月: 金額}}
    thresholds: dict = field(default_factory=dict)
