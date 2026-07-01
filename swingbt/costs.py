"""
各市場成本模型 — 讓跨市場比較「誠實」。

不同市場的交易成本差很多,直接影響策略能不能活:
  - 加密貨幣:taker 手續費約 0.055%/邊,可做空、可槓桿,一年 365 交易日。
  - 美股:券商幾乎零手續費,但有滑點;無證交稅,一年約 252 交易日。
  - 台股:券商 0.1425%/邊 + 賣出證交稅 0.3%(當沖減半,這裡以波段的 0.3% 計),
          一年約 252 交易日。這條若不建模,台股回測會過度樂觀。
  - 指數:本身不可直接買賣,實務用 ETF 代理(美股 SPY/QQQ、台股 0050),
          成本套對應的股票市場。

commission_pct / tax_sell_pct / slippage_pct 皆為「小數比例」(0.001 = 0.1%)。
"""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class CostModel:
    name: str
    commission_pct: float      # 每邊手續費(買、賣各收一次)
    tax_sell_pct: float        # 賣出時的稅(如台股證交稅);買進不收
    slippage_pct: float        # 每邊估計滑點
    allow_short: bool          # 該市場散戶是否方便做空
    max_leverage: float        # 倉位規模上限(1.0 = 不用槓桿)
    periods_per_year: int      # 年化用(交易日/年);crypto=365

    def entry_cost_pct(self) -> float:
        """進場每 1 元名目金額的成本比例。"""
        return self.commission_pct + self.slippage_pct

    def exit_cost_pct(self) -> float:
        """出場每 1 元名目金額的成本比例(含賣出稅)。"""
        return self.commission_pct + self.slippage_pct + self.tax_sell_pct


# ── 預設市場成本 ────────────────────────────────────────────────
MARKETS = {
    "crypto": CostModel(
        name="crypto", commission_pct=0.00055, tax_sell_pct=0.0,
        slippage_pct=0.0005, allow_short=True, max_leverage=5.0,
        periods_per_year=365,
    ),
    "us_stock": CostModel(
        name="us_stock", commission_pct=0.0001, tax_sell_pct=0.0,
        slippage_pct=0.0005, allow_short=True, max_leverage=1.0,
        periods_per_year=252,
    ),
    "tw_stock": CostModel(
        name="tw_stock", commission_pct=0.001425, tax_sell_pct=0.003,
        slippage_pct=0.0008, allow_short=False, max_leverage=1.0,
        periods_per_year=252,
    ),
    # 指數:用 ETF 代理。美股指數走 us_stock 成本,台股指數(0050/加權)走 tw_stock。
    "index_us": CostModel(
        name="index_us", commission_pct=0.0001, tax_sell_pct=0.0,
        slippage_pct=0.0004, allow_short=True, max_leverage=1.0,
        periods_per_year=252,
    ),
    "index_tw": CostModel(
        name="index_tw", commission_pct=0.001425, tax_sell_pct=0.003,
        slippage_pct=0.0006, allow_short=False, max_leverage=1.0,
        periods_per_year=252,
    ),
}


def get_cost_model(market: str) -> CostModel:
    if market not in MARKETS:
        raise KeyError(
            f"未知市場 '{market}';可用:{', '.join(MARKETS)}"
        )
    return MARKETS[market]
