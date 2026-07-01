"""
事件驅動回測引擎。

執行語義(與 TradingView `process_orders_on_close=true` 及你的 validate_all.py 一致):
  對每根 bar i:
    A. 若持倉:先用「上一根定案的停損」檢查本根是否觸損(用盤中 high/low)。
       - 跳空穿價 → 以開盤價成交(對你不利的一方);否則以停損價成交。
    B. 未觸損 → 呼叫策略更新「下一根生效」的移動停損。
    C. 若空手且已過暖身 → 問策略要不要進場;要 → 以 close[i] 進場。
    D. 記錄權益(現金已實現 + 未實現)。

無未來函數:進場只用 <= i 的資料;觸損檢查用的停損是上一根就定好的。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

import numpy as np

from .costs import CostModel
from .data import OHLCV
from .strategies.base import Strategy, Position


@dataclass
class BacktestResult:
    strategy: str
    ticker: str
    market: str
    dates: list
    equity: np.ndarray          # 每根的權益(mark-to-market)
    in_position: np.ndarray     # bool,每根是否持倉(算曝險率)
    trades: list                # list[dict]
    init_capital: float
    cost: CostModel
    data: OHLCV = field(repr=False, default=None)


def run(data: OHLCV, strategy: Strategy, cost: CostModel,
        init_capital: float = 1_000_000.0, risk_pct: float = 5.0,
        min_stop_frac: float = 0.005) -> BacktestResult:
    N = len(data)
    o, h, l, c = data.o, data.h, data.l, data.c

    strategy.prepare(data)
    # 讓 base.on_entry 能存取高低(策略內部也各自持有指標)
    strategy._h, strategy._l, strategy._c = h, l, c  # type: ignore[attr-defined]

    warmup = max(strategy.warmup, 1)

    cash = float(init_capital)      # 已實現權益
    pos: Optional[Position] = None
    equity = np.empty(N)
    in_pos = np.zeros(N, dtype=bool)
    trades: list = []

    def open_notional(p: Position) -> float:
        return p.qty * p.entry_price

    max_hold = strategy.max_hold

    for i in range(N):
        # ── A/B:持倉管理 ──
        if pos is not None:
            exit_price = None
            reason = ""
            # 1) 停損(悲觀假設:同一根若停損與目標都可能觸及,先算停損)
            if pos.direction > 0:
                if l[i] <= pos.stop:
                    exit_price = o[i] if o[i] < pos.stop else pos.stop
                    reason = "stop_long"
            else:
                if h[i] >= pos.stop:
                    exit_price = o[i] if o[i] > pos.stop else pos.stop
                    reason = "stop_short"
            # 2) 目標價(均值回歸)
            if exit_price is None and pos.target is not None:
                if pos.direction > 0 and h[i] >= pos.target:
                    exit_price = o[i] if o[i] > pos.target else pos.target
                    reason = "target_long"
                elif pos.direction < 0 and l[i] <= pos.target:
                    exit_price = o[i] if o[i] < pos.target else pos.target
                    reason = "target_short"
            # 3) 訊號式出場(如逆勢「回到均值就走」)→ 收盤平倉
            if exit_price is None and (i - pos.entry_i) >= 1 and strategy.exit_signal(i, pos):
                exit_price = c[i]
                reason = "signal"
            # 4) 時間停損(持太久 → 收盤平倉)
            if exit_price is None and max_hold is not None and (i - pos.entry_i) >= max_hold:
                exit_price = c[i]
                reason = "time"

            if exit_price is not None:
                gross = (exit_price - pos.entry_price) * pos.qty * pos.direction
                entry_cost = open_notional(pos) * cost.entry_cost_pct()
                exit_cost = pos.qty * exit_price * cost.exit_cost_pct()
                pnl = gross - entry_cost - exit_cost
                cash += pnl
                trades.append(_mk_trade(data, pos, i, exit_price, pnl, reason))
                pos = None
            else:
                # 未觸損 → 更新下一根生效的停損
                pos.stop = strategy.update_stop(i, pos)

        # ── C:進場(最後一根不開:沒有後續 bar 可管理,只會變成 bars=0 的幽靈單)──
        if pos is None and warmup <= i < N - 1:
            e = strategy.entry(i)
            if e is not None:
                if e.direction < 0 and not cost.allow_short:
                    e = None  # 該市場不方便做空 → 跳過空單
            if e is not None:
                entry_price = c[i]
                stop_dist = abs(entry_price - e.stop)
                stop_frac = max(stop_dist / entry_price, min_stop_frac)
                stop_dist = stop_frac * entry_price
                risk_amount = cash * (risk_pct / 100.0) * e.risk_mult
                qty_risk = risk_amount / stop_dist
                qty_cap = cost.max_leverage * cash / entry_price
                qty = min(qty_risk, qty_cap)
                if qty > 0:
                    # 停損距離 = 用來 sizing 的距離(已套 min_stop_frac 下限)。
                    # 把實際停損放在「和 sizing 一致」的距離上,實現風險才會等於 risk_pct;
                    # 且此式在正常情況(原始停損比下限寬)會還原策略原本的停損價。
                    stop_price = (entry_price - stop_dist if e.direction > 0
                                  else entry_price + stop_dist)
                    pos = Position(direction=e.direction, entry_price=entry_price,
                                   qty=qty, entry_i=i, stop=stop_price)
                    strategy.on_entry(i, pos)

        # ── D:記錄權益 ──
        if pos is not None:
            unreal = (c[i] - pos.entry_price) * pos.qty * pos.direction
            entry_cost = open_notional(pos) * cost.entry_cost_pct()
            equity[i] = cash - entry_cost + unreal
            in_pos[i] = True
        else:
            equity[i] = cash

    # ── 收尾:仍在場 → 以最後收盤平倉 ──
    if pos is not None:
        i = N - 1
        exit_price = c[i]
        gross = (exit_price - pos.entry_price) * pos.qty * pos.direction
        entry_cost = open_notional(pos) * cost.entry_cost_pct()
        exit_cost = pos.qty * exit_price * cost.exit_cost_pct()
        pnl = gross - entry_cost - exit_cost
        cash += pnl
        trades.append(_mk_trade(data, pos, i, exit_price, pnl, "eod"))
        equity[i] = cash

    return BacktestResult(
        strategy=strategy.name, ticker=data.ticker, market=cost.name,
        dates=data.dates, equity=equity, in_position=in_pos, trades=trades,
        init_capital=init_capital, cost=cost, data=data,
    )


def _mk_trade(data: OHLCV, pos: Position, exit_i: int, exit_price: float,
              pnl: float, reason: str) -> dict:
    notional = pos.qty * pos.entry_price
    return {
        "dir": pos.direction,
        "entry_i": pos.entry_i,
        "exit_i": exit_i,
        "entry_date": data.dates[pos.entry_i],
        "exit_date": data.dates[exit_i],
        "entry": pos.entry_price,
        "exit": exit_price,
        "qty": pos.qty,
        "pnl": pnl,
        "ret_on_notional": pnl / notional if notional else 0.0,
        "bars": exit_i - pos.entry_i,
        "reason": reason,
    }


# 便利類別包裝
class Backtester:
    def __init__(self, cost: CostModel, init_capital: float = 1_000_000.0,
                 risk_pct: float = 5.0):
        self.cost = cost
        self.init_capital = init_capital
        self.risk_pct = risk_pct

    def run(self, data: OHLCV, strategy: Strategy) -> BacktestResult:
        return run(data, strategy, self.cost, self.init_capital, self.risk_pct)
