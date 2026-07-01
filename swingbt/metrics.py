"""
績效指標 — 從權益曲線 + 交易紀錄算出完整成績單,並附「買進持有」基準。

為什麼一定要附 Buy & Hold:
    在長多市場(如 0050 二十年)裡,策略「賺 +700%」聽起來很威,
    但若同期買進持有就賺 +900%,那策略其實是在扣分。
    跨市場比較若不對照 B&H,會把「市場的 beta」誤當成「策略的 alpha」。
"""
from __future__ import annotations

import numpy as np


def _returns(equity: np.ndarray) -> np.ndarray:
    equity = np.asarray(equity, float)
    with np.errstate(divide="ignore", invalid="ignore"):
        r = equity[1:] / equity[:-1] - 1.0
    r = r[np.isfinite(r)]
    return r


def _max_drawdown(equity: np.ndarray):
    """回傳 (最大回撤比例, 最長回撤持續 bar 數)。"""
    equity = np.asarray(equity, float)
    peak = -np.inf
    max_dd = 0.0
    longest = 0
    cur_len = 0
    for x in equity:
        # 只有「嚴格低於前高」才算回撤中;持平(= 前高,如暖身期)不算乾旱。
        if x >= peak:
            peak = x
            cur_len = 0
        else:
            cur_len += 1
            longest = max(longest, cur_len)
        if peak > 0:
            dd = (peak - x) / peak
            max_dd = max(max_dd, dd)
    return max_dd, longest


def _trim_flat_prefix(equity: np.ndarray) -> np.ndarray:
    """去掉最前面「持平的暖身段」(權益還沒動),讓風險指標與 B&H 可比。
    保留最後一根持平 bar,使第一筆變動能被算進報酬序列。"""
    e = np.asarray(equity, float)
    i = 0
    while i < len(e) - 1 and e[i] == e[0]:
        i += 1
    start = max(i - 1, 0)
    return e[start:]


def equity_metrics(equity: np.ndarray, years: float, ppy: float) -> dict:
    equity = np.asarray(equity, float)
    init, final = equity[0], equity[-1]
    total_ret = final / init - 1.0 if init else 0.0
    cagr = (final / init) ** (1.0 / years) - 1.0 if (init > 0 and final > 0 and years > 0) else float("nan")
    # 風險指標用「去掉暖身持平段」的報酬序列(避免被前導 0 稀釋、且與 B&H 對齊)
    r = _returns(_trim_flat_prefix(equity))
    ann_vol = float(np.std(r) * np.sqrt(ppy)) if len(r) else 0.0
    mean_r = float(np.mean(r)) if len(r) else 0.0
    sharpe = (mean_r / np.std(r) * np.sqrt(ppy)) if (len(r) and np.std(r) > 0) else 0.0
    downside = r[r < 0]
    dstd = float(np.std(downside)) if len(downside) else 0.0
    sortino = (mean_r / dstd * np.sqrt(ppy)) if dstd > 0 else 0.0
    max_dd, longest = _max_drawdown(equity)
    calmar = (cagr / max_dd) if (max_dd > 0 and not np.isnan(cagr)) else float("nan")
    return {
        "total_return": total_ret,
        "cagr": cagr,
        "ann_vol": ann_vol,
        "sharpe": sharpe,
        "sortino": sortino,
        "max_dd": max_dd,
        "longest_dd_bars": longest,
        "calmar": calmar,
        "return_over_maxdd": (total_ret / max_dd) if max_dd > 0 else float("nan"),
    }


def trade_metrics(trades: list) -> dict:
    n = len(trades)
    if n == 0:
        return {"n_trades": 0, "win_rate": 0.0, "profit_factor": 0.0,
                "expectancy": 0.0, "avg_win": 0.0, "avg_loss": 0.0,
                "payoff": 0.0, "avg_bars": 0.0, "pct_long": 0.0}
    pnls = np.array([t["pnl"] for t in trades], float)
    wins = pnls[pnls > 0]
    losses = pnls[pnls < 0]        # 嚴格虧損才算 loss;打平(pnl==0)不列入輸贏
    gross_win = float(wins.sum())
    gross_loss = float(-losses.sum())
    pf = gross_win / gross_loss if gross_loss > 0 else (float("inf") if gross_win > 0 else 0.0)
    avg_win = float(wins.mean()) if len(wins) else 0.0
    avg_loss = float(losses.mean()) if len(losses) else 0.0
    return {
        "n_trades": n,
        "win_rate": len(wins) / n,
        "profit_factor": pf,
        "expectancy": float(pnls.mean()),
        "avg_win": avg_win,
        "avg_loss": avg_loss,
        "payoff": (avg_win / abs(avg_loss)) if avg_loss != 0 else float("inf"),
        "avg_bars": float(np.mean([t["bars"] for t in trades])),
        "pct_long": float(np.mean([1 if t["dir"] > 0 else 0 for t in trades])),
    }


def bars_per_year(data) -> float:
    """實測「每年幾根 bar」= 資料筆數 / 年數。

    用實測值年化,才能同時對上真實資料(股票 ~252 交易日/年、幣 ~365)
    與合成資料(連續日曆日 ~365),避免硬編 ppy 與實際 bar 間距打架。
    """
    y = data.years()
    return (len(data) / y) if y > 0 else 252.0


def buy_and_hold(data, ppy: float = None) -> dict:
    """買進持有基準:第一根收盤全押到最後。"""
    c = np.asarray(data.c, float)
    equity = c / c[0]
    if ppy is None:
        ppy = bars_per_year(data)
    m = equity_metrics(equity, data.years(), ppy)
    m["strategy"] = "buy_hold"
    return m


def compute_metrics(result) -> dict:
    ppy = bars_per_year(result.data)   # 實測年化,策略與 B&H 用同一把尺
    years = result.data.years()
    m = equity_metrics(result.equity, years, ppy)
    m.update(trade_metrics(result.trades))
    m["exposure"] = float(np.mean(result.in_position))
    m["years"] = years
    m["strategy"] = result.strategy
    m["ticker"] = result.ticker
    m["market"] = result.market
    # 附基準與超額
    bh = buy_and_hold(result.data)
    m["bh_total_return"] = bh["total_return"]
    m["bh_cagr"] = bh["cagr"]
    m["bh_max_dd"] = bh["max_dd"]
    m["excess_cagr"] = (m["cagr"] - bh["cagr"]) if not (np.isnan(m["cagr"]) or np.isnan(bh["cagr"])) else float("nan")
    return m
