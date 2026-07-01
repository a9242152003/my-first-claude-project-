"""
策略 × 市場比較報表。

核心用途:把「同一支策略跑不同市場」的結果並排,直接看
「哪個策略配哪個市場」——用數據回答『同策略能否跨市場通用』。
"""
from __future__ import annotations

from typing import Optional

import numpy as np

from . import strategies as strat
from .costs import get_cost_model
from .data import load, gen_synthetic
from .engine import run
from .metrics import compute_metrics, buy_and_hold


def _median(xs):
    xs = [x for x in xs if x is not None and np.isfinite(x)]
    return float(np.median(xs)) if xs else float("nan")


def _mean(xs):
    xs = [x for x in xs if x is not None and np.isfinite(x)]
    return float(np.mean(xs)) if xs else float("nan")


def _aggregate(strategy, label, market, metric_dicts, beats_key="total_return",
               bh_key="bh_total_return"):
    """把同一(策略,市場)在多個 seed 的成績,聚合成中位數 / 平均。"""
    g = lambda k: [m.get(k) for m in metric_dicts]
    beats = [1 for m in metric_dicts if m.get(beats_key, -1) > m.get(bh_key, 0)]
    return {
        "label": label, "market": market, "strategy": strategy,
        "cagr": _median(g("cagr")), "total_return": _median(g("total_return")),
        "max_dd": _median(g("max_dd")), "calmar": _median(g("calmar")),
        "sharpe": _median(g("sharpe")), "profit_factor": _mean(g("profit_factor")),
        "win_rate": _mean(g("win_rate")), "n_trades": _mean(g("n_trades")),
        "exposure": _mean(g("exposure")), "excess_cagr": _median(g("excess_cagr")),
        "beats_bh_rate": len(beats) / len(metric_dicts) if metric_dicts else float("nan"),
        "seeds": len(metric_dicts),
    }


def run_grid_mc(markets: list, strategy_names: list, seeds: list, n: int = 1500,
                init_capital: float = 1_000_000.0, risk_pct: float = 5.0,
                verbose: bool = True) -> list:
    """蒙地卡羅:每個市場性格跑多個 seed,取中位數 —— 避免單一 seed 的運氣誤導。
    markets: list[dict(label, kind, market)]。"""
    rows = []
    for mk in markets:
        cost = get_cost_model(mk["market"])
        # buy & hold 基準(逐 seed)
        bh_ms = []
        for sd in seeds:
            d = gen_synthetic(mk["kind"], n=n, seed=sd, market=mk["market"])
            bh_ms.append(buy_and_hold(d, cost.periods_per_year))
        bh_row = _aggregate("buy_hold", mk["label"], mk["market"], bh_ms,
                            beats_key="total_return", bh_key="total_return")
        bh_row["beats_bh_rate"] = float("nan")
        rows.append(bh_row)

        for sname in strategy_names:
            ms = []
            for sd in seeds:
                d = gen_synthetic(mk["kind"], n=n, seed=sd, market=mk["market"])
                res = run(d, strat.build(sname), cost, init_capital=init_capital,
                          risk_pct=risk_pct)
                ms.append(compute_metrics(res))
            agg = _aggregate(sname, mk["label"], mk["market"], ms)
            rows.append(agg)
            if verbose:
                print(f"  {mk['label']:<22} {sname:<9} "
                      f"CAGR中位 {_pct(agg['cagr']):>8}  MaxDD {_pct(agg['max_dd']):>7}  "
                      f"Calmar {_f(agg['calmar']):>5}  PF均 {_f(agg['profit_factor']):>5}  "
                      f"勝過B&H {_pct(agg['beats_bh_rate']):>5}  ({agg['seeds']} seeds)")
    return rows


def run_grid(cases: list, strategy_names: list, init_capital: float = 1_000_000.0,
             risk_pct: float = 5.0, verbose: bool = True) -> list:
    """cases: list[dict],每個含 label + load 參數(source/market/...)。
    回傳:每個 (策略, 市場) 一列的指標 dict;另含每個市場的 buy_hold 基準列。"""
    rows = []
    for case in cases:
        label = case["label"]
        market = case["market"]
        cost = get_cost_model(market)
        load_kw = {k: v for k, v in case.items() if k not in ("label",)}
        try:
            data = load(**load_kw)
        except Exception as e:
            if verbose:
                print(f"  ⚠️  {label}: 載入失敗 {e}")
            continue

        # 基準:買進持有
        bh = buy_and_hold(data, cost.periods_per_year)
        rows.append({"label": label, "market": market, "strategy": "buy_hold",
                     "cagr": bh["cagr"], "max_dd": bh["max_dd"],
                     "total_return": bh["total_return"], "sharpe": bh["sharpe"],
                     "calmar": bh["calmar"], "n_trades": np.nan, "win_rate": np.nan,
                     "profit_factor": np.nan, "exposure": 1.0,
                     "excess_cagr": 0.0, "years": data.years()})

        for sname in strategy_names:
            s = strat.build(sname)
            res = run(data, s, cost, init_capital=init_capital, risk_pct=risk_pct)
            m = compute_metrics(res)
            m["label"] = label
            rows.append(m)
            if verbose:
                print(f"  {label:<22} {sname:<9} "
                      f"CAGR {_pct(m['cagr']):>8}  MaxDD {_pct(m['max_dd']):>7}  "
                      f"Calmar {_f(m['calmar']):>5}  PF {_f(m['profit_factor']):>5}  "
                      f"trades {m['n_trades']:>3}  超額CAGR {_pct(m['excess_cagr']):>8}")
    return rows


# ── 格式化小工具 ──
def _pct(x) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return f"{x*100:.1f}%"


def _f(x) -> str:
    if x is None or (isinstance(x, float) and (np.isnan(x) or np.isinf(x))):
        return "—"
    return f"{x:.2f}"


def matrix_markdown(rows: list, metric: str = "calmar",
                    metric_label: str = "Calmar(CAGR/MaxDD)") -> str:
    """把結果攤成 策略×市場 矩陣(值為指定指標)。"""
    markets = []
    for r in rows:
        if r["label"] not in markets:
            markets.append(r["label"])
    strategies = []
    for r in rows:
        if r["strategy"] not in strategies:
            strategies.append(r["strategy"])

    def cell(strategy, label):
        for r in rows:
            if r["strategy"] == strategy and r["label"] == label:
                return r.get(metric)
        return None

    out = [f"### 比較矩陣 — {metric_label}(越高越好)\n",
           "| 策略＼市場 | " + " | ".join(markets) + " |",
           "|" + "---|" * (len(markets) + 1)]
    for s in strategies:
        cells = [_f(cell(s, m)) if metric != "cagr" else _pct(cell(s, m)) for m in markets]
        name = "買進持有(基準)" if s == "buy_hold" else s
        out.append(f"| **{name}** | " + " | ".join(cells) + " |")
    return "\n".join(out)


def full_markdown(rows: list, title: str = "跨市場波段策略回測") -> str:
    out = [f"# {title}\n",
           "> 由 `swingbt` 產出。**同一套策略跑不同市場,看誰配誰。**",
           "> 合成資料僅驗證邏輯與展示論點;真實績效請在本機用 `--source yfinance`。\n"]
    out.append(matrix_markdown(rows, "calmar", "Calmar = CAGR / MaxDD"))
    out.append("\n")
    out.append(matrix_markdown(rows, "cagr", "年化報酬 CAGR"))
    out.append("\n### 完整明細\n")
    out.append("| 市場 | 策略 | 年化CAGR | 總報酬 | MaxDD | Calmar | Sharpe | PF | 勝率 | 筆數 | 曝險 | 超額CAGR | 勝過B&H |")
    out.append("|---|---|---|---|---|---|---|---|---|---|---|---|---|")
    for r in rows:
        name = "買進持有" if r["strategy"] == "buy_hold" else r["strategy"]
        nt = r.get("n_trades", np.nan)
        out.append(
            f"| {r['label']} | {name} | {_pct(r.get('cagr'))} | {_pct(r.get('total_return'))} | "
            f"{_pct(r.get('max_dd'))} | {_f(r.get('calmar'))} | {_f(r.get('sharpe'))} | "
            f"{_f(r.get('profit_factor'))} | {_pct(r.get('win_rate'))} | "
            f"{'—' if (nt is None or np.isnan(nt)) else f'{nt:.0f}'} | "
            f"{_pct(r.get('exposure'))} | {_pct(r.get('excess_cagr'))} | "
            f"{_pct(r.get('beats_bh_rate'))} |")
    return "\n".join(out)
