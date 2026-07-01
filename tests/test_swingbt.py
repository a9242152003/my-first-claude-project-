#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
swingbt 健全性測試。可用 pytest 跑,也可直接 `python tests/test_swingbt.py`。

重點測「回測最容易出的錯」:未來函數(look-ahead)、進出場時點、成本、指標。
"""
import os
import sys

import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from swingbt import indicators as ind
from swingbt.data import gen_synthetic
from swingbt.costs import MARKETS, get_cost_model
from swingbt.engine import run
from swingbt.metrics import equity_metrics, _max_drawdown
from swingbt import strategies as strat


# ── 指標 ───────────────────────────────────────────────
def test_sma_ema_basic():
    x = np.array([1, 2, 3, 4, 5], float)
    s = ind.sma(x, 3)
    assert np.isnan(s[0]) and np.isnan(s[1])
    assert abs(s[2] - 2.0) < 1e-9 and abs(s[4] - 4.0) < 1e-9
    e = ind.ema(x, 3)
    assert abs(e[0] - 1.0) < 1e-9  # seed = 第一筆
    assert e[-1] > e[0]


def test_rsi_bounds():
    x = np.cumsum(np.ones(100)) + np.sin(np.arange(100))
    r = ind.rsi(x, 14)
    r = r[~np.isnan(r)]
    assert (r >= 0).all() and (r <= 100).all()


def test_rolling_max_excludes_current():
    h = np.array([1, 5, 2, 8, 3, 9], float)
    rm = ind.rolling_max(h, 2, shift=1)  # 近2根、不含當根
    # i=3: max(h[1],h[2]) = max(5,2)=5,不含 h[3]=8
    assert abs(rm[3] - 5.0) < 1e-9
    # i=5: max(h[3],h[4]) = max(8,3)=8,不含 h[5]=9
    assert abs(rm[5] - 8.0) < 1e-9


# ── 未來函數洩漏測試(最關鍵)──────────────────────────
def test_no_lookahead_indicators():
    """把資料截短,前段指標值不應改變 —— 證明指標是因果的、沒偷看未來。"""
    d = gen_synthetic("trending_hivol", n=600, seed=42)
    k = 400
    for fn in [
        lambda h, l, c, v: ind.ema(c, 200),
        lambda h, l, c, v: ind.atr(h, l, c, 14),
        lambda h, l, c, v: ind.rsi(c, 14),
        lambda h, l, c, v: ind.rolling_max(h, 10, shift=1),
        lambda h, l, c, v: ind.rvol(v, 20),
        lambda h, l, c, v: ind.stdev(c, 100),
    ]:
        full = fn(d.h, d.l, d.c, d.v)
        trunc = fn(d.h[:k], d.l[:k], d.c[:k], d.v[:k])
        a, b = full[:k], trunc
        both = ~(np.isnan(a) | np.isnan(b))
        assert np.allclose(a[both], b[both], atol=1e-8), "指標偷看了未來!"


def test_no_lookahead_signals():
    """進場訊號也不能因未來資料而改變。逐根重算 entry,截短前後訊號須一致。"""
    d = gen_synthetic("mean_reverting", n=500, seed=11)
    k = 350
    s_full = strat.build("breakout"); s_full.prepare(d)
    d2 = gen_synthetic("mean_reverting", n=500, seed=11)
    # 截短版
    from swingbt.data import OHLCV
    dt = OHLCV(d2.ticker, d2.market, d2.dates[:k], d2.o[:k], d2.h[:k],
               d2.l[:k], d2.c[:k], d2.v[:k])
    s_tr = strat.build("breakout"); s_tr.prepare(dt)
    for i in range(250, k):
        e_full = s_full.entry(i)
        e_tr = s_tr.entry(i)
        df = None if e_full is None else (e_full.direction, round(e_full.stop, 6))
        dtr = None if e_tr is None else (e_tr.direction, round(e_tr.stop, 6))
        assert df == dtr, f"bar {i} 訊號受未來影響:{df} vs {dtr}"


# ── 引擎 ───────────────────────────────────────────────
def test_engine_runs_and_equity_wellformed():
    d = gen_synthetic("trending_hivol", n=800, seed=1)
    cost = get_cost_model("crypto")
    res = run(d, strat.build("breakout"), cost, init_capital=1_000_000)
    assert len(res.equity) == len(d)
    assert np.isfinite(res.equity).all()
    assert res.equity[0] == 1_000_000  # 第一根必空手 = 本金
    # 每筆交易 pnl 與進出價方向一致(扣成本後允許小負)
    for t in res.trades:
        gross = (t["exit"] - t["entry"]) * t["qty"] * t["dir"]
        assert t["pnl"] <= gross + 1e-6  # 成本只會讓 pnl <= gross


def test_realized_risk_matches_intended_with_tight_stop():
    """回歸:策略停損比 min_stop_frac 還緊時,實際虧損仍應 ≈ 預定風險(不可被少算)。"""
    import datetime as dt
    from swingbt.strategies.base import Strategy, Entry
    from swingbt.costs import CostModel
    from swingbt.data import OHLCV

    class TightStop(Strategy):
        name = "tight"
        @property
        def warmup(self): return 2
        def prepare(self, data): pass
        def entry(self, i):
            return Entry(direction=+1, stop=100.0 * 0.999) if i == 2 else None  # 0.1% 停損
        def update_stop(self, i, pos): return pos.stop

    n = 6
    o = np.array([100, 100, 100, 100.0, 100, 100])
    h = np.array([100, 100, 100, 100.0, 100, 100])
    l = np.array([100, 100, 100, 99.0, 100, 100])   # bar3 跌破被觸損
    c = np.array([100, 100, 100, 99.5, 100, 100])
    dates = [dt.date(2020, 1, 1) + dt.timedelta(days=k) for k in range(n)]
    d = OHLCV("T", "crypto", dates, o, h, l, c, np.ones(n) * 1e6)
    free = CostModel("free", 0, 0, 0, True, 10.0, 365)  # 零成本、槓桿夠大不受 cap
    res = run(d, TightStop(), free, init_capital=1_000_000, risk_pct=5.0)
    assert len(res.trades) == 1
    loss = res.trades[0]["pnl"]
    # 預定風險 = 5% * 100萬 = 5萬;停損距離用 min_stop_frac(0.5%),實際觸損虧損應 ≈ -5萬
    assert abs(loss + 50_000) < 500, f"實現風險與預定不符:{loss}"


def test_no_zero_bar_phantom_trades():
    """回歸:最後一根不開倉,不應出現 bars==0 的幽靈單。"""
    from swingbt.costs import get_cost_model
    for kind, market in [("trending_hivol", "crypto"), ("mean_reverting", "index_us")]:
        for sname in ["breakout", "meanrev", "macross"]:
            d = gen_synthetic(kind, n=800, seed=1)
            res = run(d, strat.build(sname), get_cost_model(market))
            assert all(t["bars"] >= 1 for t in res.trades), f"{sname}/{kind} 出現 bars=0 幽靈單"


def test_short_blocked_when_market_disallows():
    """台股成本模型 allow_short=False → 不應有任何空單成交。"""
    d = gen_synthetic("bear", n=800, seed=2)
    cost = get_cost_model("tw_stock")
    assert cost.allow_short is False
    res = run(d, strat.build("breakout"), cost)
    assert all(t["dir"] > 0 for t in res.trades), "台股不該出現空單"


def test_costs_reduce_pnl():
    """同一筆行情,成本越高總報酬越低。"""
    d = gen_synthetic("trending_hivol", n=1000, seed=4)
    r_lo = run(d, strat.build("breakout"), MARKETS["us_stock"]).equity[-1]
    r_hi = run(d, strat.build("breakout"), MARKETS["tw_stock"]).equity[-1]
    # us_stock 成本遠低於 tw_stock(後者有 0.3% 證交稅);但 tw 禁空,
    # 為公平只比多單影響,改用同為可空的 crypto vs 人工高成本不易;
    # 這裡只驗證「成本存在會咬掉報酬」:關掉成本 vs 開成本
    from swingbt.costs import CostModel
    free = CostModel("free", 0, 0, 0, True, 5.0, 365)
    r_free = run(d, strat.build("breakout"), free).equity[-1]
    r_costly = run(d, strat.build("breakout"), MARKETS["crypto"]).equity[-1]
    assert r_free >= r_costly - 1e-6, "加了成本,報酬不應更高"


# ── 指標:回撤 ─────────────────────────────────────────
def test_max_drawdown():
    eq = np.array([100, 120, 90, 110, 80, 130], float)
    dd, longest = _max_drawdown(eq)
    # 峰值 120 -> 谷 80,回撤 = (120-80)/120 = 0.3333
    assert abs(dd - (40 / 120)) < 1e-9


def test_metrics_sane():
    eq = np.linspace(100, 200, 253)  # 一年翻倍、單調上升
    m = equity_metrics(eq, years=1.0, ppy=252)
    assert abs(m["total_return"] - 1.0) < 1e-9
    assert m["max_dd"] == 0.0
    assert m["sharpe"] > 0


# ── 出場機制 ───────────────────────────────────────────
def test_signal_exit_used_by_meanrev():
    """均值回歸應以『訊號式出場』(回到均值就走)為主。"""
    from swingbt.costs import get_cost_model
    d = gen_synthetic("mean_reverting", n=1500, seed=1)
    res = run(d, strat.build("meanrev"), get_cost_model("index_us"))
    reasons = {t["reason"] for t in res.trades}
    assert "signal" in reasons, "均值回歸沒有用到訊號出場"


# ── 核心論點回歸測試(跨市場結果不同)──────────────────
def _median_calmar(kind, market, sname, seeds=range(9), n=1500):
    from swingbt.costs import get_cost_model
    from swingbt.metrics import compute_metrics
    cals = []
    for s in seeds:
        d = gen_synthetic(kind, n=n, seed=s, market=market)
        m = compute_metrics(run(d, strat.build(sname), get_cost_model(market)))
        if np.isfinite(m["calmar"]):
            cals.append(m["calmar"])
    return float(np.median(cals)) if cals else float("nan")


def test_meanrev_beats_trend_in_ranging_market():
    """均值回歸市場:逆勢應優於趨勢跟蹤(macross 在盤整會被巴)。"""
    mr = _median_calmar("mean_reverting", "index_us", "meanrev")
    ma = _median_calmar("mean_reverting", "index_us", "macross")
    assert mr > ma, f"盤整市逆勢應勝趨勢:meanrev {mr:.2f} vs macross {ma:.2f}"


def test_trend_beats_meanrev_in_trending_market():
    """趨勢市場:趨勢跟蹤應優於逆勢(fade 趨勢會虧)。"""
    ma = _median_calmar("long_bull", "tw_stock", "macross")
    mr = _median_calmar("long_bull", "tw_stock", "meanrev")
    assert ma > mr, f"趨勢市趨勢跟蹤應勝逆勢:macross {ma:.2f} vs meanrev {mr:.2f}"


def _run_all():
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_") and callable(v)]
    passed = 0
    for fn in fns:
        fn()
        print(f"  ✓ {fn.__name__}")
        passed += 1
    print(f"\n全部 {passed} 項測試通過 ✅")


if __name__ == "__main__":
    _run_all()
