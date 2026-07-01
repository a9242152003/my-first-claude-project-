"""
統一資料層 — 三種來源產出同一個 OHLCV 物件,策略/引擎完全不需要知道資料哪來的。

  1. yfinance : 真實歷史資料(你在本機跑,網路通時用這個)
  2. csv      : 自備的 CSV(離線、或自訂資料源)
  3. synthetic: 合成資料(沙盒/CI/展示用;可造出不同「市場性格」)

真實 vs 合成的意義:
  合成資料是用來「驗證程式邏輯正確、並展示跨市場論點」的;
  要看真實績效,在本機用 source="yfinance"(見 run_backtest.py)。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

import numpy as np


@dataclass
class OHLCV:
    ticker: str
    market: str                 # 成本模型 key(見 costs.MARKETS)
    dates: list                 # 長度 N 的日期(datetime.date 或字串)
    o: np.ndarray
    h: np.ndarray
    l: np.ndarray
    c: np.ndarray
    v: np.ndarray
    meta: dict = field(default_factory=dict)

    def __len__(self) -> int:
        return len(self.c)

    def years(self) -> float:
        """資料涵蓋的年數(用實際日期;算不出時退回 交易日/252)。"""
        try:
            import datetime as _dt
            d0, d1 = self.dates[0], self.dates[-1]
            if isinstance(d0, str):
                d0 = _dt.date.fromisoformat(d0[:10])
                d1 = _dt.date.fromisoformat(d1[:10])
            days = (d1 - d0).days
            if days > 0:
                return days / 365.25
        except Exception:
            pass
        return len(self.c) / 252.0


# ── yfinance 真實資料 ──────────────────────────────────────────
def load_yfinance(ticker: str, market: str, period: str = "max",
                  interval: str = "1d") -> OHLCV:
    import yfinance as yf
    df = yf.download(ticker, period=period, interval=interval,
                     progress=False, auto_adjust=True)  # auto_adjust: 還原股息,對齊 TradingView ADJ
    if df is None or len(df) < 50:
        raise ValueError(f"{ticker}: 資料不足或抓取失敗")
    try:
        import pandas as pd
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)
    except Exception:
        pass
    df = df.dropna(subset=["Open", "High", "Low", "Close"])
    dates = [d.date() if hasattr(d, "date") else d for d in df.index.tolist()]
    v = df["Volume"].fillna(0).to_numpy(float) if "Volume" in df.columns else np.zeros(len(df))
    return OHLCV(
        ticker=ticker, market=market, dates=dates,
        o=df["Open"].to_numpy(float), h=df["High"].to_numpy(float),
        l=df["Low"].to_numpy(float), c=df["Close"].to_numpy(float), v=v,
    )


# ── CSV ────────────────────────────────────────────────────────
def load_csv(path: str, ticker: str, market: str) -> OHLCV:
    import pandas as pd
    df = pd.read_csv(path)
    cols = {c.lower(): c for c in df.columns}
    def col(*names):
        for n in names:
            if n in cols:
                return cols[n]
        raise KeyError(f"CSV 缺少欄位 {names};有的欄位:{list(df.columns)}")
    date_c = col("date", "datetime", "time")
    dates = [str(x)[:10] for x in df[date_c].tolist()]
    v_c = cols.get("volume")
    v = df[v_c].fillna(0).to_numpy(float) if v_c else np.zeros(len(df))
    return OHLCV(
        ticker=ticker, market=market, dates=dates,
        o=df[col("open")].to_numpy(float), h=df[col("high")].to_numpy(float),
        l=df[col("low")].to_numpy(float), c=df[col("close")].to_numpy(float), v=v,
    )


# ── 合成資料 ────────────────────────────────────────────────────
# kind 對應不同「市場性格」,用來展示「同策略跨市場結果不同」:
#   trending_hivol : 高波動 + 動能延續 + 肥尾(像 BTC/成長股)→ 突破的主場
#   long_bull      : 穩定長多 + 中波動 + 真回撤(像 0050)      → 突破做多騎趨勢
#   mean_reverting : 圍繞慢速均值來回、負自相關(像大盤指數)  → 突破常被巴,逆勢佔優
#   low_vol_range  : 低波動箱型(像躺平的權值股)              → 兩者都難賺,突破慢性失血
#   bear           : 結構性長空(給空單測試)
_KINDS = ("trending_hivol", "long_bull", "mean_reverting", "low_vol_range", "bear")


def gen_synthetic(kind: str, n: int = 1500, seed: int = 0,
                  start_price: float = 100.0, market: str = "crypto") -> OHLCV:
    if kind not in _KINDS:
        raise KeyError(f"未知合成類型 '{kind}';可用:{', '.join(_KINDS)}")
    rng = np.random.default_rng(seed)

    logp = np.empty(n)
    logp[0] = np.log(start_price)

    if kind in ("mean_reverting", "low_vol_range"):
        # OU 均值回歸:圍繞一條慢速漂移的均值來回;報酬呈負自相關 → 突破常被巴、逆勢佔優。
        kappa = 0.16 if kind == "mean_reverting" else 0.22
        sigma = 0.013 if kind == "mean_reverting" else 0.008
        slow_drift = 0.00012 if kind == "mean_reverting" else 0.0
        mean = logp[0]
        for t in range(1, n):
            mean += slow_drift
            logp[t] = logp[t - 1] + kappa * (mean - logp[t - 1]) + sigma * rng.standard_normal()
    else:
        # 動能:regime-switching(持續性牛/熊段)+ 動能自相關 + 肥尾跳空。
        # 持續的方向性趨勢 → 突破能吃到、逆勢(fade)會被趨勢輾過。
        # (bull_drift, bear_drift, p_stay_bull, p_stay_bear, p_start_bull, sigma, rho, jump_p, jump_s)
        # 牛段黏、熊段短而急 → 淨值長期向上但有大回撤(像 BTC / 狂飆成長股)。
        params = {
            "trending_hivol": (0.0022, -0.0050, 0.985, 0.915, 0.85, 0.026, 0.15, 0.018, 0.08),
            "long_bull":      (0.0018, -0.0045, 0.990, 0.910, 0.90, 0.013, 0.06, 0.010, 0.05),
            "bear":           (0.0040, -0.0026, 0.910, 0.988, 0.12, 0.026, 0.12, 0.018, 0.08),
        }
        bud, bed, psb, pse, p0, sigma, rho, jp, js = params[kind]
        state = 1 if rng.random() < p0 else -1
        r_prev = 0.0
        for t in range(1, n):
            if state == 1:
                if rng.random() >= psb:
                    state = -1
            else:
                if rng.random() >= pse:
                    state = 1
            drift = bud if state == 1 else bed
            eps = rng.standard_normal()
            jump = (rng.standard_normal() * js) if rng.random() < jp else 0.0
            r = drift + rho * r_prev + sigma * eps + jump
            logp[t] = logp[t - 1] + r
            r_prev = r

    c = np.exp(logp)

    # 由收盤序列構造 OHLC:開盤貼近前收(小跳空),高低含盤中噪音。
    o = np.empty(n); h = np.empty(n); l = np.empty(n)
    daily_ret = np.diff(c, prepend=c[0]) / np.maximum(c, 1e-9)
    for t in range(n):
        prev_c = c[t - 1] if t > 0 else c[0]
        gap = rng.normal(0, 0.2) * abs(daily_ret[t]) * prev_c
        o[t] = prev_c + gap
        rng_span = (abs(sigma) + 0.5 * abs(daily_ret[t])) * c[t]
        hi_ext = abs(rng.normal(0, 0.6)) * rng_span
        lo_ext = abs(rng.normal(0, 0.6)) * rng_span
        h[t] = max(o[t], c[t]) + hi_ext
        l[t] = min(o[t], c[t]) - lo_ext
        l[t] = max(l[t], 0.01)

    # 成交量:基準量 + 噪音,並在大漲跌日放量(讓 RVOL 突破確認有意義)。
    base = 1_000_000.0
    noise = rng.lognormal(mean=0.0, sigma=0.3, size=n)
    spike = 1.0 + 6.0 * np.abs(daily_ret) / (np.abs(daily_ret).mean() + 1e-9) * (rng.random(n) < 0.15)
    v = base * noise * spike

    # 日期:從一個固定起點往後排(週一~週五近似;此處用連續日曆日即可,年化用實際天數)
    import datetime as _dt
    start = _dt.date(2015, 1, 1)
    dates = [start + _dt.timedelta(days=int(i)) for i in range(n)]

    return OHLCV(ticker=f"SYN::{kind}", market=market, dates=dates,
                 o=o, h=h, l=l, c=c, v=v, meta={"kind": kind, "seed": seed})


# ── 統一入口 ────────────────────────────────────────────────────
def load(source: str, **kw) -> OHLCV:
    if source == "yfinance":
        return load_yfinance(kw["ticker"], kw["market"],
                             kw.get("period", "max"), kw.get("interval", "1d"))
    if source == "csv":
        return load_csv(kw["path"], kw["ticker"], kw["market"])
    if source == "synthetic":
        return gen_synthetic(kw["kind"], kw.get("n", 1500), kw.get("seed", 0),
                             kw.get("start_price", 100.0), kw.get("market", "crypto"))
    raise KeyError(f"未知資料來源 '{source}';可用:yfinance / csv / synthetic")
