"""
技術指標 — 全部「無未來函數」(no look-ahead)。

保證:每個輸出 out[i] 只用到輸入的 <= i(或明確 shift 過的更早)資料。
回傳一律是與輸入等長的 numpy 陣列,暖身不足處填 np.nan。
"""
from __future__ import annotations

import numpy as np


def sma(x: np.ndarray, n: int) -> np.ndarray:
    x = np.asarray(x, dtype=float)
    out = np.full(len(x), np.nan)
    if n <= 0 or len(x) < n:
        return out
    csum = np.cumsum(np.insert(x, 0, 0.0))
    out[n - 1:] = (csum[n:] - csum[:-n]) / n
    return out


def ema(x: np.ndarray, n: int) -> np.ndarray:
    """指數移動平均。第一個值用第一筆 seed(與 TradingView ta.ema 慣例一致)。"""
    x = np.asarray(x, dtype=float)
    out = np.full(len(x), np.nan)
    if len(x) == 0:
        return out
    a = 2.0 / (n + 1.0)
    e = x[0]
    out[0] = e
    for i in range(1, len(x)):
        e = a * x[i] + (1 - a) * e
        out[i] = e
    return out


def rma(x: np.ndarray, n: int) -> np.ndarray:
    """Wilder 平滑(RSI / ATR 用)。a = 1/n。"""
    x = np.asarray(x, dtype=float)
    out = np.full(len(x), np.nan)
    if len(x) == 0:
        return out
    a = 1.0 / n
    e = x[0]
    out[0] = e
    for i in range(1, len(x)):
        e = a * x[i] + (1 - a) * e
        out[i] = e
    return out


def true_range(h: np.ndarray, l: np.ndarray, c: np.ndarray) -> np.ndarray:
    h = np.asarray(h, float); l = np.asarray(l, float); c = np.asarray(c, float)
    n = len(c)
    tr = np.empty(n)
    tr[0] = h[0] - l[0]
    for i in range(1, n):
        pc = c[i - 1]
        tr[i] = max(h[i] - l[i], abs(h[i] - pc), abs(l[i] - pc))
    return tr


def atr(h: np.ndarray, l: np.ndarray, c: np.ndarray, n: int = 14) -> np.ndarray:
    return rma(true_range(h, l, c), n)


def rsi(c: np.ndarray, n: int = 14) -> np.ndarray:
    c = np.asarray(c, float)
    N = len(c)
    out = np.full(N, np.nan)
    if N < 2:
        return out
    delta = np.diff(c, prepend=c[0])
    gain = np.where(delta > 0, delta, 0.0)
    loss = np.where(delta < 0, -delta, 0.0)
    ag = rma(gain, n)
    al = rma(loss, n)
    with np.errstate(divide="ignore", invalid="ignore"):
        rs = np.where(al == 0, np.inf, ag / al)
        out = 100.0 - 100.0 / (1.0 + rs)
    out[np.isinf(rs)] = 100.0
    return out


def stdev(x: np.ndarray, n: int) -> np.ndarray:
    """母體標準差(對齊 TradingView ta.stdev 預設)。"""
    x = np.asarray(x, float)
    N = len(x)
    out = np.full(N, np.nan)
    if n <= 0 or N < n:
        return out
    csum = np.cumsum(np.insert(x, 0, 0.0))
    csum2 = np.cumsum(np.insert(x * x, 0, 0.0))
    s = csum[n:] - csum[:-n]
    s2 = csum2[n:] - csum2[:-n]
    var = s2 / n - (s / n) ** 2
    var = np.clip(var, 0.0, None)
    out[n - 1:] = np.sqrt(var)
    return out


def rolling_max(x: np.ndarray, n: int, shift: int = 1) -> np.ndarray:
    """近 n 根最高值;shift=1 表示「不含當根」(= 前一根結束時已知,無未來函數)。

    out[i] = max(x[i-shift-n+1 .. i-shift])。
    這正是突破策略要的「近 N 根箱頂」— 用在 bar i 判斷突破時,參考的是之前的箱體。
    """
    x = np.asarray(x, float)
    N = len(x)
    out = np.full(N, np.nan)
    for i in range(N):
        hi = i - shift
        lo = hi - n + 1
        if lo < 0:
            continue
        out[i] = np.max(x[lo:hi + 1])
    return out


def rolling_min(x: np.ndarray, n: int, shift: int = 1) -> np.ndarray:
    x = np.asarray(x, float)
    N = len(x)
    out = np.full(N, np.nan)
    for i in range(N):
        hi = i - shift
        lo = hi - n + 1
        if lo < 0:
            continue
        out[i] = np.min(x[lo:hi + 1])
    return out


def rvol(vol: np.ndarray, n: int = 20) -> np.ndarray:
    """相對成交量 = 當根量 / 近 n 根均量。均量用 sma(不含未來)。"""
    vol = np.asarray(vol, float)
    base = sma(vol, n)
    with np.errstate(divide="ignore", invalid="ignore"):
        out = np.where((base > 0) & ~np.isnan(base), vol / base, np.nan)
    return out


def macd(c: np.ndarray, fast: int = 12, slow: int = 26, signal: int = 9):
    c = np.asarray(c, float)
    macd_line = ema(c, fast) - ema(c, slow)
    signal_line = ema(macd_line, signal)
    hist = macd_line - signal_line
    return macd_line, signal_line, hist


def bbands(c: np.ndarray, n: int = 20, k: float = 2.0):
    basis = sma(c, n)
    dev = stdev(c, n)
    return basis + k * dev, basis, basis - k * dev
