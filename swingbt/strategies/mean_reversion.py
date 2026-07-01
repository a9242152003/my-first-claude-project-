"""
SFP 沒力反轉(均值回歸)—— 忠實移植你的 reversal_v1_sfp.pine 核心。

假設:大盤指數是均值回歸市,「買回檔 / 空反彈」比硬套突破合適。
  觸發:SFP 掃單收回(掃過近10根高/低又收回)+ 有意義影線。
  順勢逆勢:EMA200 上彎只做多回檔、下彎只做空反彈(牛市抓回檔沒力、熊市抓反彈沒力)。
  否決閘:布林擠壓連3收在帶外 / 單根 >3ATR 巨棒 / 逆波幅度不足 → 不進。
  加分:RSI/MACD/z 背離收斂(convScore),≥2 時風險加成 1.5x(不單獨觸發)。
  停損:結構停損(近5根擺動極值 ±0.5ATR);吊燈移動停利 3.0ATR。
"""
from __future__ import annotations

from typing import Optional

import numpy as np

from .. import indicators as ind
from .base import Strategy, Entry, Position


class SFPMeanReversion(Strategy):
    name = "meanrev"

    def __init__(self, sweep_n: int = 10, wick_atr: float = 0.25, slope_lb: int = 20,
                 wave_lb: int = 3, swing_len: int = 5, stop_buf: float = 0.5,
                 big_atr: float = 3.0, min_wave: float = 1.5,
                 rsi_hi: float = 58, rsi_lo: float = 42, z_th: float = 2.0,
                 conv_min: int = 0, conv_boost: float = 1.5, ema_len: int = 200,
                 target_r_min: float = 1.0, target_r_max: float = 3.0,
                 hold_bars: int = 20, be_after_r: float = 1.0):
        self.sweep_n = sweep_n; self.wick_atr = wick_atr; self.slope_lb = slope_lb
        self.wave_lb = wave_lb; self.swing_len = swing_len; self.stop_buf = stop_buf
        self.big_atr = big_atr; self.min_wave = min_wave
        self.rsi_hi = rsi_hi; self.rsi_lo = rsi_lo; self.z_th = z_th
        self.conv_min = conv_min; self.conv_boost = conv_boost; self.ema_len = ema_len
        # 均值回歸專屬出場:目標回到均值(以 R 為單位夾住)、保本、時間停損
        self.target_r_min = target_r_min; self.target_r_max = target_r_max
        self.hold_bars = hold_bars; self.be_after_r = be_after_r

    @property
    def warmup(self) -> int:
        return self.ema_len

    @property
    def max_hold(self):
        return self.hold_bars

    def prepare(self, data) -> None:
        o, h, l, c = data.o, data.h, data.l, data.c
        self._o, self._h, self._l, self._c = o, h, l, c
        self._ema200 = ind.ema(c, self.ema_len)
        self._ema50 = ind.ema(c, 50)
        self._atr = ind.atr(h, l, c, 14)
        self._rsi = ind.rsi(c, 14)
        _, _, self._hist = ind.macd(c, 12, 26, 9)
        self._bbup, _, self._bbdn = ind.bbands(c, 20, 2.0)
        # z = (close-ema50) 相對其 100 日均值/標準差
        devm = c - self._ema50
        zmean = ind.sma(devm, 100)
        zstd = ind.stdev(devm, 100)
        with np.errstate(divide="ignore", invalid="ignore"):
            self._z = np.where((zstd > 0), (devm - zmean) / zstd, 0.0)
        self._sweep_hi = ind.rolling_max(h, self.sweep_n, shift=1)
        self._sweep_lo = ind.rolling_min(l, self.sweep_n, shift=1)
        self._swing_hi = ind.rolling_max(h, self.swing_len, shift=1)
        self._swing_lo = ind.rolling_min(l, self.swing_len, shift=1)

    def _conv_short(self, i: int) -> int:
        rsi, hist, z = self._rsi, self._hist, self._z
        s = 0
        if rsi[i - 1] >= self.rsi_hi and rsi[i] < rsi[i - 1]:
            s += 1
        if hist[i] > 0 and hist[i] < hist[i - 1] and hist[i - 1] < hist[i - 2]:
            s += 1
        if z[i - 1] >= self.z_th and z[i] < z[i - 1]:
            s += 1
        return s

    def _conv_long(self, i: int) -> int:
        rsi, hist, z = self._rsi, self._hist, self._z
        s = 0
        if rsi[i - 1] <= self.rsi_lo and rsi[i] > rsi[i - 1]:
            s += 1
        if hist[i] < 0 and hist[i] > hist[i - 1] and hist[i - 1] > hist[i - 2]:
            s += 1
        if z[i - 1] <= -self.z_th and z[i] > z[i - 1]:
            s += 1
        return s

    def entry(self, i: int) -> Optional[Entry]:
        if i < max(self.slope_lb, self.wave_lb, 2):
            return None
        o, h, l, c = self._o, self._h, self._l, self._c
        atr = self._atr[i]
        if np.isnan(atr) or np.isnan(self._sweep_hi[i]) or np.isnan(self._z[i]):
            return None

        ema200 = self._ema200
        bear = ema200[i] < ema200[i - self.slope_lb]
        bull = ema200[i] > ema200[i - self.slope_lb]

        rally_up = c[i] > c[i - self.wave_lb] and c[i] < ema200[i]  # 熊市反彈但仍在均線下
        pull_dn = c[i] < c[i - self.wave_lb] and c[i] > ema200[i]   # 牛市回檔但仍在均線上

        sweep_high = h[i] > self._sweep_hi[i] and c[i] < self._sweep_hi[i]
        sweep_low = l[i] < self._sweep_lo[i] and c[i] > self._sweep_lo[i]
        wick_up = (h[i] - max(o[i], c[i])) > self.wick_atr * atr
        wick_dn = (min(o[i], c[i]) - l[i]) > self.wick_atr * atr

        sfp_short = bear and rally_up and sweep_high and wick_up
        sfp_long = bull and pull_dn and sweep_low and wick_dn

        # 否決閘
        big_bar = (h[i] - l[i]) > self.big_atr * atr
        squeeze_short = int(np.sum(c[i - 4:i + 1] > self._bbup[i - 4:i + 1])) >= 3
        squeeze_long = int(np.sum(c[i - 4:i + 1] < self._bbdn[i - 4:i + 1])) >= 3
        wave_ok_short = (h[i] - self._sweep_lo[i]) >= self.min_wave * atr
        wave_ok_long = (self._sweep_hi[i] - l[i]) >= self.min_wave * atr
        veto_short = squeeze_short or big_bar or not wave_ok_short
        veto_long = squeeze_long or big_bar or not wave_ok_long

        # 結構停損
        stop_hi = max(self._swing_hi[i], h[i]) + self.stop_buf * atr
        stop_lo = min(self._swing_lo[i], l[i]) - self.stop_buf * atr
        short_dist = (stop_hi - c[i]) / c[i]
        long_dist = (c[i] - stop_lo) / c[i]

        if sfp_short and not veto_short and short_dist > 0:
            conv = self._conv_short(i)
            if conv >= self.conv_min:
                mult = self.conv_boost if conv >= 2 else 1.0
                return Entry(direction=-1, stop=stop_hi, risk_mult=mult, reason="sfp_short")
        if sfp_long and not veto_long and long_dist > 0:
            conv = self._conv_long(i)
            if conv >= self.conv_min:
                mult = self.conv_boost if conv >= 2 else 1.0
                return Entry(direction=+1, stop=stop_lo, risk_mult=mult, reason="sfp_long")
        return None

    def on_entry(self, i: int, pos: Position) -> None:
        pos.entry_atr = self._atr[i]
        mean = self._ema50[i]           # 回歸目標 = 均值
        risk = abs(pos.entry_price - pos.stop)
        if pos.direction > 0:
            desired = max(mean, pos.entry_price + self.target_r_min * risk)
            desired = min(desired, pos.entry_price + self.target_r_max * risk)
            pos.target = desired
        else:
            desired = min(mean, pos.entry_price - self.target_r_min * risk)
            desired = max(desired, pos.entry_price - self.target_r_max * risk)
            pos.target = desired
        pos.scratch["be"] = False

    def update_stop(self, i: int, pos: Position) -> float:
        """均值回歸不用吊燈移動停利(那是趨勢的出場);改用固定結構停損 + 到 1R 後保本。"""
        risk = abs(pos.entry_price - pos.stop)
        if pos.direction > 0:
            if not pos.scratch["be"] and self._h[i] >= pos.entry_price + self.be_after_r * risk:
                pos.scratch["be"] = True
            return max(pos.stop, pos.entry_price) if pos.scratch["be"] else pos.stop
        else:
            if not pos.scratch["be"] and self._l[i] <= pos.entry_price - self.be_after_r * risk:
                pos.scratch["be"] = True
            return min(pos.stop, pos.entry_price) if pos.scratch["be"] else pos.stop
