"""
放量突破(動能)—— 忠實移植你的 breakout_v5 / validate_all.py。

DNA:收割「高波動 × 強趨勢延續 × 肥尾」。
  進場:收盤突破近10根箱頂/底 + 0.25ATR、RVOL≥1.3、站上/跌破 EMA200;
        做空另需 EMA200 下彎(牛市不做空,全專案最值錢的一步)。
  停損:初始 = 前一根低/高;吊燈移動停利 3.5ATR;獲利達 3.0ATR 後保本。
  倉位:風險定額(由引擎按 risk_pct 計),槓桿上限由市場成本模型決定。
"""
from __future__ import annotations

from typing import Optional

import numpy as np

from .. import indicators as ind
from .base import Strategy, Entry, Position


class BreakoutMomentum(Strategy):
    name = "breakout"

    def __init__(self, box: int = 10, rvol_min: float = 1.3, buf_atr: float = 0.25,
                 trail_atr: float = 3.5, be_arm_atr: float = 3.0, slope_lb: int = 20,
                 ema_len: int = 200):
        self.box = box
        self.rvol_min = rvol_min
        self.buf_atr = buf_atr
        self.trail_atr = trail_atr
        self.be_arm_atr = be_arm_atr
        self.slope_lb = slope_lb
        self.ema_len = ema_len

    @property
    def warmup(self) -> int:
        return self.ema_len

    def prepare(self, data) -> None:
        self._o, self._h, self._l, self._c = data.o, data.h, data.l, data.c
        self._ema = ind.ema(data.c, self.ema_len)
        self._atr = ind.atr(data.h, data.l, data.c, 14)
        self._rvol = ind.rvol(data.v, 20)
        self._hhbox = ind.rolling_max(data.h, self.box, shift=1)
        self._llbox = ind.rolling_min(data.l, self.box, shift=1)

    def entry(self, i: int) -> Optional[Entry]:
        c, atr, ema = self._c, self._atr, self._ema
        if np.isnan(atr[i]) or np.isnan(self._hhbox[i]) or np.isnan(self._rvol[i]):
            return None
        if i < self.slope_lb:
            return None
        bear = ema[i] < ema[i - self.slope_lb]
        rvol_ok = self._rvol[i] >= self.rvol_min

        long_sig = (c[i] > self._hhbox[i] + self.buf_atr * atr[i]
                    and rvol_ok and c[i] > ema[i])
        short_sig = (bear and c[i] < self._llbox[i] - self.buf_atr * atr[i]
                     and rvol_ok and c[i] < ema[i])

        if long_sig:
            d = max((c[i] - self._l[i - 1]) / c[i], 0.005)
            return Entry(direction=+1, stop=c[i] * (1 - d), reason="breakout_long")
        if short_sig:
            d = max((self._h[i - 1] - c[i]) / c[i], 0.005)
            return Entry(direction=-1, stop=c[i] * (1 + d), reason="breakout_short")
        return None

    def on_entry(self, i: int, pos: Position) -> None:
        pos.scratch["hh"] = self._h[i]
        pos.scratch["ll"] = self._l[i]
        pos.scratch["be"] = False
        pos.entry_atr = self._atr[i]

    def update_stop(self, i: int, pos: Position) -> float:
        atr = self._atr[i]
        s = pos.scratch
        if pos.direction > 0:
            s["hh"] = max(s["hh"], self._h[i])
            if not s["be"] and s["hh"] >= pos.entry_price + self.be_arm_atr * pos.entry_atr:
                s["be"] = True
            stop = max(pos.stop, s["hh"] - self.trail_atr * atr)
            if s["be"]:
                stop = max(stop, pos.entry_price)
            return stop
        else:
            s["ll"] = min(s["ll"], self._l[i])
            if not s["be"] and s["ll"] <= pos.entry_price - self.be_arm_atr * pos.entry_atr:
                s["be"] = True
            stop = min(pos.stop, s["ll"] + self.trail_atr * atr)
            if s["be"]:
                stop = min(stop, pos.entry_price)
            return stop
