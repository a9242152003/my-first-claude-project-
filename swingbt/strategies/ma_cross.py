"""
均線交叉(對照基準)—— 最經典的波段訊號,拿來當「有沒有比隨便一招好」的對照組。

  進場:快 EMA 上穿慢 EMA(且站上慢線)做多;下穿做空。
  停損:初始 = 進場價 - 3ATR(多);吊燈移動停利 3ATR。
純基準用,參數不調校 —— 目的是給突破/逆勢一個誠實的比較底線。
"""
from __future__ import annotations

from typing import Optional

import numpy as np

from .. import indicators as ind
from .base import Strategy, Entry, Position


class MACrossover(Strategy):
    name = "macross"

    def __init__(self, fast: int = 20, slow: int = 50, stop_atr: float = 3.0,
                 trail_atr: float = 3.0):
        self.fast = fast
        self.slow = slow
        self.stop_atr = stop_atr
        self.trail_atr = trail_atr

    @property
    def warmup(self) -> int:
        return self.slow + 5

    def prepare(self, data) -> None:
        self._h, self._l, self._c = data.h, data.l, data.c
        self._fast = ind.ema(data.c, self.fast)
        self._slow = ind.ema(data.c, self.slow)
        self._atr = ind.atr(data.h, data.l, data.c, 14)

    def entry(self, i: int) -> Optional[Entry]:
        atr = self._atr[i]
        if np.isnan(atr) or i < 1:
            return None
        f, s, c = self._fast, self._slow, self._c
        cross_up = f[i] > s[i] and f[i - 1] <= s[i - 1]
        cross_dn = f[i] < s[i] and f[i - 1] >= s[i - 1]
        if cross_up and c[i] > s[i]:
            return Entry(direction=+1, stop=c[i] - self.stop_atr * atr, reason="ma_up")
        if cross_dn and c[i] < s[i]:
            return Entry(direction=-1, stop=c[i] + self.stop_atr * atr, reason="ma_dn")
        return None

    def update_stop(self, i: int, pos: Position) -> float:
        atr = self._atr[i]
        s = pos.scratch
        if pos.direction > 0:
            s["hh"] = max(s["hh"], self._h[i])
            return max(pos.stop, s["hh"] - self.trail_atr * atr)
        else:
            s["ll"] = min(s["ll"], self._l[i])
            return min(pos.stop, s["ll"] + self.trail_atr * atr)
