"""
RSI/布林 均值回歸(逆勢)—— 教科書級「買超賣、賣超買、回到均值就走」的波段逆勢。
(精神同 Larry Connors RSI-2:高勝率、小幅獲利、快進快出。)

代表「均值回歸」這一整類策略,用來對照突破(動能)類,回答
『同策略能否跨市場通用』:動能市場它該輸,盤整/回檔市場它該贏。

  進場(順大勢逆小勢):
    多 = 站在 EMA200 之上(多頭)且 RSI(2) 超賣 / 跌破布林下軌 → 買回檔。
    空 = 在 EMA200 之下(空頭)且 RSI(2) 超買 / 突破布林上軌 → 空反彈。
  出場(關鍵!逆勢不用趨勢跟蹤):
    - 訊號出場:價格「收復」短均線(回到均值)→ 立刻走(小賺、高勝率)。
    - 停損:進場 ∓ k×ATR 的災難停損(較寬,避免被正常噪音掃掉)。
    - 時間停損:太久沒回歸就認賠出場。
  「回到均值就走」是它跟突破(讓獲利奔跑)最大的分水嶺。
"""
from __future__ import annotations

from typing import Optional

import numpy as np

from .. import indicators as ind
from .base import Strategy, Entry, Position


class BollingerReversion(Strategy):
    name = "meanrev"

    def __init__(self, bb_len: int = 20, bb_k: float = 2.0, rsi_len: int = 2,
                 rsi_buy: float = 10.0, rsi_sell: float = 90.0, trend_ema: int = 200,
                 stop_atr: float = 3.0, exit_ma: int = 5, hold_bars: int = 12):
        self.bb_len = bb_len; self.bb_k = bb_k; self.rsi_len = rsi_len
        self.rsi_buy = rsi_buy; self.rsi_sell = rsi_sell; self.trend_ema = trend_ema
        self.stop_atr = stop_atr; self.exit_ma = exit_ma; self.hold_bars = hold_bars

    @property
    def warmup(self) -> int:
        return self.trend_ema

    @property
    def max_hold(self):
        return self.hold_bars

    def prepare(self, data) -> None:
        c = data.c
        self._c = c
        self._ema = ind.ema(c, self.trend_ema)
        self._atr = ind.atr(data.h, data.l, data.c, 14)
        self._rsi = ind.rsi(c, self.rsi_len)
        self._bbup, _, self._bbdn = ind.bbands(c, self.bb_len, self.bb_k)
        self._exit_ma = ind.sma(c, self.exit_ma)

    def entry(self, i: int) -> Optional[Entry]:
        atr = self._atr[i]
        c = self._c
        if np.isnan(atr) or np.isnan(self._bbdn[i]) or np.isnan(self._rsi[i]) or np.isnan(self._exit_ma[i]):
            return None
        up_trend = c[i] > self._ema[i]
        dn_trend = c[i] < self._ema[i]
        oversold = self._rsi[i] <= self.rsi_buy or c[i] < self._bbdn[i]
        overbought = self._rsi[i] >= self.rsi_sell or c[i] > self._bbup[i]

        if up_trend and oversold:
            return Entry(direction=+1, stop=c[i] - self.stop_atr * atr, reason="rev_long")
        if dn_trend and overbought:
            return Entry(direction=-1, stop=c[i] + self.stop_atr * atr, reason="rev_short")
        return None

    def on_entry(self, i: int, pos: Position) -> None:
        pos.target = None  # 逆勢用訊號出場,不設固定目標

    def exit_signal(self, i: int, pos: Position) -> bool:
        """價格收復短均線 = 已回到均值 → 走。"""
        if pos.direction > 0:
            return self._c[i] >= self._exit_ma[i]
        else:
            return self._c[i] <= self._exit_ma[i]

    def update_stop(self, i: int, pos: Position) -> float:
        return pos.stop  # 災難停損固定不動
