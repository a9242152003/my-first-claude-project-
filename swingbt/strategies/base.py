"""策略介面(Strategy interface)。

生命週期(由引擎呼叫):
    prepare(data)          一次;預先算好所有指標(必須無未來函數)
    entry(i)               每根「空手」時呼叫;回傳 Entry 或 None(在 close[i] 進場)
    on_entry(i, pos)       進場當根;讓策略初始化這筆單的移動停利暫存(如 hh/ll)
    update_stop(i, pos)    每根「持倉且未觸損」時呼叫;回傳「下一根生效」的新停損

約定:
    - entry(i) 只能用到 <= i 的資料(引擎已在 close[i] 定案)。
    - 停損距離用 entry 與 stop 的差,倉位由引擎按「風險定額」計算。
    - direction: +1 做多, -1 做空。
"""
from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Entry:
    direction: int          # +1 / -1
    stop: float             # 初始停損價
    risk_mult: float = 1.0  # 風險加成(如 SFP convScore>=2 加碼)
    reason: str = ""


@dataclass
class Position:
    direction: int
    entry_price: float
    qty: float
    entry_i: int
    stop: float
    entry_atr: float = 0.0
    target: Optional[float] = None   # 目標價(均值回歸用;None=不設目標,純靠移動停損)
    scratch: dict = field(default_factory=dict)  # 策略自用暫存(hh/ll/be...)


class Strategy(ABC):
    name: str = "base"

    @property
    def warmup(self) -> int:
        """需要多少根暖身才開始交易。"""
        return 200

    @property
    def max_hold(self) -> Optional[int]:
        """時間停損:持倉超過 N 根強制以收盤平倉(None = 不設)。均值回歸常用。"""
        return None

    @abstractmethod
    def prepare(self, data) -> None:
        ...

    @abstractmethod
    def entry(self, i: int) -> Optional[Entry]:
        ...

    def on_entry(self, i: int, pos: Position) -> None:
        """預設:記錄進場當根極值,供吊燈停利使用。"""
        pos.scratch["hh"] = self._h[i]
        pos.scratch["ll"] = self._l[i]

    def exit_signal(self, i: int, pos: Position) -> bool:
        """訊號式出場:回傳 True 則在 close[i] 平倉(用於「回到均值就走」的逆勢)。
        只用 <= i 的資料,故無未來函數。預設不用。"""
        return False

    @abstractmethod
    def update_stop(self, i: int, pos: Position) -> float:
        ...
