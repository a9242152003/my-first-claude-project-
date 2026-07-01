"""可插拔策略。每支策略只描述「何時進場、停損怎麼設、怎麼移動停利」,
與市場/資料/成本完全解耦 —— 這正是回答『同策略能否跨市場通用』的前提。"""
from .base import Strategy, Entry, Position  # noqa: F401
from .breakout import BreakoutMomentum  # noqa: F401
from .bollinger_reversion import BollingerReversion  # noqa: F401
from .mean_reversion import SFPMeanReversion  # noqa: F401
from .ma_cross import MACrossover  # noqa: F401

# 策略登記表:名稱 -> 建構子
#   breakout : 放量突破(動能)—— 你的主策略
#   meanrev  : 布林/RSI 逆勢(均值回歸)—— 代表逆勢類
#   sfp      : SFP 沒力反轉 —— 你 pine 的忠實移植(專門化的逆勢)
#   macross  : 均線交叉 —— 對照基準
REGISTRY = {
    "breakout": BreakoutMomentum,
    "meanrev": BollingerReversion,
    "sfp": SFPMeanReversion,
    "macross": MACrossover,
}


def build(name: str, **kw) -> Strategy:
    if name not in REGISTRY:
        raise KeyError(f"未知策略 '{name}';可用:{', '.join(REGISTRY)}")
    return REGISTRY[name](**kw)
