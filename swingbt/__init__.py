"""
swingbt — 跨市場波段策略回測框架
================================

一套「策略與市場分離」的回測引擎,用來回答一個問題:

    「同一套波段策略,能不能跨市場(美股 / 台股 / 大盤 / 加密貨幣)通用?」

用法:同一支策略,跑不同市場,輸出比較矩陣,用數據看誰配誰。

模組:
    data        — 統一資料層(yfinance 真實資料 / CSV / 合成資料三選一)
    indicators  — 技術指標(全部無未來函數,bar i 只用 <= i 的資料)
    costs       — 各市場成本模型(手續費 / 證交稅 / 滑點 / 可否做空 / 槓桿)
    engine      — 事件驅動回測引擎(close 進場、盤中觸損出場)
    strategies  — 可插拔策略(突破動能 / SFP 均值回歸 / 均線基準)
    metrics     — 績效指標(CAGR / Sharpe / Sortino / MaxDD / Calmar / PF...)
    report      — 策略 × 市場比較報表(console + markdown)
"""

__version__ = "0.1.0"

from .data import OHLCV, load, gen_synthetic  # noqa: F401
from .costs import CostModel, MARKETS  # noqa: F401
from .engine import Backtester, BacktestResult  # noqa: F401
from .metrics import compute_metrics, buy_and_hold  # noqa: F401
