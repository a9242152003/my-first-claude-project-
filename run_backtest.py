#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
swingbt CLI — 跨市場波段策略回測。

回答一個問題:同一套策略,能不能跨市場(美股/台股/大盤/加密貨幣)通用?

── 快速上手 ─────────────────────────────────────────────
  # 合成資料(離線可跑,展示論點、驗證邏輯)
  python run_backtest.py

  # 真實資料(本機、網路通時;會用 yfinance 抓)
  pip install yfinance pandas numpy
  python run_backtest.py --source yfinance

  # 只跑某些策略
  python run_backtest.py --strategies breakout,meanrev
"""
from __future__ import annotations

import argparse
import sys

from swingbt import report


# ── 合成展示盤:每個「市場性格」用一種合成資料,跑多個 seed 取中位數 ──
# 目的:即使離線,也能穩健看到「同策略跨市場結果不同」(不靠單一 seed 的運氣)。
SYNTHETIC_MARKETS = [
    {"label": "加密貨幣(高波動動能)", "kind": "trending_hivol", "market": "crypto"},
    {"label": "美股成長股(狂飆單股)", "kind": "trending_hivol", "market": "us_stock"},
    {"label": "台股長多ETF(0050型)", "kind": "long_bull", "market": "tw_stock"},
    {"label": "美股大盤(均值回歸)", "kind": "mean_reverting", "market": "index_us"},
    {"label": "台股權值股(低波箱型)", "kind": "low_vol_range", "market": "tw_stock"},
]

# ── 真實資料盤:ticker -> 市場成本模型。與 validate_all.py 的清單同源。 ──
YFINANCE_CASES = [
    # 加密貨幣
    {"label": "BTC-USD", "source": "yfinance", "ticker": "BTC-USD", "market": "crypto"},
    {"label": "ETH-USD", "source": "yfinance", "ticker": "ETH-USD", "market": "crypto"},
    {"label": "SOL-USD", "source": "yfinance", "ticker": "SOL-USD", "market": "crypto"},
    # 美股成長股
    {"label": "NVDA", "source": "yfinance", "ticker": "NVDA", "market": "us_stock"},
    {"label": "TSLA", "source": "yfinance", "ticker": "TSLA", "market": "us_stock"},
    # 美股大盤(指數/ETF)
    {"label": "S&P500(^GSPC)", "source": "yfinance", "ticker": "^GSPC", "market": "index_us"},
    {"label": "Nasdaq100(QQQ)", "source": "yfinance", "ticker": "QQQ", "market": "index_us"},
    # 台股個股 / ETF / 大盤
    {"label": "台積電(2330)", "source": "yfinance", "ticker": "2330.TW", "market": "tw_stock"},
    {"label": "台灣50(0050)", "source": "yfinance", "ticker": "0050.TW", "market": "index_tw"},
    {"label": "加權指數(^TWII)", "source": "yfinance", "ticker": "^TWII", "market": "index_tw"},
]


def main(argv=None):
    ap = argparse.ArgumentParser(description="跨市場波段策略回測")
    ap.add_argument("--source", choices=["synthetic", "yfinance"], default="synthetic",
                    help="資料來源(預設 synthetic,離線可跑)")
    ap.add_argument("--strategies", default="breakout,meanrev,macross",
                    help="逗號分隔:breakout,meanrev,macross")
    ap.add_argument("--risk-pct", type=float, default=5.0, help="每筆風險 %%(當前權益)")
    ap.add_argument("--init-capital", type=float, default=1_000_000.0)
    ap.add_argument("--out", default="swingbt_report.md", help="輸出 markdown 檔")
    ap.add_argument("--n", type=int, default=1500, help="合成資料長度")
    ap.add_argument("--seeds", type=int, default=9, help="合成:每個市場跑幾個 seed 取中位數")
    args = ap.parse_args(argv)

    strategy_names = [s.strip() for s in args.strategies.split(",") if s.strip()]

    print(f"=== swingbt 回測 | 來源={args.source} | 策略={strategy_names} "
          f"| 風險={args.risk_pct}% ===\n")
    if args.source == "synthetic":
        rows = report.run_grid_mc(SYNTHETIC_MARKETS, strategy_names,
                                  seeds=list(range(args.seeds)), n=args.n,
                                  init_capital=args.init_capital, risk_pct=args.risk_pct)
    else:
        rows = report.run_grid(YFINANCE_CASES, strategy_names,
                               init_capital=args.init_capital, risk_pct=args.risk_pct)
    if not rows:
        print("沒有任何結果(資料都載入失敗?)")
        return 1

    md = report.full_markdown(rows)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(md)
    print("\n" + report.matrix_markdown(rows, "calmar", "Calmar = CAGR / MaxDD"))
    print(f"\n✅ 已寫出完整報表:{args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
