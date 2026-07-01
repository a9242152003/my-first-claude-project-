# swingbt — 跨市場波段策略回測框架

> 回答一個問題:**同一套波段策略,能不能跨市場(美股 / 台股 / 大盤 / 加密貨幣)通用?**

這是本專案既有 Pine 策略(`breakout_v5`、`reversal_v1_sfp`)的 Python 回測版,
把「策略」與「市場」徹底分離,讓你用**同一支策略跑不同市場、直接並排比較**。

---

## TL;DR — 這套框架給出的答案

**不行,不能無腦通用。策略要配市場性格:**

| 市場性格 | 該用的策略類型 | 為什麼 |
|---|---|---|
| 高波動 × 強趨勢(幣、狂飆成長股、長多 ETF) | **動能 / 趨勢**(突破、均線) | 趨勢會延續,讓獲利奔跑 |
| 均值回歸(大盤指數) | **逆勢 / 均值回歸**(RSI/布林) | 漲多會拉回、跌深會反彈 |
| 低波動箱型(躺平權值股) | **幾乎都難賺** | 波幅吃不過交易成本 |

合成資料的蒙地卡羅比較矩陣(每格 = 9 個 seed 的中位數 Calmar):

| 策略＼市場 | 加密(動能) | 成長股(動能) | 長多ETF | 大盤(均值回歸) | 低波箱型 |
|---|---|---|---|---|---|
| 突破 breakout | -0.05 | -0.05 | **0.35** | -0.25 | -0.25 |
| 逆勢 meanrev | -0.23 | -0.23 | -0.25 | **-0.02(最佳)** | -0.21 |
| 均線 macross | **0.13** | **0.14** | **0.58** | -0.38(被巴慘) | -0.38 |

看點:**趨勢策略(macross/breakout)在趨勢市賺、在盤整市被巴到爆(-0.38);
逆勢(meanrev)反過來。同一支策略換個市場,好壞完全翻盤。**

> ⚠️ 這是**合成資料**的結果,用來驗證引擎邏輯與展示論點。
> **真實績效**請在本機用 `--source yfinance`(你的實測 BTC/0050 成績才是真的)。

---

## 安裝 & 快速上手

```bash
pip install numpy pandas          # 核心
pip install yfinance              # 要抓真實資料才需要

# 1) 合成資料(離線可跑,展示 + 驗證邏輯)
python run_backtest.py

# 2) 真實資料(本機、網路通時)
python run_backtest.py --source yfinance

# 只跑某些策略
python run_backtest.py --strategies breakout,meanrev

# 調風險/資料長度/seed 數
python run_backtest.py --risk-pct 3 --n 2000 --seeds 15
```

輸出:console 比較矩陣 + `swingbt_report.md`(完整明細表)。

---

## 架構(策略與市場分離)

```
swingbt/
  data.py         統一資料層:yfinance / csv / synthetic 三選一 → 同一個 OHLCV
  indicators.py   技術指標(全部無未來函數:bar i 只用 <= i 的資料)
  costs.py        各市場成本模型(手續費/證交稅/滑點/可否做空/槓桿/交易日)
  engine.py       事件驅動引擎(close 進場;盤中觸損;停損→目標→訊號→時間 出場)
  metrics.py      績效(CAGR/Sharpe/Sortino/MaxDD/Calmar/PF... + 買進持有基準)
  report.py       策略 × 市場比較矩陣(含蒙地卡羅多 seed 聚合)
  strategies/
    base.py               策略介面
    breakout.py           放量突破(動能)= 你的 breakout_v5
    bollinger_reversion.py RSI/布林 逆勢(均值回歸)
    mean_reversion.py     SFP 沒力反轉 = 你的 reversal_v1_sfp(忠實移植)
    ma_cross.py           均線交叉(對照基準)
run_backtest.py   CLI
tests/test_swingbt.py  健全性測試(含未來函數洩漏測試、論點回歸測試)
```

### 為什麼成本模型很重要
台股有 **0.3% 賣出證交稅** + 0.1425%/邊手續費;加密貨幣約 0.055%/邊、可槓桿可空;
美股近零手續費。不建模這些,跨市場比較就是假的(台股會被高估、指數的空單在美股可行、在台股不可行)。

---

## 內建策略

| 名稱 | 類型 | 進場 | 出場 |
|---|---|---|---|
| `breakout` | 動能 | 放量突破近10根箱頂/底 + EMA200 濾網;牛市不做空 | 吊燈移動停利 3.5ATR + 保本 |
| `meanrev` | 均值回歸 | 順大勢逆小勢:多頭買 RSI(2) 超賣/破下軌;空頭空超買/破上軌 | **回到均值(短均線)就走** + 災難停損 + 時間停損 |
| `sfp` | 均值回歸(專門化) | SFP 掃單收回 + 影線 + 否決閘(你的 pine 忠實移植) | 目標回均值 + 結構停損 + 時間停損 |
| `macross` | 趨勢 | 快慢 EMA 交叉 | 吊燈移動停利(純基準,不調校) |

**關鍵設計**:逆勢用「回到均值就走」的訊號出場,趨勢用「吊燈讓獲利奔跑」——
出場哲學不同,是兩類策略的分水嶺(這也是為什麼把 SFP 硬套吊燈會虧)。

---

## 加自己的策略(3 步)

1. 在 `swingbt/strategies/` 新增一個檔案,繼承 `Strategy`,實作
   `prepare(data)` / `entry(i)`(可選 `on_entry`/`update_stop`/`exit_signal`/`max_hold`)。
2. 在 `strategies/__init__.py` 的 `REGISTRY` 登記名稱。
3. `python run_backtest.py --strategies 你的名字` 即可跨市場比較。

> 鐵律:`entry(i)` 與所有指標**只能用 `<= i` 的資料**。`tests/test_swingbt.py`
> 有自動的「截短資料前後值不變」洩漏測試會抓你。

---

## 誠實前提(和主 README 一致)

1. **回測 ≠ 未來**:標的會「變性」,靠健康儀表板 + 定期重驗 + 熔斷應對(見 `../健康檢查.md`)。
2. **報酬靠少數怪物單**:趨勢策略勝率常 <50%,靠賺賠比。
3. **合成資料只驗邏輯**:真實 edge 要用 `--source yfinance` 在真資料上看。
4. **成本已建模但仍偏樂觀**:未計永續資金費、借券成本、大單衝擊。
