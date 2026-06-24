#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
標的資格自動驗證 — Claude
自己抓資料(yfinance)→ 跑嚴格回測 → 輸出 PASS/FAIL 表。
真回測,不是市場體質的燈,所以不會有「2330 綠燈但回測是屎」。

用法:
    pip install yfinance
    python validate_all.py
輸出:console 表格 + 同目錄 health_table.md(可被 GitHub Action 自動 commit)。

判準(嚴格):前後半各 交易≥15 且 PF≥1.3 且 報酬/回撤≥2,且全期 報酬/回撤≥3 → PASS。
"""
import sys, math
try:
    import yfinance as yf
except ImportError:
    sys.exit("請先安裝:pip install yfinance")

# ── 標的清單(yfinance 代碼)──
ASSETS = {
    "加密前10": ["BTC-USD","ETH-USD","BNB-USD","SOL-USD","XRP-USD","ADA-USD","DOGE-USD","AVAX-USD","TRX-USD","LINK-USD"],
    "美股大盤+前10": ["^GSPC","^IXIC","NVDA","AAPL","MSFT","AMZN","GOOGL","META","TSLA","AVGO","AMD","NFLX"],
    "台股0050+前10": ["0050.TW","^TWII","2330.TW","2317.TW","2454.TW","2308.TW","2382.TW","2891.TW","2412.TW","2881.TW","3711.TW","2603.TW"],
}

# ── 嚴格門檻 ──
MIN_TR_HALF = 15
MIN_PF_HALF = 1.3
MIN_RDD_HALF = 2.0
MIN_RDD_FULL = 3.0
COMM = 0.00055
INIT = 1_000_000.0

def ema(s, n):
    a = 2/(n+1); out=[None]*len(s); e=None
    for i,x in enumerate(s):
        e = x if e is None else a*x+(1-a)*e; out[i]=e
    return out
def rma(s, n):
    a=1/n; out=[None]*len(s); e=None
    for i,x in enumerate(s):
        e = x if e is None else a*x+(1-a)*e; out[i]=e
    return out
def sma(s, n):
    out=[None]*len(s); ss=0.0
    for i in range(len(s)):
        ss+=s[i]
        if i>=n: ss-=s[i-n]
        if i>=n-1: out[i]=ss/n
    return out

def backtest(o,h,l,c,vol,t0=None,t1=None):
    N=len(c)
    ema200=ema(c,200)
    tr=[h[i]-l[i] if i==0 else max(h[i]-l[i],abs(h[i]-c[i-1]),abs(l[i]-c[i-1])) for i in range(N)]
    atr=rma(tr,14); vs=sma(vol,20)
    rvol=[vol[i]/vs[i] if vs[i] else 0 for i in range(N)]
    hh10=[max(h[i-10:i]) if i>=10 else None for i in range(N)]
    ll10=[min(l[i-10:i]) if i>=10 else None for i in range(N)]
    cash=INIT;pos=0;qty=0.0;entry=0.0;pst=0.0;hh=0.0;ll=0.0;atrE=0.0;beA=False
    peak=INIT;maxdd=0.0;tr_=[]
    for i in range(N):
        if pos!=0 and atr[i] is not None:
            ex=None
            if pos>0:
                if l[i]<=pst: ex=o[i] if o[i]<pst else pst
            else:
                if h[i]>=pst: ex=o[i] if o[i]>pst else pst
            if ex is not None:
                pnl=(ex-entry)*qty if pos>0 else (entry-ex)*qty
                pnl-=COMM*qty*(entry+ex); cash+=pnl; tr_.append(pnl); pos=0
            else:
                if pos>0:
                    hh=max(hh,h[i])
                    if not beA and hh>=entry+3.0*atrE: beA=True
                    pst=max(pst,hh-3.5*atr[i])
                    if beA: pst=max(pst,entry)
                else:
                    ll=min(ll,l[i])
                    if not beA and ll<=entry-3.0*atrE: beA=True
                    pst=min(pst,ll+3.5*atr[i])
                    if beA: pst=min(pst,entry)
        if pos==0 and i>=200 and atr[i] is not None and ema200[i-20] is not None and hh10[i] is not None:
            inwin=(t0 is None or i>=t0) and (t1 is None or i<t1)
            if inwin:
                bear=ema200[i]<ema200[i-20]
                longS=c[i]>hh10[i]+0.25*atr[i] and rvol[i]>=1.3 and c[i]>ema200[i]
                shortS=bear and c[i]<ll10[i]-0.25*atr[i] and rvol[i]>=1.3 and c[i]<ema200[i]
                if longS:
                    d=max((c[i]-l[i-1])/c[i],0.005); qn=min(cash*5/100/(d*c[i]),5*cash/c[i])
                    pos=1;qty=qn;entry=c[i];pst=c[i]-d*c[i];atrE=atr[i];hh=h[i];beA=False
                elif shortS:
                    d=max((h[i-1]-c[i])/c[i],0.005); qn=min(cash*5/100/(d*c[i]),5*cash/c[i])
                    pos=-1;qty=qn;entry=c[i];pst=c[i]+d*c[i];atrE=atr[i];ll=l[i];beA=False
        mtm=cash+(0 if pos==0 else (c[i]-entry)*qty if pos>0 else (entry-c[i])*qty)
        peak=max(peak,mtm); maxdd=max(maxdd,(peak-mtm)/peak)
    w=[x for x in tr_ if x>0]; ls=[-x for x in tr_ if x<=0]
    pf=sum(w)/sum(ls) if sum(ls)>0 else (99 if w else 0)
    net=(cash-INIT)/INIT*100; dd=maxdd*100
    return dict(net=net,dd=dd,pf=pf,n=len(tr_),rdd=net/dd if dd>0 else 0)

def validate(ticker):
    try:
        df=yf.download(ticker, period="max", interval="1d", progress=False, auto_adjust=False)
    except Exception as e:
        return None, f"抓取失敗 {e}"
    if df is None or len(df)<300:
        return None, "資料不足"
    o=df["Open"].astype(float).values.tolist()
    h=df["High"].astype(float).values.tolist()
    l=df["Low"].astype(float).values.tolist()
    c=df["Close"].astype(float).values.tolist()
    v=df["Volume"].astype(float).fillna(0).values.tolist() if "Volume" in df else [0]*len(c)
    o=[x[0] if isinstance(x,(list,tuple)) else x for x in o]  # 防多層欄位
    N=len(c); mid=N//2
    full=backtest(o,h,l,c,v)
    A=backtest(o,h,l,c,v,None,mid)
    B=backtest(o,h,l,c,v,mid,None)
    def okhalf(r): return r['n']>=MIN_TR_HALF and r['pf']>=MIN_PF_HALF and r['rdd']>=MIN_RDD_HALF
    passed = okhalf(A) and okhalf(B) and full['rdd']>=MIN_RDD_FULL
    return dict(full=full,A=A,B=B,passed=passed), None

def main():
    rows=[]
    print(f"{'標的':<12}{'全PF':>6}{'全RDD':>7}{'前PF':>6}{'前n':>5}{'後PF':>6}{'後n':>5}  判決")
    for cat, tickers in ASSETS.items():
        print(f"\n=== {cat} ===")
        rows.append(("CAT",cat,None))
        for tk in tickers:
            res,err=validate(tk)
            if err:
                print(f"{tk:<12} {err}")
                rows.append((tk,err,None)); continue
            f,A,B=res['full'],res['A'],res['B']
            verdict="✅ PASS" if res['passed'] else "❌ FAIL"
            print(f"{tk:<12}{f['pf']:>6.2f}{f['rdd']:>7.1f}{A['pf']:>6.2f}{A['n']:>5}{B['pf']:>6.2f}{B['n']:>5}  {verdict}")
            rows.append((tk,res,None))
    # 寫 markdown
    with open("health_table.md","w",encoding="utf-8") as fp:
        fp.write("# 標的資格自動驗證表\n\n")
        fp.write(f"判準:前後半各 交易≥{MIN_TR_HALF} 且 PF≥{MIN_PF_HALF} 且 報酬/回撤≥{MIN_RDD_HALF},全期 報酬/回撤≥{MIN_RDD_FULL} → PASS\n\n")
        for tag,a,b in rows:
            if tag=="CAT":
                fp.write(f"\n## {a}\n\n| 標的 | 全期PF | 全期報酬/回撤 | 前半PF/筆 | 後半PF/筆 | 判決 |\n|---|---|---|---|---|---|\n")
            elif isinstance(a,dict):
                f,A,B=a['full'],a['A'],a['B']
                v="✅ PASS" if a['passed'] else "❌ FAIL"
                fp.write(f"| {tag} | {f['pf']:.2f} | {f['rdd']:.1f} | {A['pf']:.2f}/{A['n']} | {B['pf']:.2f}/{B['n']} | **{v}** |\n")
            else:
                fp.write(f"| {tag} | — | — | — | — | {a} |\n")
    print("\n已寫出 health_table.md")

if __name__=="__main__":
    main()
