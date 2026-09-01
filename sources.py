"""台股資料來源 —— FinMind 為主，yfinance / yahooquery 補位。

沿用美股版的契約：沒有的能力回傳空表 / None，sync.py 會自動往下一層找。
但台股的資料現實跟美股不同，有三點是重寫的理由：

  1. FinMind 一家就能補滿 Raw 層（實測 2330 / 2408 / 5289 都有 30 季），
     不像美股要靠四層瀑布才湊得齊。所以瀑布變短，FinMind 之後只剩補位。
  2. 台股一律曆年制，沒有 52/53 週制，fiscal_from_period_end 大幅簡化。
  3. 還原股價是付費功能，改成「存原始價 + 存除權息事件」，係數在 Calc 層算。


★★★ 三個會靜靜算錯的坑，動這個檔案前務必先讀 ★★★

  坑 A：`EquityAttributableToOwnersOfParent` 這個 key 在兩個 dataset 意義完全不同
        · TaiwanStockFinancialStatements → 「淨利（淨損）歸屬於母公司業主」= 淨利
        · TaiwanStockBalanceSheet        → 「歸屬於母公司業主之權益合計」  = 權益
        台積電 2026Q2：淨利 7,066 億 vs 權益 6.43 兆，差 900 倍。
        剛好就是 ROE 的分子和分母。接錯不會報錯，只會得到一個荒謬但看似合理的數字。
        本檔的防線：兩個 dataset 分開 pivot，各自從自己的白名單取值，
        再加 _sanity_net_income() 用「淨利不可能是營收的數倍」擋一次。

  坑 B：資產負債表每個欄位都有一個 `_per` 影子欄位（佔資產總額的百分比），
        而且 origin_name 跟本尊完全相同：
            Equity      權益總額  6.474471e+12
            Equity_per  權益總額  6.906000e+01
        用中文欄名比對會抓到哪一筆看順序決定，抓到 69.06 當權益 → ROE 爆掉。
        本檔的防線：_wide() 一律先剔除 type 結尾為 "_per" 的列。

  坑 C：FinMind 的財報是**單季制**，不是累計制。
        實測 2330 的 2024 四季營收加總 = 2.894 兆 = 全年營收；
        EPS 加總 = 45.26 = 全年 EPS。所以拿到什麼就是什麼，**不要再 diff**。
        （原 Colab 的 finmind_get_eps_rolling 多做了一次 groupby.diff()，
          會望遠鏡式全部抵銷，TTM 只剩最新一季，約為真值的 1/3。）


其他實測到的事實（2026-08 驗證）：
  · 免費層可用：TaiwanStockFinancialStatements / BalanceSheet / MonthRevenue /
    Dividend / Price
  · 需付費：TaiwanStockPriceAdj（回 HTTP 400 "Your level is free"）
  · 額度：無 token 300 次/小時，有 token 600 次/小時；超過回 HTTP 402
"""
import os
import re
import time
import datetime as dt

import numpy as np
import pandas as pd
import requests

# ───────────────────────── 資料契約 ─────────────────────────
# 欄名刻意沿用美股版（shares_diluted / eps_diluted_adj），讓 build_xlsx.py 與
# export_json.py 不用改。台股的實際口徑差異記在 eps_basis / roe_basis 欄位裡。
FIN_COLS = ["period_end", "fy", "fq", "revenue", "gross_profit", "net_income",
            "shares_diluted", "treasury_shares", "total_equity"]
EPS_COLS = ["period_end", "eps"]
MREV_COLS = ["month_end", "revenue"]                    # 月營收（台股獨有）
DIV_COLS = ["ex_date", "cash_dividend", "stock_ratio"]  # 除權息事件（還原用）

PAR_VALUE = 10.0        # 台股普通股面額固定 10 元 → 股數 = 股本 ÷ 10
TW_CODE = re.compile(r"^[1-9]\d{3}$")


def empty(cols):
    return pd.DataFrame(columns=cols)


def symbol_candidates(code):
    """yfinance 的台股寫法：上市 .TW、上櫃 .TWO。兩個都試，能回資料的就是它。"""
    code = str(code).strip()
    return [f"{code}.TW", f"{code}.TWO"]


def fiscal_from_period_end(period_end):
    """由期末日推出 (fy, fq)。

    台股一律曆年制，沒有美股那種 52/53 週制的漂移，所以不需要美股版那套
    「分數月距四捨五入」的換算。但期末日偶爾會落在次月初（例如來源把
    2024-03-31 寫成 2024-04-01），所以往前退 10 天再取月份，吸收邊界誤差。
    """
    d = period_end - dt.timedelta(days=10)
    return int(d.year), int((d.month - 1) // 3 + 1)


def quarter_end(fy, fq):
    """(fy, fq) → 該季的期末日。取代美股版 `+91 天` 的近似算法。"""
    m = fq * 3
    last = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]
    if m == 2 and (fy % 4 == 0 and (fy % 100 != 0 or fy % 400 == 0)):
        last = 29
    return dt.date(fy, m, last)


def statutory_deadline(fy, fq):
    """台股財報的法定申報期限。

    美股是逐檔公布 next_earnings；台股沒有這種東西，只有法規期限：
        Q1 / Q2 / Q3 季報 → 季後 45 天
        Q4（年報）        → 年度結束後 75 天
    所以 sync.py 的 plan() 用這個推算「財報可能已公布」，取代 meta.next_earnings。
    （金控／銀行／保險的期限法規上不同，如果之後納入要在這裡分流。）
    """
    end = quarter_end(fy, fq)
    return end + dt.timedelta(days=75 if fq == 4 else 45)


def _f(v):
    """寬鬆轉 float，轉不動回 NaN。"""
    try:
        if v is None or v in ("", "None", "-"):
            return np.nan
        x = float(v)
        return x if np.isfinite(x) else np.nan
    except (TypeError, ValueError):
        return np.nan


def log(msg):
    print(msg, flush=True)


# ───────────────────────── FinMind 底層 client ─────────────────────────
class _FM:
    """FinMind 的 HTTP 層：快取、節流、額度與付費層的處理集中在這裡。

    三種要分開處理的失敗：
      · HTTP 402 → 額度用盡。整趟同步剩下的呼叫全部放棄，不要繼續空轉，
                   讓 sync.py 用既有資料，下一輪再補。
      · HTTP 400 且訊息含 "level is free" → 這個 dataset 要付費。
                   記進 _paid，之後同一個 dataset 直接跳過，不要每檔都撞一次。
      · 其他      → 單次失敗，回空表讓瀑布往下走。
    """
    BASE = "https://api.finmindtrade.com/api/v4/data"
    _cache = {}
    _paid = set()
    _last_call = 0.0
    quota_exhausted = False

    # 免費層 300 次/小時、有 token 600 次/小時 → 每次間隔 0.3 秒綽綽有餘，
    # 主要目的是不要在 GitHub Actions 上瞬間打爆。
    MIN_INTERVAL = float(os.getenv("FINMIND_MIN_INTERVAL", "0.3"))

    @classmethod
    def token(cls):
        return os.getenv("FINMIND_TOKEN", "").strip()

    @classmethod
    def reset(cls):
        """給測試用；正式流程一趟只跑一次，不需要呼叫。"""
        cls._cache.clear()
        cls._paid.clear()
        cls.quota_exhausted = False

    @classmethod
    def get(cls, dataset, stock_id, start_date, end_date=None):
        """回傳長格式 DataFrame（date / stock_id / type / value / origin_name）。"""
        key = (dataset, str(stock_id), start_date, end_date)
        if key in cls._cache:
            return cls._cache[key]
        if cls.quota_exhausted or dataset in cls._paid:
            return pd.DataFrame()

        gap = cls.MIN_INTERVAL - (time.time() - cls._last_call)
        if gap > 0:
            time.sleep(gap)

        params = {"dataset": dataset, "data_id": str(stock_id),
                  "start_date": start_date}
        if end_date:
            params["end_date"] = end_date
        headers = {}
        if cls.token():
            headers["Authorization"] = f"Bearer {cls.token()}"

        out = pd.DataFrame()
        try:
            r = requests.get(cls.BASE, params=params, headers=headers, timeout=30)
            cls._last_call = time.time()
            if r.status_code == 200:
                out = pd.DataFrame(r.json().get("data", []))
            elif r.status_code == 402:
                cls.quota_exhausted = True
                log("      ✗ FinMind 額度用盡（402），本輪剩餘呼叫全部略過")
            elif r.status_code == 400 and "level is free" in r.text:
                cls._paid.add(dataset)
                log(f"      ⚠ FinMind {dataset} 需要付費方案，已略過（不再重試）")
            else:
                log(f"      ! FinMind {dataset} {stock_id} → HTTP {r.status_code}")
        except Exception as e:
            cls._last_call = time.time()
            log(f"      ! FinMind {dataset} {stock_id} 連線失敗: {e}")

        cls._cache[key] = out
        return out

    @staticmethod
    def wide(df):
        """長格式 → 寬格式（index=期末日, columns=type）。

        ★ 坑 B ★ 資產負債表每欄都有 `_per` 影子欄位（佔資產總額的百分比），
        origin_name 與本尊完全相同。這裡一律先剔除，是整個檔案最重要的一行。
        """
        if df is None or df.empty or "type" not in df.columns:
            return pd.DataFrame()
        df = df[~df["type"].astype(str).str.endswith("_per")].copy()
        if df.empty:
            return pd.DataFrame()
        df["value"] = pd.to_numeric(df["value"], errors="coerce")
        w = df.pivot_table(index="date", columns="type", values="value",
                           aggfunc="last")
        w.index = pd.to_datetime(w.index, errors="coerce")
        return w[~w.index.isna()].sort_index(ascending=False)


def _pick(wide, names):
    """依白名單順序取一欄，回傳 Series；都沒有就回 None。"""
    for n in names:
        if n in wide.columns:
            s = pd.to_numeric(wide[n], errors="coerce")
            if s.notna().any():
                return s
    return None


# 損益表：白名單有順序，前面的優先
REVENUE_T = ["Revenue"]
GROSS_T = ["GrossProfit"]
COGS_T = ["CostOfGoodsSold"]
# ★ 坑 A ★ 這裡的 EquityAttributableToOwnersOfParent 是「淨利歸屬母公司」。
# 沒有子公司的公司可能只有 IncomeAfterTaxes，所以往下補。
NET_INCOME_T = ["EquityAttributableToOwnersOfParent", "IncomeAfterTaxes",
                "IncomeFromContinuingOperations"]
EPS_T = ["EPS"]

# 金融業（金控／銀行／保險）的損益表沒有「營業收入」「營業毛利」。
# 實際的 type 名稱還沒實測過，先留白 —— 寧可空著也不要憑印象填錯的字串。
# 填法：拿一檔金控（例如 2882）跑 dump_types()，把「淨收益」那一欄的 type 貼進來。
FINANCIAL_REVENUE_T = []          # TODO: 待金控實測後填入

# 資產負債表
# ★ 坑 A ★ 這裡的 EquityAttributableToOwnersOfParent 是「歸屬母公司權益」。
EQUITY_T = ["EquityAttributableToOwnersOfParent", "Equity"]
CAPITAL_T = ["OrdinaryShare", "CapitalStock"]
TREASURY_T = ["NumberOfSharesInEntityHeldByEntityAndByItsSubsidiaries"]


def _sanity_net_income(rev, ni, ticker):
    """★ 坑 A 的第二道防線 ★

    如果不小心把「權益」當成「淨利」取走，數字會大到離譜（台積電 2026Q2
    權益 6.43 兆 vs 營收 1.27 兆，比值 5）。真實的淨利率極少超過營收的 1.5 倍
    （業外一次性收益才可能），所以用這條線攔一次。
    攔到就把該季淨利設成 NaN —— 寧可缺一欄，也不要讓 ROE 帶著鬼數字上線。
    """
    if rev is None or ni is None:
        return ni
    bad = rev.notna() & (rev > 0) & ni.notna() & (ni.abs() > rev * 1.5)
    n = int(bad.sum())
    if n:
        log(f"      ⚠ {ticker} 有 {n} 季的淨利大於營收 1.5 倍，已捨棄"
            f"（極可能是抓到權益欄位，請檢查 NET_INCOME_T）")
        ni = ni.copy()
        ni[bad] = np.nan
    return ni


class Source:
    """沒有的能力回傳空表 / None，sync.py 會自動往下一層找。"""
    name = "base"
    needs_key = False

    @staticmethod
    def quarterly_financials(ticker, n=24):
        return empty(FIN_COLS)

    @staticmethod
    def quarterly_eps_street(ticker, n=24):
        return empty(EPS_COLS)

    @staticmethod
    def quarterly_eps_gaap(ticker, n=24):
        return empty(EPS_COLS)

    @staticmethod
    def month_revenue(ticker, months=36):
        return empty(MREV_COLS)

    @staticmethod
    def dividend_events(ticker, since=None):
        return empty(DIV_COLS)

    @staticmethod
    def dividends(ticker, since=None):
        return pd.Series(dtype="float64")

    @staticmethod
    def price_history(ticker, since=None):
        return pd.Series(dtype="float64")

    @staticmethod
    def estimates(ticker):
        return {}


# ───────────────────────── ① FinMind ─────────────────────────
class FM(Source):
    """主力來源。實測 2330 / 2408 / 5289 三檔的財報、資產負債表、月營收
    都是 2019-03 起共 30 季／92 個月，零缺漏。免金鑰可用（額度較低）。"""
    name = "fm"
    needs_key = False

    @staticmethod
    def _start(n_quarters):
        """多抓兩年當緩衝，避免邊界剛好切掉最舊那一季。"""
        years = max(2, n_quarters // 4 + 2)
        return (dt.date.today() - dt.timedelta(days=365 * years)).isoformat()

    @staticmethod
    def dump_types(ticker, dataset="TaiwanStockFinancialStatements"):
        """診斷用：印出某檔某 dataset 的 type ↔ origin_name 對照。

        金融股的欄位對應就是靠這個補的 —— 跑一次金控，把「淨收益」那一欄的
        type 名稱填進 FINANCIAL_REVENUE_T，而不是憑印象猜字串。
        """
        raw = _FM.get(dataset, ticker, FM._start(4))
        if raw.empty:
            log(f"{ticker} {dataset} 無資料")
            return raw
        latest = raw[raw["date"] == raw["date"].max()]
        cols = [c for c in ["type", "origin_name", "value"] if c in latest.columns]
        print(latest[cols].to_string(index=False))
        return latest

    @staticmethod
    def quarterly_financials(ticker, n=24):
        start = FM._start(n)
        inc = _FM.wide(_FM.get("TaiwanStockFinancialStatements", ticker, start))
        bs = _FM.wide(_FM.get("TaiwanStockBalanceSheet", ticker, start))
        if inc.empty:
            return empty(FIN_COLS)

        rev = _pick(inc, REVENUE_T)
        if rev is None and FINANCIAL_REVENUE_T:
            rev = _pick(inc, FINANCIAL_REVENUE_T)      # 金融業
        gp = _pick(inc, GROSS_T)
        if gp is None:
            cogs = _pick(inc, COGS_T)
            if rev is not None and cogs is not None:
                gp = rev - cogs
        ni = _sanity_net_income(rev, _pick(inc, NET_INCOME_T), ticker)

        eq = _pick(bs, EQUITY_T) if not bs.empty else None
        cap = _pick(bs, CAPITAL_T) if not bs.empty else None
        tre = _pick(bs, TREASURY_T) if not bs.empty else None

        rows = []
        for ts in list(inc.index)[:n]:
            pe = ts.date()
            fy, fq = fiscal_from_period_end(pe)
            # 股本 ÷ 面額 10 元 = 期末股數。台積電實測：
            #   股本 2.5932e11 ÷ 10 = 259.32 億股；淨利 7.0656e11 ÷ 259.32 億股
            #   = 27.25 元，與 EPS 欄位完全吻合 → 這條路是通的。
            shares = np.nan
            if cap is not None and ts in cap.index:
                c = _f(cap.get(ts))
                shares = c / PAR_VALUE if pd.notna(c) and c > 0 else np.nan
            rows.append(dict(
                period_end=pe, fy=fy, fq=fq,
                revenue=_f(rev.get(ts)) if rev is not None else np.nan,
                gross_profit=_f(gp.get(ts)) if gp is not None else np.nan,
                net_income=_f(ni.get(ts)) if ni is not None else np.nan,
                shares_diluted=shares,
                treasury_shares=_f(tre.get(ts)) if tre is not None else np.nan,
                total_equity=_f(eq.get(ts)) if eq is not None else np.nan))
        return pd.DataFrame(rows, columns=FIN_COLS)

    @staticmethod
    def quarterly_eps_street(ticker, n=24):
        """★ 坑 C ★ FinMind 是單季制，拿到什麼就是什麼，**不要 diff**。

        台股沒有街頭 / GAAP 兩套口徑，EPS 就是財報上那一個（基本每股盈餘）。
        這裡沿用美股版的方法名，是為了讓 sync.py 的 waterfall_eps 不用改。
        """
        raw = _FM.wide(_FM.get("TaiwanStockFinancialStatements",
                               ticker, FM._start(n)))
        if raw.empty:
            return empty(EPS_COLS)
        s = _pick(raw, EPS_T)
        if s is None:
            return empty(EPS_COLS)
        s = s.dropna()
        return pd.DataFrame({"period_end": [t.date() for t in s.index],
                             "eps": s.values})[:n]

    @staticmethod
    def month_revenue(ticker, months=36):
        """月營收 —— 台股獨有，美股完全沒有這一層。

        只回傳原始的「當月營收」。月增率／年增率／累計年增率全部是衍生指標，
        留給 Calc 層算 —— 跟季度那張表同一個原則。
        （FinMind 的 revenue_month / revenue_year 是「月份」「年份」的數字，
          不是月營收／年營收，名字很容易誤會。)
        """
        start = (dt.date.today() - dt.timedelta(days=31 * (months + 6))).isoformat()
        raw = _FM.get("TaiwanStockMonthRevenue", ticker, start)
        if raw.empty or "revenue" not in raw.columns:
            return empty(MREV_COLS)
        df = raw.copy()
        df["month_end"] = pd.to_datetime(df["date"], errors="coerce")
        df = df.dropna(subset=["month_end"]).sort_values("month_end",
                                                         ascending=False)
        df["revenue"] = pd.to_numeric(df["revenue"], errors="coerce")
        df = df.dropna(subset=["revenue"]).head(months)
        df["month_end"] = df["month_end"].dt.date
        return df[MREV_COLS].reset_index(drop=True)

    @staticmethod
    def dividend_events(ticker, since=None):
        """除權息事件 —— 這張表同時解掉兩個坑。

        坑 4（除權息／還原股價）：還原股價是付費功能（實測回 400
              "Your level is free"），所以我們存原始收盤價，用這張表在
              Calc 層自己算還原係數。而且係數每次同步重算，不會過期 ——
              存成快照的話，下次配息後所有歷史係數就全錯了。

        坑 3（股票股利稀釋股數）：配股率是明確的原始數字，股數還原變成
              確定性的累乘，不是美股版那種「猜倍率 ±4%」的啟發式。

        單位：股利欄位是「元／股」，面額 10 元，所以配 1 元股票股利
              = 每股配 0.1 股 → stock_ratio = 股票股利 ÷ 10。
        """
        start = since.isoformat() if since else "2015-01-01"
        raw = _FM.get("TaiwanStockDividend", ticker, start)
        if raw.empty:
            return empty(DIV_COLS)

        def col(name):
            return pd.to_numeric(raw[name], errors="coerce").fillna(0.0) \
                if name in raw.columns else 0.0

        cash = col("CashEarningsDistribution") + col("CashStatutorySurplus")
        stock = col("StockEarningsDistribution") + col("StockStatutorySurplus")

        # 除息日與除權日欄位是分開的，多數台股同一天，但不保證。
        # 取兩者中有值的那一個；都有就取較早的（配股配息通常同日）。
        ex = pd.to_datetime(raw.get("CashExDividendTradingDate"), errors="coerce")
        ex_s = pd.to_datetime(raw.get("StockExDividendTradingDate"), errors="coerce")
        ex = ex.fillna(ex_s) if ex is not None else ex_s

        out = pd.DataFrame({"ex_date": ex,
                            "cash_dividend": cash,
                            "stock_ratio": stock / PAR_VALUE})
        out = out.dropna(subset=["ex_date"])
        out = out[(out.cash_dividend > 0) | (out.stock_ratio > 0)]
        if out.empty:
            return empty(DIV_COLS)
        out["ex_date"] = out["ex_date"].dt.date
        out = (out.groupby("ex_date", as_index=False)
                  .agg({"cash_dividend": "sum", "stock_ratio": "sum"})
                  .sort_values("ex_date", ascending=False))
        return out[DIV_COLS].reset_index(drop=True)

    @staticmethod
    def dividends(ticker, since=None):
        """相容用：回傳「除息日 → 現金股利」的 Series，讓 sync.py 的 dps
        邏輯不用改。配股的部分在 dividend_events 裡，兩者不重複計算。"""
        ev = FM.dividend_events(ticker, since)
        if ev.empty:
            return pd.Series(dtype="float64")
        s = pd.Series(ev.cash_dividend.values,
                      index=pd.to_datetime(ev.ex_date)).sort_index()
        return s[s > 0]

    @staticmethod
    def price_history(ticker, since=None):
        """原始收盤價（未還原）。

        還原股價 TaiwanStockPriceAdj 是付費方案（實測三檔全部 HTTP 400），
        所以 Raw 層存原始價，還原留給 Calc 層用 dividend_events 算。
        這樣 PE 用的是「當時的實際股價配當時的 EPS」—— 同一個時點的真實估值，
        混用還原價會讓歷史 PE 失真。
        """
        start = since.isoformat() if since else \
            (dt.date.today() - dt.timedelta(days=365 * 7)).isoformat()
        raw = _FM.get("TaiwanStockPrice", ticker, start)
        if raw.empty or "close" not in raw.columns:
            return pd.Series(dtype="float64")
        df = raw.copy()
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["close"] = pd.to_numeric(df["close"], errors="coerce")
        df = df.dropna(subset=["date", "close"]).sort_values("date")
        return pd.Series(df["close"].values, index=df["date"])


# ───────────────────────── ② yfinance（.TW / .TWO） ─────────────────────────
class YF(Source):
    """補位用。台股的季報 yfinance 也有，但欄位比 FinMind 少、季數比較淺，
    所以排在 FinMind 後面只補缺的。它真正不可取代的是分析師共識。"""
    name = "yf"
    needs_key = False
    _sym = {}

    @staticmethod
    def _t(ticker):
        """.TW 與 .TWO 都試，能回季報的就是它。結果快取起來 ——
        每檔判斷一次要打兩次 API，不快取的話整趟會多出好幾十次呼叫。"""
        import yfinance as yf
        code = str(ticker).strip()
        if code in YF._sym:
            return yf.Ticker(YF._sym[code])
        for sym in symbol_candidates(code):
            try:
                t = yf.Ticker(sym)
                q = t.quarterly_income_stmt
                if q is not None and len(q) > 0:
                    YF._sym[code] = sym
                    return t
            except Exception:
                continue
        YF._sym[code] = f"{code}.TW"
        return yf.Ticker(YF._sym[code])

    @staticmethod
    def _norm(s):
        return re.sub(r"[^a-z0-9]", "", str(s).lower())

    @staticmethod
    def _row(df, names):
        if df is None or len(df) == 0:
            return None
        m = {YF._norm(i): i for i in df.index}
        for n in names:
            k = YF._norm(n)
            if k in m:
                return pd.to_numeric(df.loc[m[k]], errors="coerce")
        return None

    @staticmethod
    def quarterly_financials(ticker, n=24):
        try:
            t = YF._t(ticker)
            inc = t.quarterly_income_stmt
            bs = t.quarterly_balance_sheet
            if inc is None or len(inc) == 0:
                return empty(FIN_COLS)
            inc = inc.copy()
            inc.columns = pd.to_datetime(inc.columns, errors="coerce")
            inc = inc.loc[:, ~pd.isna(inc.columns)].sort_index(axis=1,
                                                               ascending=False)
            rev = YF._row(inc, ["Total Revenue", "Revenue", "Operating Revenue"])
            gp = YF._row(inc, ["Gross Profit"])
            if gp is None:
                cogs = YF._row(inc, ["Cost Of Revenue"])
                if rev is not None and cogs is not None:
                    gp = rev - cogs
            ni = YF._row(inc, ["Net Income", "Net Income Common Stockholders"])
            eq = None
            if bs is not None and len(bs):
                bs = bs.copy()
                bs.columns = pd.to_datetime(bs.columns, errors="coerce")
                eq = YF._row(bs, ["Stockholders Equity",
                                  "Total Stockholder Equity",
                                  "Total Equity Gross Minority Interest"])
            rows = []
            for c in list(inc.columns)[:n]:
                pe = pd.Timestamp(c).date()
                fy, fq = fiscal_from_period_end(pe)
                rows.append(dict(
                    period_end=pe, fy=fy, fq=fq,
                    revenue=_f(rev.get(c)) if rev is not None else np.nan,
                    gross_profit=_f(gp.get(c)) if gp is not None else np.nan,
                    net_income=_f(ni.get(c)) if ni is not None else np.nan,
                    shares_diluted=np.nan,
                    treasury_shares=np.nan,
                    total_equity=_f(eq.get(c)) if eq is not None else np.nan))
            return pd.DataFrame(rows, columns=FIN_COLS)
        except Exception:
            return empty(FIN_COLS)

    @staticmethod
    def quarterly_eps_street(ticker, n=24):
        try:
            inc = YF._t(ticker).quarterly_income_stmt
            if inc is None or len(inc) == 0:
                return empty(EPS_COLS)
            inc = inc.copy()
            inc.columns = pd.to_datetime(inc.columns, errors="coerce")
            s = YF._row(inc, ["Diluted EPS", "Basic EPS"])
            if s is None:
                return empty(EPS_COLS)
            s = s.dropna()
            return pd.DataFrame({"period_end": [pd.Timestamp(c).date()
                                                for c in s.index],
                                 "eps": s.values})[:n]
        except Exception:
            return empty(EPS_COLS)

    @staticmethod
    def price_history(ticker, since=None):
        try:
            h = YF._t(ticker).history(period="max", auto_adjust=False)
            if h is None or len(h) == 0:
                return pd.Series(dtype="float64")
            s = h["Close"]
            if since is not None:
                s = s[s.index.date >= since]
            return s
        except Exception:
            return pd.Series(dtype="float64")

    @staticmethod
    def estimates(ticker):
        """台股的分析師共識覆蓋率低，抓不到是常態 —— 回空 dict，
        讓 PE Forward / PEG_F 誠實顯示 N/M，不要用自己的推估去填。"""
        out = {}
        try:
            info = YF._t(ticker).info or {}
            if info.get("forwardEps"):
                out["eps_f1"] = float(info["forwardEps"])
        except Exception:
            pass
        return out


# ───────────────────────── ③ yahooquery ─────────────────────────
class YQ(Source):
    """只為了共識預估而存在。earnings_trend 對台股是時有時無，
    大型權值股（2330、2454）多半有，中小型多半沒有 —— 這是預期內的。"""
    name = "yq"
    needs_key = False

    @staticmethod
    def _sym(ticker):
        code = str(ticker).strip()
        return YF._sym.get(code) or f"{code}.TW"

    @staticmethod
    def estimates(ticker):
        out = {}
        try:
            from yahooquery import Ticker
            sym = YQ._sym(ticker)
            d = Ticker(sym, country="taiwan", formatted=False).earnings_trend
            if not isinstance(d, dict):
                return out
            obj = d.get(sym)
            if not isinstance(obj, dict):
                return out                       # 'Invalid Crumb' 會是字串
            for r in obj.get("trend", []) or []:
                per = str(r.get("period", "")).strip().lower()
                est = r.get("earningsEstimate") or {}
                avg = est.get("avg")
                if isinstance(avg, dict):
                    avg = avg.get("raw")
                if avg is None:
                    continue
                na = est.get("numberOfAnalysts")
                if isinstance(na, dict):
                    na = na.get("raw")
                if per in ("0q", "currentquarter"):
                    out["eps_q0"] = float(avg)
                    if na:
                        out["n_analysts"] = int(na)
                elif per in ("0y", "currentyear"):
                    out["eps_f1"] = float(avg)
                    if r.get("endDate"):
                        try:
                            out["fy1_end"] = pd.to_datetime(r["endDate"]).date()
                        except Exception:
                            pass
                elif per in ("+1y", "nextyear"):
                    out["eps_f2"] = float(avg)
        except Exception:
            pass
        return out


REGISTRY = {"fm": FM, "yf": YF, "yq": YQ}


def chain(names):
    return [REGISTRY[n] for n in names if n in REGISTRY]
