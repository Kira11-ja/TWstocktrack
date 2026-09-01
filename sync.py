#!/usr/bin/env python3
"""增量同步（台股版）：只抓缺的，抓完清理，清理後才寫回 master。

骨架與美股版相同 —— 三層架構、seq 樞紐、逐列 sanitize、空值不覆蓋好資料。
台股改的是四處：

  1. plan() 用**法定申報期限**推算財報是否已公布，取代美股逐檔的 next_earnings。
     台股沒有公司預告財報日這種東西，只有法規：Q1~Q3 季後 45 天、Q4 年報 75 天。

  2. 多一條**月營收**的同步線。台股每月 10 日前公布上月營收，是最高頻的基本面
     指標，美股完全沒有這一層。它的節奏跟季報不同步，所以是獨立的一張表
     （data/raw_m.csv），不是季度表的附加欄位。

  3. 拿掉 fix_split_shares。美股要猜分割倍率，是因為 yfinance 的股數沒有回溯調整、
     EPS 卻是調整後的，兩者不同基期。台股不存在這個問題：
        · 股數 = 當期期末股本 ÷ 10
        · EPS  = 當期財報上的每股盈餘（用當期加權平均股數算的）
     兩者是**同一期的口徑**，相乘反推淨利本來就對得上（台積電 2026Q2 實測
     7,065.6 億 ÷ 259.32 億股 = 27.25，與 EPS 欄位完全吻合）。
     配股造成的稀釋只影響「跨期比較」（EPS YoY、股價漲跌），那是 Calc 層的事，
     用 data/raw_div.csv 的除權息事件算還原係數即可。Raw 層一律存原始值。

  4. 多存一張**除權息事件表**。它同時支撐還原股價與 EPS 還原 —— 還原股價是
     FinMind 的付費功能（實測回 400），而且係數會隨每次新配息而全部改變，
     存成快照下次配息就全錯，所以只存事件、係數每次重算。

另外新增一道免費的正確性檢查：ROE 的兩條路徑（直接抓淨利 vs EPS×股數）算出來
應該一致，不一致就示警 —— 這是坑 A（同名 key 一邊是淨利一邊是權益）的第三道防線。
"""
import os
import sys
import argparse
import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd
import yaml

import sources

ROOT = Path(__file__).parent
DATA = ROOT / "data"
MASTER = DATA / "master.csv"          # 用 CSV 不用 parquet：git 可以 diff，
                                      # 財報被重編時 git log 直接告訴你哪天變的

RAW_Q_COLS = ["ticker", "period", "fy", "fq", "period_end", "is_est", "est_source",
              "eps_basis", "revenue", "gross_profit", "net_income",
              "eps_diluted_adj", "shares_diluted", "treasury_shares",
              "total_equity", "dps", "price_at_end"]
OUT_COLS = RAW_Q_COLS + ["seq", "key"]

MASTER_M = DATA / "master_m.csv"      # 月營收
RAW_M_COLS = ["ticker", "ym", "month_end", "revenue"]
OUT_M_COLS = RAW_M_COLS + ["mseq"]

RAW_DIV_COLS = ["ticker", "ex_date", "cash_dividend", "stock_ratio"]


def log(msg):
    print(msg, flush=True)


# ───────────────────────── seq 樞紐 ─────────────────────────
def add_seq(df):
    """seq：is_est="Y" 記為 0（當季預估）；實際季由新到舊排 1、2、3…
    key：ticker|seq，給 Excel 端做文字型 INDEX/MATCH。"""
    if df.empty:
        out = df.copy()
        out["seq"] = []
        out["key"] = []
        return out[OUT_COLS]
    df = df.copy()
    actual = df.is_est == "N"
    df["seq"] = 0
    df.loc[actual, "seq"] = (df[actual]
                             .groupby("ticker")["period_end"]
                             .rank(ascending=False, method="first")
                             .astype(int))
    df["key"] = df.ticker.astype(str) + "|" + df.seq.astype(str)
    return df[OUT_COLS]


def add_mseq(df):
    """月營收的 seq：由新到舊 1、2、3…（沒有預估列，所以不需要 0）。
    月增率取 mseq 相鄰兩列，年增率取相差 12 的兩列 —— 跟季度表同一個取法。"""
    if df.empty:
        out = df.copy()
        out["mseq"] = []
        return out[OUT_M_COLS]
    df = df.copy()
    df["mseq"] = (df.groupby("ticker")["month_end"]
                    .rank(ascending=False, method="first").astype(int))
    return df[OUT_M_COLS]


# ───────────────────────── 缺口判斷 ─────────────────────────
def latest_reportable(today, lag_days):
    """今天為止，法定期限已過（因此應該已經公布）的最新一季。

    美股是逐檔查 next_earnings；台股統一用法規期限推算：
        Q1~Q3 → 季末 + 45 天
        Q4    → 年末 + 75 天
    再加上 config 的 report_lag_days 當緩衝（台股常在期限前幾天集中公布，
    但也有拖到最後一天的，寬鬆一點不會有壞處）。
    """
    y = today.year
    best = None
    for fy in (y - 1, y):
        for fq in (1, 2, 3, 4):
            due = sources.statutory_deadline(fy, fq) + dt.timedelta(days=lag_days)
            if due <= today:
                best = (fy, fq)
    return best


def plan(ticker, master, cfg, today):
    have = master[master.ticker == ticker] if len(master) else master
    actual = have[have.is_est == "N"] if len(have) else have
    target = cfg["target_quarters"]
    if len(actual) == 0:
        return dict(mode="backfill", quarters=target, why="新標的，抓滿歷史")
    if len(actual) < target:
        return dict(mode="backfill", quarters=target,
                    why=f"季數不足（{len(actual)}/{target}），補歷史")

    want = latest_reportable(today, cfg.get("report_lag_days", 3))
    if want is None:
        return dict(mode="price_only", quarters=0, why="無可判斷的財季")
    newest = actual.sort_values("period_end", ascending=False).iloc[0]
    have_idx = int(newest.fy) * 4 + int(newest.fq)
    want_idx = want[0] * 4 + want[1]
    if have_idx >= want_idx:
        return dict(mode="price_only", quarters=0,
                    why=f"已有 FY{int(newest.fy)}Q{int(newest.fq)}，"
                        f"法定期限內沒有更新的財報")
    return dict(mode="incremental", quarters=cfg["restatement_window"],
                why=f"FY{want[0]}Q{want[1]} 的申報期限已過，抓最近幾季")


# ───────────────────────── 瀑布抓取 ─────────────────────────
def waterfall_fin(ticker, n, chain):
    """依序嘗試各來源，後面的只填前面缺的期別（不覆寫已有值）。"""
    out = sources.empty(sources.FIN_COLS)
    used = []
    for src in chain:
        if getattr(src, "needs_key", False) and not os.getenv("FINMIND_TOKEN"):
            continue
        try:
            df = src.quarterly_financials(ticker, n)
        except Exception as e:
            log(f"      ! {src.name} 財報失敗: {e}")
            continue
        if df is None or df.empty:
            continue
        used.append(src.name)
        if out.empty:
            out = df.copy()
        else:
            # 用 (fy, fq) 當鍵，不能用 period_end ——
            # 不同來源給同一季的期末日會差幾天，用日期當鍵會讓同一季被當成兩季。
            out = out.dropna(subset=["fy", "fq"]).drop_duplicates(subset=["fy", "fq"])
            df = df.dropna(subset=["fy", "fq"]).drop_duplicates(subset=["fy", "fq"])
            out = (out.set_index(["fy", "fq"])
                   .combine_first(df.set_index(["fy", "fq"]))
                   .reset_index())
        need = ["revenue", "net_income", "shares_diluted", "total_equity"]
        if out[need].notna().all().all() and len(out) >= n:
            break
    return out, used


def _merge_eps(a, b):
    if a.empty:
        return b.copy()
    a = a.drop_duplicates(subset=["period_end"]).set_index("period_end")
    b = b.drop_duplicates(subset=["period_end"]).set_index("period_end")
    return a.combine_first(b).reset_index()


def waterfall_eps(ticker, n, chain):
    """跨來源互補 —— 不是拿到第一個非空就停。

    台股沒有街頭 / GAAP 兩套口徑，EPS 就是財報上的基本每股盈餘，
    所以 eps_basis 一律記 "basic"。方法名沿用美股版，讓上層不用改。
    """
    got, used = sources.empty(sources.EPS_COLS), []
    for src in chain:
        if getattr(src, "needs_key", False) and not os.getenv("FINMIND_TOKEN"):
            continue
        try:
            df = src.quarterly_eps_street(ticker, n)
        except Exception as e:
            log(f"      ! {src.name} EPS 失敗: {e}")
            continue
        if df is None or df.empty:
            continue
        used.append(src.name)
        got = _merge_eps(got, df)
        if len(got) >= n:
            break
    return got, "basic", "+".join(used) or None


def is_blank(v):
    """型別安全的「空」判斷。
    不能寫成 `if v:` —— pandas 的 Series / DataFrame 會做逐元素比較，
    丟給 and / if 就會拋 'truth value of a Series is ambiguous'。"""
    if v is None:
        return True
    if isinstance(v, (pd.Series, pd.DataFrame, pd.Index)):
        return len(v) == 0
    if isinstance(v, (dict, list, tuple, set, str)):
        return len(v) == 0
    return False


def first_of(chain, fname, *args):
    for src in chain:
        if getattr(src, "needs_key", False) and not os.getenv("FINMIND_TOKEN"):
            continue
        fn = getattr(src, fname, None)
        if fn is None:
            continue
        try:
            v = fn(*args)
        except Exception as e:
            log(f"      ! {src.name}.{fname} 失敗: {e}")
            continue
        if not is_blank(v):
            return v, src.name
    return None, None


def merged_estimates(chain, ticker):
    """分析師共識 —— 台股覆蓋率低，抓不到是常態，不硬填。

    有標明期別的來源（yq）先填，yf 的 forwardEps 沒標明是哪個財年，只當保底。
    這條規則沿用美股版：讓 yf 先佔走 eps_f1 會導致 f1 == f2、成長率 0、
    PEG_F 全部變 N/M。
    """
    labeled, fallback = {}, {}
    for src in chain:
        fn = getattr(src, "estimates", None)
        if not fn:
            continue
        try:
            got = fn(ticker) or {}
        except Exception:
            continue
        target = fallback if src.name == "yf" else labeled
        for k, v in got.items():
            if v is not None:
                target.setdefault(k, v)
    out = dict(labeled)
    for k, v in fallback.items():
        out.setdefault(k, v)
    if out.get("eps_f1") is None and out.get("eps_f2") is not None:
        out["eps_f1"] = out["eps_f2"]
    if out.get("eps_f2") is None and out.get("eps_f1") is not None:
        out["eps_f2"] = out["eps_f1"]
    return out


# ───────────────────────── 組成 Raw_Q 列 ─────────────────────────
def nearest_price(prices, when):
    if prices is None or len(prices) == 0:
        return np.nan
    idx = pd.to_datetime(pd.Series(prices.index)).dt.date
    ok = idx[idx <= when]
    if ok.empty:
        return np.nan
    return float(prices.iloc[ok.index[-1]])


def align_eps(eps_df, ends):
    """把 EPS 對到財季。

    台股比美股單純：FinMind 的 index 就是期末日（2024-03-31 這種），
    不像 yfinance earnings_dates 那樣是「公布日」。但 yfinance 補位時
    仍可能給出偏移幾天的日期，所以保留 ±12 天的容忍。
    """
    if eps_df is None or eps_df.empty:
        return {}
    asc = sorted(ends)
    out = {}
    for _, r in eps_df.iterrows():
        d, v = r["period_end"], r["eps"]
        if pd.isna(v) or d is None:
            continue
        near = [pe for pe in asc if abs((d - pe).days) <= 12]
        if not near:
            continue
        pe = min(near, key=lambda p: abs((d - p).days))
        gap = abs((d - pe).days)
        if pe not in out or gap < out[pe][1]:
            out[pe] = (float(v), gap)
    return {k: v[0] for k, v in out.items()}


def build_rows(ticker, fin, eps, basis, divs, prices, cfg):
    # combine_first 合併不同來源後，缺漏的期別會讓 fy / fq 變 NaN，int() 會炸。
    fin = fin.dropna(subset=["period_end", "fy", "fq"])
    fin = fin.sort_values("period_end", ascending=False)
    fin = fin.head(cfg["target_quarters"]).copy()
    ends = sorted(fin["period_end"].tolist(), reverse=True)
    epsmap = align_eps(eps, ends)

    rows = []
    for _, r in fin.iterrows():
        pe = r["period_end"]
        prev = next((d for d in ends if d < pe), None)
        dps = np.nan
        if divs is not None and len(divs) and prev is not None:
            d_idx = pd.to_datetime(pd.Series(divs.index)).dt.date
            sel = divs[(d_idx > prev).values & (d_idx <= pe).values]
            dps = float(sel.sum()) if len(sel) else 0.0
        rows.append(dict(
            ticker=ticker, period=f"FY{int(r.fy)}Q{int(r.fq)}",
            fy=int(r.fy), fq=int(r.fq), period_end=pe,
            is_est="N", est_source="actual", eps_basis=basis,
            revenue=r.revenue, gross_profit=r.gross_profit,
            net_income=r.net_income,
            eps_diluted_adj=epsmap.get(pe, np.nan),
            shares_diluted=r.shares_diluted,
            treasury_shares=r.treasury_shares,
            total_equity=r.total_equity, dps=dps,
            price_at_end=nearest_price(prices, pe)))
    return pd.DataFrame(rows, columns=RAW_Q_COLS)


def estimate_row(ticker, actual_rows, est, basis):
    """當季（尚未公布）的預估列 —— seq 會被算成 0。

    期末日用 quarter_end() 精確算，不用美股版的 `+91 天` 近似 ——
    台股是曆年制，季末日是確定的。
    """
    if actual_rows.empty or not est.get("eps_q0"):
        return sources.empty(RAW_Q_COLS)
    ok = actual_rows.dropna(subset=["fy", "fq"])
    if ok.empty:
        return sources.empty(RAW_Q_COLS)
    last = ok.sort_values("period_end", ascending=False).iloc[0]
    fy, fq = int(last.fy), int(last.fq) + 1
    if fq > 4:
        fy, fq = fy + 1, 1
    return pd.DataFrame([dict(
        ticker=ticker, period=f"FY{fy}Q{fq}E", fy=fy, fq=fq,
        period_end=sources.quarter_end(fy, fq),
        is_est="Y", est_source="consensus", eps_basis=basis,
        revenue=np.nan, gross_profit=np.nan, net_income=np.nan,
        eps_diluted_adj=float(est["eps_q0"]), shares_diluted=np.nan,
        treasury_shares=np.nan, total_equity=np.nan, dps=np.nan,
        price_at_end=np.nan)], columns=RAW_Q_COLS)


# ───────────────────────── 清理（不是全有全無的驗證）─────────────────────────
def cross_check_roe(df):
    """ROE 兩條路徑的交叉驗算 —— 坑 A 的第三道防線。

    分子有兩種算法：
      · 直接抓「淨利歸屬母公司」
      · EPS × 期末股數（股本 ÷ 10）
    台股沒有街頭 / GAAP 兩套帳，這兩條應該得到幾乎相同的數字
    （台積電 2026Q2 實測完全吻合）。差太多就代表某一欄接錯了口徑，
    最可能的就是把資產負債表的「權益」當成損益表的「淨利」。

    只示警不修改 —— 這裡的目的是讓錯誤浮上檯面，不是靜靜地修掉它。
    """
    a = df[(df.is_est == "N")].copy()
    for c in ("net_income", "eps_diluted_adj", "shares_diluted"):
        a[c] = pd.to_numeric(a[c], errors="coerce")
    ok = a.dropna(subset=["net_income", "eps_diluted_adj", "shares_diluted"])
    ok = ok[(ok.shares_diluted > 0) & (ok.net_income.abs() > 0)]
    if ok.empty:
        return []
    implied = ok.eps_diluted_adj * ok.shares_diluted
    diff = (implied - ok.net_income).abs() / ok.net_income.abs()
    bad = ok[diff > 0.05]
    notes = []
    for tk, g in bad.groupby("ticker"):
        notes.append(f"{tk} 有 {len(g)} 季的「EPS×股數」與淨利差距超過 5%"
                     f"（最舊 {g.period.iloc[-1]}）—— 請確認淨利欄位沒有接到權益")
    return notes


def sanitize(df):
    """把「不可用」和「會算錯」的列處理掉，而不是把整批更新擋下來。

    三種問題，三種處理：
      · 營收缺失或非正數 → 那一列什麼指標都撐不起來，直接丟掉
      · 季度有缺口       → 只留「最新的連續一段」。所有指標都是用 seq
                          （由新到舊的名次）去取區間，中間缺一季會讓 seq 5
                          不再是去年同期，YoY 會靜靜地算錯
      · 同一財季有兩列   → 期別換算出錯，會汙染 seq。這檔整個退回舊資料
    """
    notes, quarantine = [], set()
    est = df[df.is_est == "Y"]
    act = df[df.is_est == "N"].copy()

    act["revenue"] = pd.to_numeric(act["revenue"], errors="coerce")
    bad = act[act.revenue.isna() | (act.revenue <= 0)]
    if len(bad):
        for r in bad.itertuples():
            notes.append(f"{r.ticker} {r.period} 沒有營收，丟掉這一列")
        act = act.drop(bad.index)

    keep = []
    for tk, g in act.groupby("ticker"):
        g = g.sort_values("period_end", ascending=False)
        idx = (g.fy.astype(int) * 4 + g.fq.astype(int)).tolist()
        if len(set(idx)) != len(idx):
            quarantine.add(tk)
            notes.append(f"{tk} 有兩筆對到同一個財季，這次的資料整批退回")
            continue
        cut = len(g)
        for i in range(1, len(idx)):
            if idx[i - 1] - idx[i] != 1:      # 由新到舊，正常是每次減 1
                cut = i
                break
        if cut < len(g):
            notes.append(f"{tk} 第 {cut + 1} 季往前有缺口，只保留最近 {cut} 季"
                         f"（保留缺口前的資料會讓 YoY 對錯季）")
        keep.append(g.head(cut))
    act = pd.concat(keep, ignore_index=True) if keep else act.iloc[0:0]

    out = pd.concat([act, est[~est.ticker.isin(quarantine)]], ignore_index=True)

    a = out[out.is_est == "N"]
    n_miss = int(pd.to_numeric(a.eps_diluted_adj, errors="coerce").isna().sum())
    if n_miss:
        notes.append(f"{n_miss}/{len(a)} 季沒有對到 EPS")
    n_ni = int(pd.to_numeric(a.net_income, errors="coerce").isna().sum())
    if n_ni:
        notes.append(f"{n_ni}/{len(a)} 季沒有淨利，ROE 會退到 EPS÷BPS 的備案口徑")
    notes += cross_check_roe(out)
    return out, notes, quarantine


# ───────────────────────── 主流程 ─────────────────────────
def keep_others(new_rows, path):
    """--only 只跑部分股票時，沒跑到的那些要沿用舊資料，不能被整份蓋掉。"""
    new = pd.DataFrame(new_rows)
    if path.exists():
        old = pd.read_csv(path)
        if len(new) and "ticker" in old.columns and "ticker" in new.columns:
            old = old[~old.ticker.astype(str).isin(new.ticker.astype(str))]
        new = pd.concat([new, old], ignore_index=True)
    return new.sort_values("ticker") if "ticker" in new.columns else new


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", help="只跑指定代號（逗號分隔），用於除錯")
    ap.add_argument("--force-full", action="store_true", help="忽略快取，全部重抓")
    ap.add_argument("--dry-run", action="store_true", help="不寫檔，只印出計畫")
    ap.add_argument("--month-only", action="store_true",
                    help="只更新月營收（給每月 11 日那支 workflow 用）")
    args = ap.parse_args()

    cfg = yaml.safe_load((ROOT / "config.yaml").read_text(encoding="utf-8"))
    chain = sources.chain(cfg["sources"])
    today = dt.date.today()
    DATA.mkdir(exist_ok=True)

    tickers = pd.read_csv(ROOT / "tickers.csv", dtype={"ticker": str})
    tickers = tickers[tickers.ticker.notna()]
    tickers["ticker"] = tickers.ticker.astype(str).str.strip()
    tickers = tickers[tickers.ticker.str.fullmatch(r"[1-9]\d{3}")]
    if args.only:
        keep = {s.strip() for s in args.only.split(",")}
        tickers = tickers[tickers.ticker.isin(keep)]
    if tickers.empty:
        log("✗ tickers.csv 沒有有效的台股代號（上市上櫃普通股為四位數字）")
        return 1

    if MASTER.exists() and not args.force_full:
        master = pd.read_csv(MASTER, parse_dates=["period_end"],
                             dtype={"ticker": str})
        master["period_end"] = master["period_end"].dt.date
    else:
        master = pd.DataFrame(columns=RAW_Q_COLS)

    # ── 月營收（獨立的一條線，跟季報不同節奏）──────────────
    m_rows = []
    for tk in tickers.ticker if not args.dry_run else []:
        mr, msrc = first_of(chain, "month_revenue", tk,
                            cfg.get("month_revenue_months", 36))
        if mr is None or len(mr) == 0:
            log(f"  {tk:<6} 月營收 抓不到")
            continue
        mr = mr.copy()
        mr["ticker"] = tk
        mr["ym"] = pd.to_datetime(mr.month_end).dt.strftime("%Y-%m")
        m_rows.append(mr[RAW_M_COLS])
        log(f"  {tk:<6} 月營收 {len(mr)} 個月（{msrc}）")

    if not args.dry_run:
        m_new = pd.concat(m_rows, ignore_index=True) if m_rows \
            else pd.DataFrame(columns=RAW_M_COLS)
        if MASTER_M.exists():
            m_old = pd.read_csv(MASTER_M, dtype={"ticker": str})
            m_old = m_old[~m_old.ticker.isin(set(m_new.ticker))] if len(m_new) \
                else m_old
            m_new = pd.concat([m_new, m_old], ignore_index=True)
        m_new = (m_new.drop_duplicates(subset=["ticker", "ym"], keep="first")
                      .sort_values(["ticker", "ym"], ascending=[True, False]))
        m_new.to_csv(MASTER_M, index=False)
        add_mseq(m_new).to_csv(DATA / "raw_m.csv", index=False)
        log(f"  ✓ 月營收：{m_new.ticker.nunique()} 檔 / {len(m_new)} 個月")

    if args.month_only:
        if not args.dry_run:
            pd.read_csv(ROOT / "tickers.csv", dtype={"ticker": str}) \
              .to_csv(DATA / "tickers.csv", index=False)
        return 0

    # ── 季度 ──────────────────────────────────────────────
    est_rows, price_rows, div_rows, all_new = [], [], [], []

    for tk in tickers.ticker:
        p = plan(tk, master, cfg, today)
        log(f"  {tk:<6} {p['mode']:<12} {p['why']}")
        if args.dry_run:
            continue

        est = merged_estimates(chain, tk)
        prices, _ = first_of(chain, "price_history", tk, None)
        px = float(prices.iloc[-1]) if prices is not None and len(prices) else np.nan
        price_rows.append(dict(ticker=tk, price=px, price_date=today, as_of=today))
        est_rows.append(dict(ticker=tk, eps_f1=est.get("eps_f1"),
                             eps_f2=est.get("eps_f2"), fy1_end=est.get("fy1_end"),
                             n_analysts=est.get("n_analysts"), as_of=today))

        ev, _ = first_of(chain, "dividend_events", tk, None)
        if ev is not None and len(ev):
            ev = ev.copy()
            ev["ticker"] = tk
            div_rows.append(ev[RAW_DIV_COLS])

        if p["mode"] == "price_only":
            continue

        fin, used = waterfall_fin(tk, p["quarters"], chain)
        if fin.empty:
            log(f"      ! {tk} 完全抓不到季度財報，跳過")
            continue
        eps, basis, esrc = waterfall_eps(tk, p["quarters"], chain)
        divs, dsrc = first_of(chain, "dividends", tk, None)
        rows = build_rows(tk, fin, eps, basis, divs, prices, cfg)
        n_eps = int(pd.to_numeric(rows.eps_diluted_adj,
                                  errors="coerce").notna().sum())
        log(f"      財報={'+'.join(used)} {len(rows)} 季 ｜ EPS={esrc} "
            f"對上 {n_eps}/{len(rows)} ｜ 除權息 {0 if ev is None else len(ev)} 筆")
        rows = pd.concat([rows, estimate_row(tk, rows, est, basis)],
                         ignore_index=True)
        all_new.append(rows)

    if args.dry_run:
        log("\n(dry-run，未寫檔)")
        return 0

    if all_new:
        new = pd.concat(all_new, ignore_index=True)
        keys = ["ticker", "fy", "fq"]
        # 新的原則上覆蓋舊的（財報會被追溯重編），但有一個例外：
        # 這次抓回來是空的、master 裡原本卻有值，那是來源當下的一次性缺漏，
        # 不能讓它把好資料洗掉。這種洞會一路傳染 —— 少一季就變成缺口，
        # 缺口前面的歷史又會被整段截掉。
        if len(master):
            mrev = pd.to_numeric(master.revenue, errors="coerce")
            usable = mrev.notna() & (mrev > 0)
            have_good = set(map(tuple, master.loc[usable, keys].values))
            nrev = pd.to_numeric(new.revenue, errors="coerce")
            hollow = nrev.isna() | (nrev <= 0)
            clash = pd.Series([tuple(v) in have_good for v in new[keys].values],
                              index=new.index)
            drop = new[hollow & clash]
            if len(drop):
                for r in drop.itertuples():
                    log(f"      ⚠ {r.ticker} {r.period} 這次抓到空的，沿用既有資料")
                new = new.drop(drop.index)
        master = (pd.concat([new, master], ignore_index=True)
                  .drop_duplicates(subset=keys, keep="first"))

    master = master.sort_values(["ticker", "period_end"], ascending=[True, False])
    before = len(master)
    prev = pd.read_csv(MASTER, dtype={"ticker": str}) if MASTER.exists() else None

    master, notes, quarantine = sanitize(master)
    for n in notes:
        log(f"  ⚠ {n}")
    if quarantine and prev is not None:
        old = prev[prev.ticker.isin(quarantine)].copy()
        if len(old):
            old["period_end"] = pd.to_datetime(old["period_end"]).dt.date
            master = pd.concat([master, old], ignore_index=True)
            log(f"  ⚠ {'、'.join(sorted(quarantine))} 沿用上一版資料")
    if len(master) < before:
        log(f"  ⚠ 共清掉 {before - len(master)} 列有問題的資料")
    if master.empty:
        log("\n✗ 清理後沒有任何資料可寫，master 保持原狀")
        return 1

    master = master.sort_values(["ticker", "period_end"], ascending=[True, False])
    master.to_csv(MASTER, index=False)
    add_seq(master).to_csv(DATA / "raw_q.csv", index=False)
    keep_others(est_rows, DATA / "raw_est.csv").to_csv(DATA / "raw_est.csv",
                                                       index=False)
    keep_others(price_rows, DATA / "raw_price.csv").to_csv(DATA / "raw_price.csv",
                                                           index=False)
    div = pd.concat(div_rows, ignore_index=True) if div_rows \
        else pd.DataFrame(columns=RAW_DIV_COLS)
    keep_others(div, DATA / "raw_div.csv").to_csv(DATA / "raw_div.csv", index=False)
    # tickers.csv 要寫「完整的那份」，不是被 --only 篩過的
    pd.read_csv(ROOT / "tickers.csv", dtype={"ticker": str}) \
      .to_csv(DATA / "tickers.csv", index=False)

    log(f"\n✓ 完成：{master.ticker.nunique()} 檔 / {len(master)} 列")
    return 0


if __name__ == "__main__":
    sys.exit(main())
