#!/usr/bin/env python3
"""自動補上公司名稱與產業分類（台股版）。

跟美股版最大的不同：**不用關鍵字猜**。

美股版是拿 yfinance 的英文 GICS 字串（Semiconductors、Software - Infrastructure）
用關鍵字規則對回中文，順序寫錯就會分類錯（"Semiconductor Equipment" 要排在
"Semiconductor" 前面，否則永遠對到後者）。台股不需要這套 —— 證交所／櫃買中心
本來就有官方的產業別，FinMind 的 TaiwanStockInfo 直接帶 industry_category 欄位，
拿到什麼就是什麼。

兩個原則沿用美股版：
1. 只填「空白」或「未分類」的欄位 —— 你自己設過的一律不動。
2. 對不到就留「未分類」，不硬猜。

★ 兩個踩過的坑 ★
  · pandas 讀 CSV 時，整欄空白會被判定成 float64（全是 NaN），往裡面塞字串會拋
    TypeError: Invalid value 'xxx' for dtype 'float64'。所以讀進來要先統一轉成
    字串欄位、NaN 變空字串。
  · 分類只是錦上添花，不該因為它讓整趟同步白跑。所以 main() 一律回 0，
    出錯只印警告。

用法：
    python classify.py            # 只補空的
    python classify.py --force    # 全部重新分類（會蓋掉你手動設的，慎用）
"""
import argparse
import os
import sys
import time
from pathlib import Path

import pandas as pd
import requests

ROOT = Path(__file__).parent
TICKERS = ROOT / "tickers.csv"
BLANK = ("", "未分類", "nan", "None", "NaN")

# tickers.csv 用英文欄名（build_xlsx.py 是用 t.get("company") 取值的）。
# 改成中文欄名會讓 Excel 整片空白。
TEXT_COLS = ["ticker", "company", "sector", "tier", "added", "note", "tags"]

FINMIND = "https://api.finmindtrade.com/api/v4/data"
TWSE_LISTED = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"


def is_blank(v):
    return str(v).strip() in BLANK


def as_text(df):
    """★ float64 陷阱 ★ 整欄空白的 CSV 欄位會被讀成 float64，
    往裡面塞字串就炸。統一轉成字串欄位、NaN 變空字串。"""
    for c in TEXT_COLS:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].astype("object").where(df[c].notna(), "").astype(str)
        df[c] = df[c].replace({"nan": "", "None": "", "NaN": ""})
    return df


def from_finmind():
    """FinMind TaiwanStockInfo → {代號: (名稱, 產業別)}。

    這張表是全市場總覽，欄位有 stock_id / stock_name / industry_category / type。
    免費層對這種小表通常可以不帶 data_id 一次抓全部；抓不到就回空，
    讓下一層接手。
    """
    headers = {}
    token = os.getenv("FINMIND_TOKEN", "").strip()
    if token:
        headers["Authorization"] = f"Bearer {token}"
    try:
        r = requests.get(FINMIND, headers=headers,
                         params={"dataset": "TaiwanStockInfo"}, timeout=30)
        if r.status_code != 200:
            print(f"⚠ FinMind TaiwanStockInfo → HTTP {r.status_code}")
            return {}
        df = pd.DataFrame(r.json().get("data", []))
    except Exception as e:
        print(f"⚠ FinMind TaiwanStockInfo 連線失敗：{e}")
        return {}
    if df.empty or "stock_id" not in df.columns:
        return {}
    out = {}
    for r in df.to_dict("records"):
        sid = str(r.get("stock_id", "")).strip()
        if not sid:
            continue
        out.setdefault(sid, (str(r.get("stock_name", "") or "").strip(),
                             str(r.get("industry_category", "") or "").strip()))
    return out


def finmind_one(sid):
    """單檔查詢 —— 全市場總表抓不到某檔時的補位。

    免費層對部分 dataset 會要求一定要帶 data_id（不帶就回 400），
    所以總表可能整個回空。逐檔查最多就是幾次呼叫，成本很低。
    """
    headers = {}
    token = os.getenv("FINMIND_TOKEN", "").strip()
    if token:
        headers["Authorization"] = f"Bearer {token}"
    try:
        r = requests.get(FINMIND, headers=headers,
                         params={"dataset": "TaiwanStockInfo", "data_id": sid},
                         timeout=20)
        if r.status_code != 200:
            return None
        rows = r.json().get("data", [])
    except Exception:
        return None
    for x in rows:
        name = str(x.get("stock_name", "") or "").strip()
        ind = str(x.get("industry_category", "") or "").strip()
        if name or ind:
            return name, ind
    return None


def from_twse():
    """證交所 OpenAPI 的公司基本資料（上市）。FinMind 抓不到時的備援。

    只有上市；上櫃抓不到就留「未分類」，不硬猜 —— 寧可空著讓你自己填，
    也不要塞一個看起來合理但錯的產業別。
    """
    try:
        r = requests.get(TWSE_LISTED, timeout=30,
                         headers={"accept": "application/json"})
        if r.status_code != 200:
            print(f"⚠ 證交所 OpenAPI → HTTP {r.status_code}")
            return {}
        rows = r.json()
    except Exception as e:
        print(f"⚠ 證交所 OpenAPI 連線失敗：{e}")
        return {}
    out = {}
    for r in rows if isinstance(rows, list) else []:
        sid = str(r.get("公司代號", "")).strip()
        if not sid:
            continue
        out[sid] = (str(r.get("公司簡稱", "") or "").strip(),
                    str(r.get("產業別", "") or "").strip())
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true",
                    help="全部重新分類（會蓋掉手動設定，慎用）")
    args = ap.parse_args()

    if not TICKERS.exists():
        print("⚠ 找不到 tickers.csv，略過分類")
        return 0

    df = as_text(pd.read_csv(TICKERS, dtype=str))
    lookup = from_finmind()
    src = "FinMind 總表"
    if not lookup:
        lookup = from_twse()
        src = "證交所 OpenAPI"
    if lookup:
        print(f"✓ 產業別來源：{src}（{len(lookup)} 檔）")
    else:
        print("⚠ 總表抓不到，改成逐檔查 FinMind")

    # 總表沒有的（上櫃、或總表整個回空）逐檔補
    for sid in df["ticker"].astype(str).str.strip():
        if sid and sid not in lookup:
            hit = finmind_one(sid)
            if hit:
                lookup[sid] = hit
            time.sleep(0.3)
    if not lookup:
        print("⚠ 完全抓不到公司基本資料，略過分類（不影響同步）")
        return 0

    n_name = n_sector = 0
    unknown = []
    for i, row in df.iterrows():
        sid = str(row.get("ticker", "")).strip()
        if not sid:
            continue
        hit = lookup.get(sid)
        if not hit:
            unknown.append(sid)
            continue
        name, sector = hit
        if name and (args.force or is_blank(row.get("company"))):
            df.at[i, "company"] = name
            n_name += 1
        if sector and (args.force or is_blank(row.get("sector"))):
            df.at[i, "sector"] = sector
            n_sector += 1
        if is_blank(df.at[i, "sector"]):
            df.at[i, "sector"] = "未分類"
        if is_blank(df.at[i, "tier"]):
            df.at[i, "tier"] = "池子"

    df.to_csv(TICKERS, index=False)
    # ★ 一定要同時更新 data/tickers.csv ★
    # workflow 的順序是 sync → classify → export_json。sync 在結束時已經把
    # 「還沒分類的」根目錄 tickers.csv 複製到 data/ 了，而 export_json.py 與
    # build_xlsx.py 讀的是 data/ 那一份。只寫根目錄的話，剛補好的公司名稱與
    # 產業別要等到下一輪同步才會出現在網頁上 —— 看起來就像 classify 沒生效。
    data_copy = ROOT / "data" / "tickers.csv"
    if data_copy.parent.exists():
        df.to_csv(data_copy, index=False)
    print(f"✓ 補上公司名稱 {n_name} 筆、產業別 {n_sector} 筆")
    if unknown:
        print(f"⚠ 查無基本資料（留未分類，可自行填）：{'、'.join(unknown[:20])}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        # 分類失敗不該讓整趟同步白跑 —— 前面可能已經花了一分多鐘抓資料。
        print(f"⚠ classify 發生非預期錯誤，已略過：{type(e).__name__}: {e}")
        sys.exit(0)
