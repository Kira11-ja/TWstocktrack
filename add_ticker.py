#!/usr/bin/env python3
"""從 GitHub Issue 新增／移除股票、或修改它的等級與備註。

網頁上的「＋新增股票」會開一張標題是 `add: NVDA` 的 Issue，
這支程式負責把它變成 tickers.csv 的一列。
移除是 `remove: NVDA`，改等級／備註是 `edit: NVDA`（新值放在 Issue 內文）。

刻意直接讀 GitHub 給的事件 JSON，不從 workflow 把標題內插進 shell——
Issue 標題是任何人都能打的字串，內插進 shell 等於開後門。
代號本身也會用白名單正規表達式驗過才使用。
"""
import csv
import datetime as dt
import json
import os
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).parent
DATA = ROOT / "data"
TICKERS = ROOT / "tickers.csv"
# 欄名沿用 repo 原本的英文鍵 —— build_xlsx.py 是用 t.get("company") / t.get("sector")
# 取值的，改成中文欄名會讓 Excel 的 Tickers 分頁整片空白。
COLS = ["ticker", "company", "sector", "tier", "added", "note", "tags"]
TAG_SEP = "、"
TAG_SPLIT = re.compile(r"[、,，;；]+")
TAG_MAX = 6
SYMBOL = re.compile(r"^[1-9]\d{3}$")   # 台股上市上櫃普通股：四位數字，不以 0 開頭（0 開頭是 ETF）
DEFAULT_TIER = "池子"
LABEL_MAX = 16


def tags(v):
    """標籤是多值的，用頓號或逗號分隔。空字串＝清空。"""
    parts = [label(x) for x in TAG_SPLIT.split(str(v or ""))]
    seen, out = set(), []
    for x in parts:
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return TAG_SEP.join(out[:TAG_MAX])


def label(v):
    """等級與產業是自由文字，你想打什麼都行——這裡只做基本清理，不設白名單。
    代號有白名單是因為它會被拿去打 API；標籤只是存進 CSV 給人看。"""
    v = re.sub(r"[\r\n\t]+", " ", str(v or "")).strip()
    return v[:LABEL_MAX]


def emit(**kv):
    """把結果交給後面的 workflow 步驟。

    值一定要壓成單行再寫進 GITHUB_ENV —— 帶換行的值可以偽造出額外的環境變數，
    而這裡的內容有一部分來自 Issue 標題（誰都能打）。
    """
    kv = {k: re.sub(r"[\r\n]+", " ", str(v))[:200] for k, v in kv.items()}
    path = os.getenv("GITHUB_ENV")
    if path:
        with open(path, "a", encoding="utf-8") as f:
            for k, v in kv.items():
                f.write(f"{k}={v}\n")
    for k, v in kv.items():
        print(f"{k}={v}")


def parse_body(body):
    out = {}
    for line in (body or "").splitlines():
        if ":" not in line:
            continue
        k, v = line.split(":", 1)
        k = k.strip().lower()
        if k in ("tier", "name", "note", "industry", "tags"):
            out[k] = v.strip()
    return out


def load_tickers():
    if TICKERS.exists():
        df = pd.read_csv(TICKERS)
    else:
        df = pd.DataFrame(columns=COLS)
    for c in COLS:
        if c not in df.columns:
            df[c] = ""
    return df


def publish(df):
    """tickers.csv 是輸入，data/tickers.csv 是網頁與 Excel 讀的那份。
    只改等級／移除時不會跑 sync.py，所以這裡要自己把兩邊同步。"""
    df.to_csv(TICKERS, index=False, quoting=csv.QUOTE_MINIMAL)
    DATA.mkdir(exist_ok=True)
    df.to_csv(DATA / "tickers.csv", index=False, quoting=csv.QUOTE_MINIMAL)


def strip_from_data(t):
    """移除時，把該檔從所有產出檔一起清掉，不然網頁還會看到它。"""
    for name in ("master.csv", "meta.csv", "raw_q.csv", "raw_est.csv", "raw_price.csv"):
        p = DATA / name
        if not p.exists():
            continue
        df = pd.read_csv(p)
        if "ticker" in df.columns:
            df[df.ticker != t].to_csv(p, index=False)


def main():
    ev_path = os.getenv("GITHUB_EVENT_PATH")
    if not ev_path or not Path(ev_path).exists():
        print("找不到 GitHub 事件檔，這支程式只能在 Actions 裡跑")
        emit(ACTION="none", TICKER="", RESULT="")
        return 0

    ev = json.loads(Path(ev_path).read_text(encoding="utf-8"))
    issue = ev.get("issue") or {}
    title = (issue.get("title") or "").strip()
    body = issue.get("body") or ""

    m = re.match(r"^\s*(add|remove|edit)\s*[:：]\s*(.+)$", title, re.I)
    if not m:
        emit(ACTION="none", TICKER="",
             RESULT="標題不是 add: / remove: / edit: 開頭，略過")
        return 0

    action = m.group(1).lower()
    t = m.group(2).strip().upper()
    if not SYMBOL.match(t):
        emit(ACTION="none", TICKER="",
             RESULT=f"「{t[:20]}」看起來不像美股代號，沒有做任何事")
        return 0

    df = load_tickers()
    have = df.ticker.astype(str).str.upper().tolist()

    if action == "add":
        if t in have:
            emit(ACTION="none", TICKER=t, RESULT=f"{t} 已經在清單裡了，沒有重複加入")
            return 0
        f = parse_body(body)
        row = {
            "ticker": t,
            "company": label(f.get("name")),
            "sector": label(f.get("industry")) or "未分類",
            "tier": label(f.get("tier")) or DEFAULT_TIER,
            "added": dt.date.today().isoformat(),
            "note": (f.get("note") or "").strip()[:200],
            "tags": tags(f.get("tags")),
        }
        publish(pd.concat([df, pd.DataFrame([row])], ignore_index=True))
        emit(ACTION="add", TICKER=t,
             RESULT=f"已加入 {t}（等級 {row['tier']}），正在抓歷史資料")
        return 0

    if t not in have:
        emit(ACTION="none", TICKER=t, RESULT=f"{t} 本來就不在清單裡")
        return 0

    if action == "edit":
        f = parse_body(body)
        hit = df.ticker.astype(str).str.upper() == t
        changed = []
        for key, col, zh in (("tier", "tier", "等級"), ("industry", "sector", "產業")):
            if key not in f:
                continue
            v = label(f[key])
            if not v:                      # 留空＝不動，避免手滑清掉分類
                continue
            df.loc[hit, col] = v
            changed.append(f"{zh} → {v}")
        if "tags" in f:                    # 標籤留空＝清空，跟等級不同
            v = tags(f["tags"])
            df.loc[hit, "tags"] = v
            changed.append(f"標籤 → {v}" if v else "標籤已清空")
        if "note" in f:
            note = f["note"].strip()[:200]
            df.loc[hit, "note"] = note
            changed.append("備註已更新" if note else "備註已清空")
        if not changed:
            emit(ACTION="none", TICKER=t, RESULT=f"{t} 沒有指定要改什麼")
            return 0
        publish(df)
        emit(ACTION="edit", TICKER=t, RESULT=f"{t}：{'、'.join(changed)}")
        return 0

    publish(df[df.ticker.astype(str).str.upper() != t])
    strip_from_data(t)
    emit(ACTION="remove", TICKER=t, RESULT=f"已移除 {t} 與它的歷史資料")
    return 0


if __name__ == "__main__":
    sys.exit(main())
