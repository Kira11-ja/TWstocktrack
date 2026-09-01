# 台股觀察表

跟 [美股版](https://github.com/Kira11-ja/stocktracker) 同一套架構，抓資料那一層整個換掉。

三層，資料往下長不往右長：

- **Raw 層**（長格式 CSV）— 一列一個「股票 × 財季」，只放原始欄位
- **Calc 層**（Excel 隱藏分頁 / 網頁的 JS）— 所有衍生指標
- **呈現層** — Excel 的 Dashboard／個股卡，以及 GitHub Pages 網頁

`seq` 是整套的樞紐：`is_est="Y"` 記 0（當季預估），實際季由新到舊排 1、2、3…
所有 TTM／YoY／QoQ 都是用 seq 區間去 SUMIFS，新增股票不必改任何公式。

---

## 第一次啟用

1. 建一個新的 repo，把這些檔案全部推上去
2. **Settings → Secrets and variables → Actions** 新增 `FINMIND_TOKEN`
   （到 [finmindtrade.com](https://finmindtrade.com) 註冊免費帳號取得。
   不設也能跑，只是額度從每小時 600 次降到 300 次）
3. **Settings → Pages** 來源選 `main` 分支的 `/docs` 目錄
4. **Actions** 分頁手動跑一次 `update-stock-data`，勾 `force_full`

跑完 `data/`、`docs/data.json` 和 `台股觀察表.xlsx` 就會出現在 repo 裡。

---

## 資料來源

FinMind 一家補滿 Raw 層（實測 2330／2408／5289 都是 2019-03 起 30 季、
月營收 92 個月零缺漏），yfinance 與 yahooquery 只補位。

| Raw 欄位 | 來源 | FinMind type |
|---|---|---|
| revenue | 綜合損益表 | `Revenue` |
| gross_profit | 綜合損益表 | `GrossProfit` |
| net_income | 綜合損益表 | `EquityAttributableToOwnersOfParent` |
| eps_diluted_adj | 綜合損益表 | `EPS`（台股是基本每股盈餘） |
| total_equity | 資產負債表 | `EquityAttributableToOwnersOfParent` |
| shares_diluted | 資產負債表 | `OrdinaryShare` ÷ 10（面額固定 10 元） |
| treasury_shares | 資產負債表 | `NumberOfSharesInEntityHeldByEntityAndByItsSubsidiaries` |
| price_at_end | 股價 | `TaiwanStockPrice`（原始收盤，未還原） |
| 月營收 | 月營收 | `TaiwanStockMonthRevenue` |
| 除權息事件 | 股利政策 | `TaiwanStockDividend` |
| 共識預估 | yahooquery | `earnings_trend`（覆蓋率低，抓不到就 N/M） |
| 產業別 | FinMind / 證交所 | 官方產業別，不用關鍵字猜 |

`TaiwanStockPriceAdj`（還原股價）是**付費方案**，所以 Raw 層存原始收盤價，
還原係數在 Calc 層用除權息事件表每次重算 —— 係數是相對於「今天」算的，
存成快照下次配息就全錯了。

---

## ★ 三個會靜靜算錯的坑 ★

這三個都不會拋例外、不會回空表，只會給出一個看起來很合理的錯數字。

**坑 A —— 同一個 key 在兩個 dataset 意義完全不同（差 900 倍）**

`EquityAttributableToOwnersOfParent`：
- 在**損益表**是「淨利（淨損）歸屬於母公司業主」= **淨利**
- 在**資產負債表**是「歸屬於母公司業主之權益合計」= **權益**

台積電 2026Q2 分別是 7,066 億與 6.43 兆，剛好就是 ROE 的分子和分母。
防線有三道：兩個 dataset 分開 pivot 各取各的白名單、`_sanity_net_income()`
用「淨利不可能是營收的 1.5 倍」擋一次、`sync.py` 的 `cross_check_roe()`
比對「EPS×股數」與淨利差距超過 5% 就示警。

**坑 B —— 資產負債表每欄都有 `_per` 影子欄位（差 100 倍）**

```
Equity      權益總額  6.474471e+12
Equity_per  權益總額  6.906000e+01   ← 佔資產總額的百分比
```

兩者的 `origin_name` **完全相同**，用中文欄名比對會抓到哪一筆看順序決定。
防線：`_FM.wide()` 一律先剔除 `type` 結尾為 `_per` 的列。

**坑 C —— FinMind 是單季制，不要再 diff**

實測 2330 的 2024 四季營收加總 = 2.894 兆 = 全年營收；EPS 加總 = 45.26 = 全年 EPS。
資料已經是單季，再 `groupby(year).diff()` 一次會望遠鏡式全部抵銷，
TTM 只剩最新一季，約為真值的 1/3。

---

## 沿用美股版踩過的坑

**GitHub Actions**

1. `bash -e` 底下 `[ test ] && cmd`，test 為假時整個步驟會以 1 結束。要用 `if…fi`
2. 步驟寫了自訂的 `if:` 會蓋掉預設的 `success()`。要寫 `if: success() && …`
3. `concurrency` 同一組只允許「一個在跑 + 一個在等」，第三個會被取消
4. 推送可能被拒（期間 main 前進了）。要先合併再重推，最多三次
5. Issue 標題是任何人都能打的字串，**絕不可以內插進 shell**
6. 寫進 `GITHUB_ENV` 的值一定要壓成單行

**資料處理**

7. pandas 的 Series 不能直接丟給 `if` 或 `and`，要寫型別安全的空值判斷
8. 跨來源合併要用 `(fy, fq)` 當鍵，**不能用期末日**
9. 瀑布式來源**不能拿到第一個非空就停**，要跨來源互補
10. master 用 CSV 不用 parquet —— 要讓 git 能 diff，才看得出財報被追溯重編
11. 整欄空白的 CSV 欄位會被 pandas 讀成 `float64`，塞字串會拋
    `TypeError: Invalid value ... for dtype 'float64'`。讀進來先統一轉字串

**Excel**

12. openpyxl 只寫公式字串、不寫快取值。線上預覽會整片空白，
    要在 CI 用 LibreOffice 算過一次再交付
13. LibreOffice 不支援 `XLOOKUP`／`FILTER`／`LET`／`SORT`／`UNIQUE`／`SEQUENCE`

**網頁**

14. `<input list=datalist>` 會用欄位既有的字過濾選項，要用真的 `<select>`
15. 同一個 `id` 重複時 `getElementById` 只認第一個

**其他**

16. `tickers.csv` 的欄名是**英文鍵**（`ticker,company,sector,tier,added,note,tags`），
    `build_xlsx.py` 是用 `t.get("company")` 取值的。改成中文欄名會讓 Excel 整片空白

---

## 台股跟美股不同的地方

- **股數不需要還原。** 美股要猜分割倍率，是因為 yfinance 的股數沒有回溯調整、
  EPS 卻是調整後的。台股沒這問題：股數是當期期末股本÷10、EPS 是當期財報值，
  同一期的口徑。台積電 2026Q2 實測 7,065.6 億 ÷ 259.32 億股 = 27.25，
  與 EPS 欄位完全吻合。配股的稀釋只影響跨期比較，那是 Calc 層的事。
- **沒有逐檔的財報預告日。** 只有法定期限：Q1~Q3 季後 45 天、Q4 年報 75 天。
  `plan()` 用 `statutory_deadline()` 推算，取代美股的 `meta.next_earnings`。
- **EPS 只有一套口徑。** 台股沒有街頭 vs GAAP 的分歧，`eps_basis` 一律 `basic`。
- **月營收。** 獨立的一張表、獨立的一支 workflow（每月 11、12、13 日各跑一次）。
- **收盤時間。** 13:30 收盤，排程是 UTC 06:30（台灣 14:30）。

---

## 還沒做的

- **金融股**：`sources.py` 的 `FINANCIAL_REVENUE_T` 刻意留白。金控的損益表沒有
  「營業收入」「營業毛利」，`sanitize` 那條「營收缺失就丟掉該列」會把整檔清空。
  拿一檔金控（例如 2882）跑 `sources.FM.dump_types("2882")`，把「淨收益」那一欄的
  type 名稱填進去 —— 不憑印象猜字串。
- **Calc 層的 ROE 主路徑**：`net_income` 已經進 Raw 層，但 Excel 與 `metrics.js`
  目前仍用「EPS×股數」反推。台股這兩條會得到同一個數字（已驗證），所以不急，
  但改成直接用 `net_income` 會更穩。
- **還原係數**：`data/raw_div.csv` 已經在存除權息事件，但 Calc 層還沒用它算
  還原股價與 EPS 還原。

---

## 手動跑

```bash
pip install -r requirements.txt

python sync.py --dry-run          # 只印計畫，不抓不寫
python sync.py --only 2330        # 只跑一檔
python sync.py --force-full       # 忽略快取全部重抓
python sync.py --month-only       # 只更新月營收

python classify.py                # 補公司名稱與產業別
python export_json.py             # 產生 docs/data.json
python build_xlsx.py              # 產生 台股觀察表.xlsx
```
