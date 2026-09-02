/* 指標計算 —— 這裡的每一條都對應 Excel「Calc」分頁的同名欄位，
   刻意連 N/M 的判斷條件都一模一樣，兩邊數字才會對得起來。 */

export const DEFAULTS = {
  months_shown: 18,        // 月營收顯示幾個月（要 >= 13 才看得到完整的年增率）
  peg_min_growth: 0.05,   // 成長率低於此值時 PEG 顯示 N/M（低成長時 PEG 會失真）
  gm_avg_quarters: 20,    // 「5 年平均毛利率」的取樣季數
  min_analysts: 5,        // 分析師家數低於此值 → 預估信心不足，出旗標
  ntm_days: 365,          // EPS_NTM 依 fy1_end 距今天數在 F1 / F2 之間加權
  accel_flat_pp: 0,       // 加速度絕對值小於此值視為「持平」（0＝完全比照 Excel）
  quarters_shown: 12,
  years_shown: 6,
};

/* 產業分類 —— 證交所／櫃買中心的官方產業別，不是 GICS。
   刻意不用 11 大類（太粗，NVDA 和 AAPL 會被歸在一起而沒得比），
   也不用 GICS 最細的 163 個子產業（太細，每檔自成一類就失去比較意義）。
   左邊是給下拉選單分群用的大類，右邊才是實際寫進 tickers.csv 的值。
   你可以自由改成任何字串，篩選器是從 tickers.csv 現有的值長出來的。 */
export const SECTORS = [
  ['半導體', ['半導體業', '半導體']],
  ['電子零組件', ['電子零組件業', '電子零組件']],
  ['電腦及週邊', ['電腦及週邊設備業', '電腦及週邊設備']],
  ['光電', ['光電業', '光電']],
  ['通信網路', ['通信網路業', '通信網路']],
  ['電子通路', ['電子通路業', '電子通路']],
  ['資訊服務', ['資訊服務業', '資訊服務']],
  ['其他電子', ['其他電子業', '電機機械', '電器電纜']],
  ['金融保險', ['金融保險業', '金融保險', '銀行', '保險', '證券']],
  ['傳產原物料', ['塑膠工業', '化學工業', '鋼鐵工業', '水泥工業', '造紙工業',
                  '玻璃陶瓷', '橡膠工業', '紡織纖維']],
  ['民生消費', ['食品工業', '貿易百貨業', '觀光餐旅', '居家生活', '運動休閒']],
  ['生技醫療', ['生技醫療業', '生技醫療', '醫療器材']],
  ['運輸營建', ['航運業', '汽車工業', '建材營造業', '油電燃氣業']],
  ['其他', ['其他業', '文化創意業', '綠能環保', '數位雲端', '未分類']],
];

/* 這兩個只是「建議值」，不是白名單。等級和產業都可以自己打新的，
   下拉選單會把你用過的值一起帶出來。 */
export const TIERS = ['核心', '觀察', '池子'];

export function sectorOf(industry) {
  for (const [sec, list] of SECTORS) if (list.includes(industry)) return sec;
  return '自訂';
}

/** 建議清單 ∪ 資料裡實際出現過的值，後者排在最後面。 */
export function withCustom(base, used) {
  const seen = new Set(base);
  const extra = [...new Set(used)].filter((v) => v && !seen.has(v)).sort();
  return { base, extra };
}

export const NM = 'N/M';
export const isNum = (v) => typeof v === 'number' && isFinite(v);

/* SUMIFS 的語意：找不到列就是 0，不是錯誤 */
function sumRange(rows, field, from, to) {
  let s = 0;
  for (const r of rows) {
    if (r.seq === null || r.seq === undefined) continue;
    if (r.seq >= from && r.seq <= to) {
      const v = r[field];
      if (typeof v === 'number' && isFinite(v)) s += v;
    }
  }
  return s;
}

/* IFERROR 的語意：分母 0 或非數字 → 空值 */
function div(a, b) {
  if (!isFinite(a) || !isFinite(b) || b === 0) return null;
  const r = a / b;
  return isFinite(r) ? r : null;
}
const growth = (a, b) => { const r = div(a, b); return r === null ? null : r - 1; };

function daysBetween(iso, today) {
  if (!iso) return null;
  const d = Date.parse(iso + 'T00:00:00Z');
  if (!isFinite(d)) return null;
  return (d - today) / 86400000;
}

/* 一檔股票的所有季列，seq 由 add_seq() 在 Python 端就算好：
   seq=0 是當季預估，seq=1 是最新一季實際，往回 2、3、4… */
export function indexQuarters(quarters) {
  const by = new Map();
  for (const r of quarters) {
    if (!by.has(r.t)) by.set(r.t, []);
    by.get(r.t).push(r);
  }
  for (const rows of by.values()) rows.sort((a, b) => (a.seq ?? 99) - (b.seq ?? 99));
  return by;
}

export function momentum(qoq, accelPp, flatPp) {
  if (!isNum(qoq) || !isNum(accelPp)) return null;
  if (flatPp > 0 && Math.abs(accelPp) < flatPp) {
    return qoq > 0 ? '成長持平' : '衰退持平';
  }
  if (qoq > 0) return accelPp > 0 ? '成長且加速' : '成長但減速';
  return accelPp > 0 ? '衰退但收斂' : '衰退且惡化';
}

/* ── 總覽用：一檔股票的所有摘要指標 ───────────────────────── */
export function computeSummary(t, rows, est, price, cfg, todayMs) {
  const S = (f, a, b) => sumRange(rows, f, a, b);
  const at = (f, s) => sumRange(rows, f, s, s);   // 單季也走 SUMIFS，缺列＝0
  const o = { t };

  o.price = price?.price ?? null;
  o.next_earnings = price?.next_earnings ?? null;
  o.as_of = price?.as_of ?? null;
  o.n_analysts = est?.n ?? null;

  o.rev_ttm = S('rev', 1, 4);
  o.eps_ttm = S('eps', 1, 4);
  o.eps_ttm_fwd = S('eps', 0, 3);
  o.eps_ttm_yoy = growth(S('eps', 1, 4), S('eps', 5, 8));

  const f1 = est?.f1 ?? null, f2 = est?.f2 ?? null;
  if (isNum(f1) && isNum(f2)) {
    const d = daysBetween(est.fy1_end, todayMs);
    const w = d === null ? 1 : Math.min(1, Math.max(0, d / cfg.ntm_days));
    o.eps_ntm = f1 * w + f2 * (1 - w);
  } else {
    o.eps_ntm = isNum(f1) ? f1 : null;
  }
  o.eps_growth_f = (isNum(f1) && isNum(f2) && f1 !== 0) ? f2 / f1 - 1 : null;

  const pe = (e) => (!isNum(o.price) || !isNum(e) || e <= 0) ? NM : o.price / e;
  o.pe_ttm = pe(o.eps_ttm);
  o.pe_ttm_fwd = pe(o.eps_ttm_fwd);
  o.pe_ntm = pe(o.eps_ntm);

  o.peg_t = (!isNum(o.pe_ttm) || !isNum(o.eps_ttm_yoy) || o.eps_ttm_yoy < cfg.peg_min_growth)
    ? NM : o.pe_ttm / (o.eps_ttm_yoy * 100);
  o.peg_f = (!isNum(o.pe_ntm) || !isNum(o.eps_growth_f) || o.eps_growth_f < cfg.peg_min_growth)
    ? NM : o.pe_ntm / (o.eps_growth_f * 100);

  o.dps_ttm = S('dps', 1, 4);
  o.yield = o.dps_ttm === 0 ? '不配息' : div(o.dps_ttm, o.price);
  const dpsPrev = S('dps', 5, 8);
  o.dps_yoy = dpsPrev === 0 ? null : growth(o.dps_ttm, dpsPrev);

  o.rev_yoy_q = growth(at('rev', 1), at('rev', 5));
  o.rev_qoq = growth(at('rev', 1), at('rev', 2));
  o.rev_qoq_prev = growth(at('rev', 2), at('rev', 3));
  o.accel_pp = (isNum(o.rev_qoq) && isNum(o.rev_qoq_prev))
    ? (o.rev_qoq - o.rev_qoq_prev) * 100 : null;
  o.momentum = momentum(o.rev_qoq, o.accel_pp, cfg.accel_flat_pp);
  const qoqYearAgo = growth(at('rev', 5), at('rev', 6));
  o.yoy_accel_pp = (isNum(o.rev_qoq) && isNum(qoqYearAgo))
    ? (o.rev_qoq - qoqYearAgo) * 100 : null;
  o.eps_yoy_q = growth(at('eps', 1), at('eps', 5));

  o.gm_ttm = div(S('gp', 1, 4), o.rev_ttm);
  o.gm_5y = div(S('gp', 1, cfg.gm_avg_quarters), S('rev', 1, cfg.gm_avg_quarters));
  o.gm_vs_5y_bps = (isNum(o.gm_ttm) && isNum(o.gm_5y)) ? (o.gm_ttm - o.gm_5y) * 10000 : null;
  const gm1 = div(at('gp', 1), at('rev', 1)), gm5 = div(at('gp', 5), at('rev', 5));
  o.dgm_yoy_bps = (isNum(gm1) && isNum(gm5)) ? (gm1 - gm5) * 10000 : null;

  o.bvps = div(at('eq', 1), at('sh', 1));

  /* 淨利用 EPS×股數反推 —— 與 Excel 同口徑，
     這樣 ROE 的分子才跟 EPS_TTM 是同一套帳。

     台股其實有現成的「淨利歸屬母公司」（Raw 層的 ni 欄），而且實測與
     EPS×股數 完全吻合（台積電 2026Q2：7,065.6 億 ÷ 259.32 億股 = 27.25 = EPS）。
     所以這裡沿用反推是安全的；改成直接讀 ni 會更穩，屬於之後的優化。
     美股當初必須反推，是因為街頭 EPS 與 GAAP 淨利不是同一套帳，台股沒這問題。 */
  const niWindow = (a, b) => {
    let s = 0;
    for (const r of rows) {
      if (r.seq >= a && r.seq <= b && isNum(r.eps) && isNum(r.sh)) s += r.eps * r.sh;
    }
    return s;
  };
  const roeWindow = (a) => {
    const eqNew = at('eq', a), eqOld = at('eq', a + 4);
    if (Math.min(eqNew, eqOld) <= 0) return null;
    return div(niWindow(a, a + 3), (eqNew + eqOld) / 2);
  };
  o.ni_ttm = niWindow(1, 4);
  o.roe_ttm = roeWindow(1) ?? NM;
  const anchors = [1, 5, 9, 13, 17].map(roeWindow).filter(isNum);
  o.roe_5y = anchors.length ? anchors.reduce((a, b) => a + b, 0) / anchors.length : null;
  o.roe_fwd = (!isNum(o.bvps) || o.bvps <= 0 || !isNum(f1)) ? NM : f1 / o.bvps;

  const e0 = rows.find((r) => r.seq === 0);
  o.est_eps = e0?.eps ?? null;
  o.est_period = e0?.period ?? null;
  o.est_src = e0?.src ?? null;
  o.eps_basis = rows.find((r) => r.basis)?.basis ?? null;

  o.flags = [];
  if (isNum(o.n_analysts) && o.n_analysts < cfg.min_analysts) o.flags.push('分析師僅 ' + o.n_analysts + ' 家');
  // 台股只有一套 EPS 口徑（財報上的基本每股盈餘），沒有街頭 vs GAAP 的分歧，
  // 所以不需要美股那個「EPS 降級」的旗標。
  o.n_actual = rows.filter((r) => !r.est).length;
  if (o.n_actual === 0) o.flags = ['資料尚未抓取（等 GitHub 跑完這一輪）'];
  else if (o.n_actual < 8) o.flags.push('僅 ' + o.n_actual + ' 季歷史');
  return o;
}

/* ── 季度檢視 ──────────────────────────────────────────── */
/* ── 月營收（台股獨有）─────────────────────────────────
   台股每月 10 日前公布上月營收，是最高頻的基本面指標，
   比季報早兩三個月透露轉折。

   三個衍生指標的定義：
     · 月增率   本月 ÷ 上月 −1。台股淡旺季明顯，單看月增率容易被季節性騙。
     · 年增率   本月 ÷ 去年同月 −1。排除季節性，這是主要看的那個。
     · 累計年增率  今年 1 月到本月的累計 ÷ 去年同期累計 −1。
                  平滑掉單月的出貨遞延（例如 3 月的貨壓到 4 月出），
                  是判斷「趨勢」而非「單月雜訊」用的。

   累計只在「今年到本月」與「去年同期」兩段都完整時才算 —— 缺一個月就寧可留空，
   不然分母少一個月會讓累計年增率無聲地虛高。 */
export function indexMonths(months) {
  const by = new Map();
  for (const r of months || []) {
    if (!by.has(r.t)) by.set(r.t, []);
    by.get(r.t).push(r);
  }
  for (const rows of by.values()) rows.sort((a, b) => (a.mseq ?? 9999) - (b.mseq ?? 9999));
  return by;
}

export function monthTable(rows, cfg) {
  if (!rows || !rows.length) return [];
  const rev = new Map();
  for (const r of rows) if (isNum(r.rev)) rev.set(r.ym, r.rev);
  const ym = (y, m) => `${y}-${String(m).padStart(2, '0')}`;

  const out = [];
  const n = cfg.months_shown ?? 18;
  for (const r of rows.slice(0, n)) {
    const parts = String(r.ym || '').split('-');
    const y = Number(parts[0]), m = Number(parts[1]);
    if (!y || !m) continue;

    const prevYm = m === 1 ? ym(y - 1, 12) : ym(y, m - 1);

    let cum = 0, cumPrev = 0, full = true, fullPrev = true;
    for (let i = 1; i <= m; i++) {
      const a = rev.get(ym(y, i));
      const b = rev.get(ym(y - 1, i));
      if (isNum(a)) cum += a; else full = false;
      if (isNum(b)) cumPrev += b; else fullPrev = false;
    }

    out.push({
      ym: r.ym,
      rev: r.rev,
      mom: growth(r.rev, rev.get(prevYm)),
      yoy: growth(r.rev, rev.get(ym(y - 1, m))),
      cum: full ? cum : null,
      cum_yoy: (full && fullPrev) ? growth(cum, cumPrev) : null,
    });
  }
  return out;
}

export function quarterTable(rows, cfg) {
  const at = (f, s) => sumRange(rows, f, s, s);
  const out = [];
  for (let s = cfg.quarters_shown; s >= 1; s--) {
    const r = rows.find((x) => x.seq === s);
    if (!r) continue;
    const qoq = growth(at('rev', s), at('rev', s + 1));
    const prev = growth(at('rev', s + 1), at('rev', s + 2));
    const accel = (isNum(qoq) && isNum(prev)) ? (qoq - prev) * 100 : null;
    const qoqYearAgo = growth(at('rev', s + 4), at('rev', s + 5));
    const gm = div(at('gp', s), at('rev', s));
    const gmYearAgo = div(at('gp', s + 4), at('rev', s + 4));
    out.push({
      period: r.period, end: r.end, rev: r.rev, qoq, prev_qoq: prev, accel_pp: accel,
      momentum: momentum(qoq, accel, cfg.accel_flat_pp),
      yoy_accel_pp: (isNum(qoq) && isNum(qoqYearAgo)) ? (qoq - qoqYearAgo) * 100 : null,
      rev_yoy: growth(at('rev', s), at('rev', s + 4)),
      gm, dgm_yoy_bps: (isNum(gm) && isNum(gmYearAgo)) ? (gm - gmYearAgo) * 10000 : null,
      eps: r.eps, eps_yoy: growth(at('eps', s), at('eps', s + 4)),
      dps: r.dps, status: '實際',
    });
  }
  const e0 = rows.find((x) => x.seq === 0);
  if (e0) {
    out.push({
      period: e0.period, end: e0.end, rev: null, qoq: null, prev_qoq: null,
      accel_pp: null, momentum: null, yoy_accel_pp: null, rev_yoy: null,
      gm: null, dgm_yoy_bps: null, eps: e0.eps, eps_yoy: null, dps: null, status: '預估',
    });
  }
  return out;
}

/* ── 年度檢視（由季度 rollup，只計實際季）───────────────── */
export function yearTable(rows, est, summary, cfg) {
  const actual = rows.filter((r) => !r.est && isNum(r.fy));
  const years = [...new Set(actual.map((r) => r.fy))].sort((a, b) => a - b);
  const shown = years.slice(-cfg.years_shown);
  const pick = (fy) => actual.filter((r) => r.fy === fy);
  const sum = (arr, f) => arr.reduce((s, r) => s + (isNum(r[f]) ? r[f] : 0), 0);
  const q4 = (fy, f) => { const r = actual.find((x) => x.fy === fy && x.fq === 4); return r && isNum(r[f]) ? r[f] : 0; };

  const out = [];
  let prev = null;
  for (const fy of shown) {
    const g = pick(fy);
    const row = { fy, n: g.length };
    if (g.length >= 4) {
      row.rev = sum(g, 'rev');
      row.gm = div(sum(g, 'gp'), row.rev);
      row.eps = sum(g, 'eps');
      row.dps = sum(g, 'dps');
      row.px = q4(fy, 'px') || null;
      row.rev_yoy = (prev && isNum(prev.rev)) ? growth(row.rev, prev.rev) : null;
      row.gm_chg_bps = (prev && isNum(prev.gm) && isNum(row.gm)) ? (row.gm - prev.gm) * 10000 : null;
      row.eps_yoy = (prev && isNum(prev.eps)) ? growth(row.eps, prev.eps) : null;
      const eqNew = q4(fy, 'eq'), eqOld = q4(fy - 1, 'eq');
      if (Math.min(eqNew, eqOld) <= 0) {
        row.roe = NM;
      } else {
        let ni = 0;
        for (const r of g) if (isNum(r.eps) && isNum(r.sh)) ni += r.eps * r.sh;
        row.roe = div(ni, (eqNew + eqOld) / 2) ?? NM;
      }
      row.yield = (isNum(row.px) && isNum(row.dps) && row.dps !== 0) ? row.dps / row.px : null;
      row.pe = (isNum(row.px) && isNum(row.eps) && row.eps > 0) ? row.px / row.eps : null;
      row.peg = (!isNum(row.pe) || !isNum(row.eps_yoy)) ? null
        : (row.eps_yoy < cfg.peg_min_growth ? NM : row.pe / (row.eps_yoy * 100));
      prev = row;
    }
    out.push(row);
  }
  return {
    years: out,
    ttm: {
      label: 'TTM / 最新季', rev: summary.rev_ttm, rev_yoy: summary.rev_yoy_q,
      gm: summary.gm_ttm, gm_chg_bps: summary.gm_vs_5y_bps, eps: summary.eps_ttm,
      eps_yoy: summary.eps_ttm_yoy, roe: summary.roe_ttm, dps: summary.dps_ttm,
      px: summary.price, yield: summary.yield, pe: summary.pe_ttm, peg: summary.peg_t,
    },
    fwd: {
      label: 'Forward（分析師共識）', eps: summary.eps_ntm, eps_yoy: summary.eps_growth_f,
      roe: summary.roe_fwd, pe: summary.pe_ntm, peg: summary.peg_f,
    },
  };
}
