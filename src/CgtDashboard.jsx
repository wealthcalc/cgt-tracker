import React, { useState, useMemo, useCallback, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  Plus, Trash2, Download, Upload, Wand2, RefreshCw, Moon, Sun,
  TableProperties, Receipt, FlaskConical, FileUp, AlertTriangle, Check,
  Wallet, TrendingUp, TrendingDown, FileText, Printer, AlertCircle, PoundSterling,
} from "lucide-react";

// Safe localStorage wrapper: persists on the deployed app, silently no-ops in
// sandboxed preview frames where storage access throws.
const store = {
  get(k, fallback) { try { const v = localStorage.getItem(k); return v == null ? fallback : JSON.parse(v); } catch { return fallback; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* sandbox */ } },
};

/* ======================================================================
   HMRC CGT engine — ported verbatim from the validated Python/JS engine
   (15/15 parity tests passing). Order: same-day -> 30-day B&B -> S104 pool.
   All amounts GBP (FX-converted upstream).
   ====================================================================== */
const MS = 86400000;
const dUTC = (s) => new Date(s + "T00:00:00Z");
const daysBetween = (a, b) => Math.round((b - a) / MS);

function matchWithPool(txns) {
  const acqs = txns.filter((t) => t.side === "BUY")
    .map((t) => ({ t, date: dUTC(t.date), remaining: t.quantity, unit: t.gbpAmount / t.quantity }))
    .sort((a, b) => a.date - b.date);
  const disps = txns.filter((t) => t.side === "SELL")
    .map((t) => ({ t, date: dUTC(t.date), remaining: t.quantity, legs: [] }))
    .sort((a, b) => a.date - b.date);
  // ERI: excess reportable income, a cost-only uplift to the S104 pool on the fund
  // distribution date. No units; excluded from same-day / 30-day matching.
  const eris = txns.filter((t) => t.side === "ERI").map((t) => ({ date: dUTC(t.date), cost: t.gbpAmount }));

  const alloc = (d, n, cost, method, acqDate) => {
    const proceeds = d.t.gbpAmount * (n / d.t.quantity);
    d.legs.push({ method, quantity: n, proceeds, cost, gain: proceeds - cost, matchedAcqDate: acqDate });
    d.remaining -= n;
  };
  for (const d of disps) for (const a of acqs) {
    if (d.remaining <= 1e-9) break;
    if (a.remaining <= 0 || +a.date !== +d.date) continue;
    const n = Math.min(d.remaining, a.remaining); alloc(d, n, a.unit * n, "SAME_DAY", a.t.date); a.remaining -= n;
  }
  for (const d of disps) for (const a of acqs) {
    if (d.remaining <= 1e-9) break;
    if (a.remaining <= 0) continue;
    const g = daysBetween(d.date, a.date);
    if (g > 0 && g <= 30) { const n = Math.min(d.remaining, a.remaining); alloc(d, n, a.unit * n, "THIRTY_DAY", a.t.date); a.remaining -= n; }
  }
  const ev = [];
  for (const a of acqs) if (a.remaining > 0) ev.push([a.date, 0, a]);
  for (const e of eris) ev.push([e.date, 0, { eri: true, cost: e.cost }]);
  for (const d of disps) if (d.remaining > 1e-9) ev.push([d.date, 1, d]);
  ev.sort((x, y) => x[0] - y[0] || x[1] - y[1]);
  let pq = 0, pc = 0;
  for (const [, kind, o] of ev) {
    if (kind === 0) { if (o.eri) { if (pq > 1e-9) pc += o.cost; continue; } pq += o.remaining; pc += o.unit * o.remaining; o.remaining = 0; }
    else {
      const n = o.remaining;
      if (n > pq + 1e-6) throw new Error(`Disposal ${o.t.date} ${o.t.ticker} exceeds shares held (needs ${round4(n)}, pool holds ${round4(pq)}).`);
      const cost = pq > 0 ? pc * (n / pq) : 0; alloc(o, n, cost, "SECTION_104", null); pq -= n; pc -= cost;
    }
  }
  const results = disps.map((x) => ({
    date: x.t.date, ticker: x.t.ticker, quantity: x.t.quantity, proceeds: x.t.gbpAmount,
    legs: x.legs, cost: x.legs.reduce((s, l) => s + l.cost, 0),
    gain: x.legs.reduce((s, l) => s + l.gain, 0), taxYear: ukTaxYear(x.t.date), id: x.t.id,
  })).sort((a, b) => (a.date < b.date ? -1 : 1));
  return { results, poolQty: pq, poolCost: pc };
}

function matchPortfolio(txns) {
  const by = {}; for (const t of txns) (by[t.ticker] ||= []).push(t);
  const all = []; const pools = {};
  for (const [tk, ts] of Object.entries(by)) {
    const { results, poolQty, poolCost } = matchWithPool(ts);
    all.push(...results); pools[tk] = { qty: poolQty, cost: poolCost };
  }
  all.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : a.ticker.localeCompare(b.ticker)));
  return { disposals: all, pools };
}

function ukTaxYear(s) {
  const [y, m, d] = s.split("-").map(Number);
  const start = m > 4 || (m === 4 && d >= 6) ? y : y - 1;
  return `${start}/${String(start + 1).slice(-2)}`;
}
// Historical UK tax-year parameters for SHARES / non-property assets.
// aea = annual exempt amount; basicLimit = income-tax basic-rate limit (band above the
// personal allowance); pa = personal allowance; reportThreshold = proceeds figure above
// which a disposal must be reported (4xAEA before 2023/24, fixed £50k after). 2024/25
// splits at the 30 Oct 2024 Budget (10/20 -> 18/24). Verified against GOV.UK / HoC Library.
const _Y = (aea, basicLimit, pa, reportThreshold, rates) => ({ aea, basicLimit, pa, reportThreshold, rates });
const _ONE = (basic, higher) => [{ from: "0000-00-00", basic, higher }];
const TAX_YEARS = {
  "2015/16": _Y(11100, 31785, 10600, 44400, _ONE(0.18, 0.28)),
  "2016/17": _Y(11100, 32000, 11000, 44400, _ONE(0.10, 0.20)),
  "2017/18": _Y(11300, 33500, 11500, 45200, _ONE(0.10, 0.20)),
  "2018/19": _Y(11700, 34500, 11850, 46800, _ONE(0.10, 0.20)),
  "2019/20": _Y(12000, 37500, 12500, 48000, _ONE(0.10, 0.20)),
  "2020/21": _Y(12300, 37500, 12500, 49200, _ONE(0.10, 0.20)),
  "2021/22": _Y(12300, 37700, 12570, 49200, _ONE(0.10, 0.20)),
  "2022/23": _Y(12300, 37700, 12570, 49200, _ONE(0.10, 0.20)),
  "2023/24": _Y(6000, 37700, 12570, 50000, _ONE(0.10, 0.20)),
  "2024/25": _Y(3000, 37700, 12570, 50000, [
    { from: "0000-00-00", basic: 0.10, higher: 0.20 },
    { from: "2024-10-30", basic: 0.18, higher: 0.24 },
  ]),
  "2025/26": _Y(3000, 37700, 12570, 50000, _ONE(0.18, 0.24)),
  "2026/27": _Y(3000, 37700, 12570, 50000, _ONE(0.18, 0.24)),
};
const LATEST_YEAR = "2026/27";
const cfgFor = (year) => TAX_YEARS[year] || { ...TAX_YEARS[LATEST_YEAR], assumed: true };
const aeaForYear = (year) => cfgFor(year).aea;
const rateForDate = (cfg, dateStr) => { let p = cfg.rates[0]; for (const r of cfg.rates) if (r.from <= dateStr) p = r; return p; };
// Personal allowance tapers by £1 for every £2 of income over £100,000.
const paFor = (pa, income) => (income <= 100000 ? pa : Math.max(0, pa - (income - 100000) / 2));

function liabilityForYear(disposals, { income = 0, carriedLosses = 0 } = {}) {
  const zero = { gains: 0, losses: 0, usedCarried: 0, aea: 0, taxable: 0, atBasic: 0, atHigher: 0, tax: 0, proceeds: 0, net: 0, reporting: false, breakdown: [], assumed: false, personalAllowance: 0, taxableIncome: 0 };
  if (!disposals.length) return zero;
  const cfg = cfgFor(disposals[0].taxYear);
  const entries = []; let losses = 0, proceeds = 0;
  for (const d of disposals) {
    proceeds += d.proceeds;
    if (d.gain > 0) { const r = rateForDate(cfg, d.date); entries.push({ amount: d.gain, basic: r.basic, higher: r.higher }); }
    else losses += -d.gain;
  }
  const gains = entries.reduce((s, e) => s + e.amount, 0);
  const net = gains - losses;
  let usedCarried = 0;
  if (net > cfg.aea && carriedLosses > 0) usedCarried = Math.min(net - cfg.aea, carriedLosses);
  // losses + carried losses + AEA reduce the highest-rate gains first (taxpayer-favourable).
  entries.sort((a, b) => b.higher - a.higher || b.basic - a.basic);
  let reductions = losses + usedCarried + cfg.aea;
  for (const e of entries) { const cut = Math.min(e.amount, reductions); e.amount -= cut; reductions -= cut; if (reductions <= 0) break; }
  // Income consumes the basic-rate band only after the personal allowance. Unused PA
  // cannot shelter gains; gains are the top slice above taxable income.
  const personalAllowance = paFor(cfg.pa, income);
  const taxableIncome = Math.max(0, income - personalAllowance);
  let bandLeft = Math.max(0, cfg.basicLimit - taxableIncome);
  const taxableEntries = entries.filter((e) => e.amount > 0).sort((a, b) => (b.higher - b.basic) - (a.higher - a.basic));
  let tax = 0, atBasic = 0, atHigher = 0; const byRate = {};
  for (const e of taxableEntries) {
    const b = Math.min(e.amount, bandLeft), h = e.amount - b;
    atBasic += b; atHigher += h; bandLeft -= b; tax += b * e.basic + h * e.higher;
    if (b > 0) byRate[e.basic] = (byRate[e.basic] || 0) + b;
    if (h > 0) byRate[e.higher] = (byRate[e.higher] || 0) + h;
  }
  const breakdown = Object.entries(byRate).map(([rate, amount]) => ({ rate: +rate, amount, tax: amount * +rate })).sort((a, b) => a.rate - b.rate);
  return { gains, losses, usedCarried, aea: cfg.aea, taxable: atBasic + atHigher, atBasic, atHigher, tax, proceeds, net, reporting: tax > 0 || proceeds > cfg.reportThreshold, breakdown, assumed: !!cfg.assumed, personalAllowance, taxableIncome };
}
const sharesForTargetGain = (q, c, p, target) => {
  const per = p - c / q; if (per <= 0) return q; return Math.min(q, Math.floor(target / per));
};
const fmtRate = (r) => `${(r * 100).toFixed(0)}%`;

// Chain per-year CGT liability across tax years, carrying losses forward.
// In-year losses offset in-year gains fully; brought-forward losses reduce net
// gains only down to the AEA; unused losses carry forward (4-year claim window).
function liabilityAllYears(disposals, { incomeByYear = {}, initialCarried = 0 } = {}) {
  const byYear = {}; for (const d of disposals) (byYear[d.taxYear] ||= []).push(d);
  const years = Object.keys(byYear).sort();
  let carried = initialCarried; const results = {};
  for (const y of years) {
    const res = liabilityForYear(byYear[y], { income: incomeByYear[y] || 0, carriedLosses: carried });
    const carriedInto = carried, inYearNetLoss = Math.max(0, res.losses - res.gains);
    carried = carried - res.usedCarried + inYearNetLoss;
    results[y] = { ...res, carriedInto, carriedOut: carried, inYearNetLoss };
  }
  return { years, results, carriedForward: carried };
}
const unitsHeldAt = (txns, dateStr, ticker) => {
  const want = ticker ? String(ticker).toUpperCase() : null;
  let q = 0; for (const t of txns) { if (t.side !== "BUY" && t.side !== "SELL") continue; if (want && String(t.ticker || "").toUpperCase() !== want) continue; if (t.date <= dateStr) q += (t.side === "BUY" ? 1 : -1) * t.quantity; } return q;
};

/* ---- UK income tax on investment income (dividends + interest), stacked on
   salary. Nil-rate allowances (PSA, dividend allowance) sit at 0% but occupy band
   space. Verified in incometax.test.mjs (11/11). ---- */
const _I = (pa, basicLimit, addl, divAllow, div, sav, psa) => ({ pa, basicLimit, addl, divAllow, div, sav, psa });
const _DO = { basic: 0.075, higher: 0.325, addl: 0.381 }, _DM = { basic: 0.0875, higher: 0.3375, addl: 0.3935 }, _DN = { basic: 0.1075, higher: 0.3575, addl: 0.3935 };
const _SAV = { basic: 0.20, higher: 0.40, addl: 0.45 }, _PSA = { basic: 1000, higher: 500, addl: 0 };
const INCOME_YEARS = {
  "2016/17": _I(11000, 32000, 150000, 5000, _DO, _SAV, _PSA), "2017/18": _I(11500, 33500, 150000, 5000, _DO, _SAV, _PSA),
  "2018/19": _I(11850, 34500, 150000, 2000, _DO, _SAV, _PSA), "2019/20": _I(12500, 37500, 150000, 2000, _DO, _SAV, _PSA),
  "2020/21": _I(12500, 37500, 150000, 2000, _DO, _SAV, _PSA), "2021/22": _I(12570, 37700, 150000, 2000, _DO, _SAV, _PSA),
  "2022/23": _I(12570, 37700, 150000, 2000, _DM, _SAV, _PSA), "2023/24": _I(12570, 37700, 125140, 1000, _DM, _SAV, _PSA),
  "2024/25": _I(12570, 37700, 125140, 500, _DM, _SAV, _PSA), "2025/26": _I(12570, 37700, 125140, 500, _DM, _SAV, _PSA),
  "2026/27": _I(12570, 37700, 125140, 500, _DN, _SAV, _PSA),
};
const incomeCfg = (year) => INCOME_YEARS[year] || { ...INCOME_YEARS["2026/27"], assumed: true };
function _walk(pos, amount, basicTop, higherTop, rates) {
  const bounds = [basicTop, higherTop, Infinity], rs = [rates.basic, rates.higher, rates.addl];
  let tax = 0, p = pos, rem = amount;
  for (let i = 0; i < 3 && rem > 1e-9; i++) { if (p >= bounds[i]) continue; const take = Math.min(rem, bounds[i] - p); tax += take * rs[i]; p += take; rem -= take; }
  return { tax, end: p };
}
function investmentIncomeTax({ salary = 0, interest = 0, dividends = 0, year } = {}) {
  const c = incomeCfg(year), ani = salary + interest + dividends;
  const pa = paFor(c.pa, ani), basicTop = c.basicLimit, higherTop = Math.max(basicTop, c.addl - pa);
  let paLeft = pa; const net = (x) => { const u = Math.min(x, paLeft); paLeft -= u; return x - u; };
  const salT = net(salary), intT = net(interest), divT = net(dividends), taxableTotal = salT + intT + divT;
  const band = taxableTotal <= basicTop ? "basic" : taxableTotal <= higherTop ? "higher" : "addl";
  const psa = c.psa[band], startRate = Math.max(0, 5000 - salT);
  let pos = salT, interestTax = 0, dividendTax = 0;
  { let rem = intT; const z = Math.min(rem, startRate + psa); pos += z; rem -= z; const r = _walk(pos, rem, basicTop, higherTop, c.sav); interestTax = r.tax; pos = r.end; }
  { let rem = divT; const z = Math.min(rem, c.divAllow); pos += z; rem -= z; const r = _walk(pos, rem, basicTop, higherTop, c.div); dividendTax = r.tax; pos = r.end; }
  const r2 = (x) => Math.round(x * 100) / 100;
  return { year, assumed: !!c.assumed, interestTax: r2(interestTax), dividendTax: r2(dividendTax), tax: r2(interestTax + dividendTax), personalAllowance: pa, band, divAllow: c.divAllow, psa };
}

/* ---- IBKR CSV import (Flex Query + Activity Statement) -> trades + income.
   Verified in ibkr.test.mjs (20/20). ---- */
function parseCSVRows(text) {
  const rows = []; let row = [], cell = "", q = false;
  const s = text.replace(/\r\n?/g, "\n");
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (q) { if (c === '"') { if (s[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += c; }
    else if (c === '"') q = true;
    else if (c === ",") { row.push(cell); cell = ""; }
    else if (c === "\n") { row.push(cell); rows.push(row); row = []; cell = ""; }
    else cell += c;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((c) => c.trim() !== ""));
}
const _ibnorm = (h) => h.toLowerCase().replace(/[^a-z0-9]/g, "");
const _ibnum = (x) => { if (x == null) return 0; const n = parseFloat(String(x).replace(/,/g, "")); return isFinite(n) ? n : 0; };
const _ibdate = (x) => {
  if (!x) return ""; const s = String(x).trim(); let m = s.match(/^(\d{4})(\d{2})(\d{2})$/); if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{4})\/(\d{2})\/(\d{2})/); if (m) return `${m[1]}-${m[2]}-${m[3]}`; return "";
};
const _IBSTOCK = new Set(["stk", "etf", "fund", "stocks", "equity", "closedendfund"]);
const _ibpick = (headerIndex) => (row, ...keys) => { for (const k of keys) { const i = headerIndex[k]; if (i != null && row[i] !== undefined) return row[i]; } return undefined; };
function _ibTrade(get, defaultWrapper, baseCurrency, warnings) {
  const symbol = (get("symbol", "underlyingsymbol") || "").trim().toUpperCase();
  const date = _ibdate(get("tradedate", "datetime", "date"));
  if (!symbol || !date) return null;
  const asset = _ibnorm(get("assetclass", "assetcategory") || "stk");
  if (asset && !_IBSTOCK.has(asset)) { warnings.push(`Skipped ${symbol} ${date}: asset class "${asset}" not supported.`); return null; }
  const qtyRaw = _ibnum(get("quantity"));
  const bs = (get("buysell") || "").trim().toUpperCase();
  const side = bs ? (bs.startsWith("S") ? "SELL" : "BUY") : (qtyRaw < 0 ? "SELL" : "BUY");
  const quantity = Math.abs(qtyRaw); if (!quantity) return null;
  const currency = (get("currencyprimary", "currency") || "GBP").trim().toUpperCase();
  const proceeds = _ibnum(get("proceeds")), commission = _ibnum(get("ibcommission", "commission", "commfee", "commissionandtax")), taxes = _ibnum(get("taxes", "tax"));
  const netcash = get("netcash");
  const native = Math.abs(netcash !== undefined && netcash !== "" ? _ibnum(netcash) + taxes : proceeds + commission + taxes);
  const fxToBase = _ibnum(get("fxratetobase", "fxrate"));
  const isin = (get("isin", "securityid") || "").trim().toUpperCase();
  let gbpAmount = null, fxRate = null, needsFx = false;
  if (currency === "GBP") { gbpAmount = native; fxRate = 1; }
  else if (baseCurrency === "GBP" && fxToBase) { gbpAmount = native * fxToBase; fxRate = fxToBase; }
  else needsFx = true;
  return { date, ticker: symbol, isin, side, quantity, nativeCurrency: currency, nativeAmount: native, fxRate, gbpAmount: gbpAmount == null ? null : Math.round(gbpAmount * 100) / 100, needsFx, wrapper: defaultWrapper };
}
function _ibCash(get, defaultWrapper, baseCurrency) {
  const typ = _ibnorm(get("type", "activitydescription", "description") || "");
  let kind = null;
  if (typ.includes("withholding")) return null;
  if (typ.includes("dividend") || typ.includes("inlieu") || typ.includes("paymentinlieu")) kind = "dividend";
  else if (typ.includes("interest")) kind = "interest"; else return null;
  const amount = _ibnum(get("amount")); if (amount <= 0) return null;
  const date = _ibdate(get("settledate", "reportdate", "date", "paydate")); if (!date) return null;
  const currency = (get("currencyprimary", "currency") || "GBP").trim().toUpperCase();
  const fxToBase = _ibnum(get("fxratetobase", "fxrate"));
  const symbol = (get("symbol", "underlyingsymbol") || "").trim().toUpperCase();
  const isin = (get("isin", "securityid") || "").trim().toUpperCase();
  let gbp = null, fxRate = null, needsFx = false;
  if (currency === "GBP") { gbp = amount; fxRate = 1; }
  else if (baseCurrency === "GBP" && fxToBase) { gbp = amount * fxToBase; fxRate = fxToBase; }
  else needsFx = true;
  return { date, ticker: symbol, isin, kind, nativeCurrency: currency, nativeAmount: amount, fxRate, amount: gbp == null ? null : Math.round(gbp * 100) / 100, needsFx, wrapper: defaultWrapper };
}
function _ibFlex(rows, defaultWrapper, baseCurrency, warnings) {
  const header = rows[0].map(_ibnorm); const headerIndex = {}; header.forEach((h, i) => { if (!(h in headerIndex)) headerIndex[h] = i; });
  const pick = _ibpick(headerIndex); const has = (...k) => k.some((x) => headerIndex[x] != null);
  const looksTrades = has("tradedate", "buysell", "tradeprice") || (has("quantity") && has("proceeds"));
  const looksCash = has("type", "amount") && !has("tradeprice", "buysell");
  const trades = [], income = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r]; const get = (...k) => pick(row, ...k);
    if (looksCash && !looksTrades) { const c = _ibCash(get, defaultWrapper, baseCurrency); if (c) income.push(c); continue; }
    const t = _ibTrade(get, defaultWrapper, baseCurrency, warnings); if (t) { trades.push(t); continue; }
    if (has("type", "amount")) { const c = _ibCash(get, defaultWrapper, baseCurrency); if (c) income.push(c); }
  }
  return { trades, income };
}
function _ibActivity(rows, defaultWrapper, baseCurrency, warnings) {
  const trades = [], income = []; const sections = {};
  for (const row of rows) { const name = row[0], tag = row[1];
    if (tag === "Header") sections[name] = { header: row.slice(2).map(_ibnorm), data: [] };
    else if (tag === "Data" && sections[name]) sections[name].data.push(row.slice(2)); }
  const build = (sec) => { const idx = {}; sec.header.forEach((h, i) => { if (!(h in idx)) idx[h] = i; }); return { idx, pick: _ibpick(idx) }; };
  if (sections["Trades"]) { const { pick } = build(sections["Trades"]);
    for (const row of sections["Trades"].data) { const get = (...k) => pick(row, ...k);
      const disc = _ibnorm(get("datadiscriminator") || "order"); if (disc && !["order", "trade"].includes(disc)) continue;
      const t = _ibTrade(get, defaultWrapper, baseCurrency, warnings); if (t) trades.push(t); } }
  for (const secName of ["Dividends", "Payment In Lieu Of Dividends", "Interest"]) {
    if (!sections[secName]) continue; const { pick } = build(sections[secName]);
    for (const row of sections[secName].data) { const get = (...k) => pick(row, ...k);
      const desc = get("description") || ""; const isinM = String(desc).match(/\(([A-Z]{2}[A-Z0-9]{9}\d)\)/); const symM = String(desc).match(/^([A-Z0-9.]+)\b/);
      const kind = secName === "Interest" ? "interest" : "dividend"; const amount = _ibnum(get("amount")); if (amount <= 0) continue;
      const date = _ibdate(get("date", "settledate", "reportdate")); if (!date) continue;
      const currency = (get("currency", "currencyprimary") || "GBP").trim().toUpperCase();
      let gbp = null, needsFx = false; if (currency === "GBP") gbp = amount; else needsFx = true;
      income.push({ date, ticker: symM ? symM[1].toUpperCase() : "", isin: isinM ? isinM[1] : "", kind, nativeCurrency: currency, nativeAmount: amount, fxRate: currency === "GBP" ? 1 : null, amount: gbp == null ? null : Math.round(gbp * 100) / 100, needsFx, wrapper: defaultWrapper }); } }
  return { trades, income };
}
function parseIBKR(text, { defaultWrapper = "GIA", baseCurrency = "GBP" } = {}) {
  const rows = parseCSVRows(text); if (!rows.length) return { trades: [], income: [], warnings: ["Empty file."], baseCurrency };
  const warnings = [];
  const sectioned = rows.some((r) => r[1] === "Header" && ["Trades", "Dividends", "Interest", "Deposits & Withdrawals", "Payment In Lieu Of Dividends"].includes(r[0]));
  const { trades, income } = sectioned ? _ibActivity(rows, defaultWrapper, baseCurrency, warnings) : _ibFlex(rows, defaultWrapper, baseCurrency, warnings);
  const needFx = trades.filter((t) => t.needsFx).length + income.filter((t) => t.needsFx).length;
  if (needFx) warnings.push(`${needFx} row(s) in a non-GBP currency need an FX rate; fetching by trade date.`);
  return { trades, income, warnings, baseCurrency, format: sectioned ? "activity" : "flex" };
}

/* ---- Multi-year AEA disposal / gain-harvesting optimiser. Verified in
   optimiser.test.mjs (14/14). ---- */
const nextTaxYear = (y) => { const a = Number(y.split("/")[0]) + 1; return `${a}/${String(a + 1).slice(-2)}`; };
function optimiseDisposals({ holdings, startYear, years = 10, income = 0, useBasicBand = false, growth = 0 }) {
  let hs = holdings.map((h) => ({ ticker: h.ticker, qty: +h.qty, avgCost: +h.qty ? +h.cost / +h.qty : 0, price: +h.price })).filter((h) => h.qty > 0 && isFinite(h.price) && h.price > 0);
  const embedded = () => hs.reduce((s, h) => s + Math.max(0, h.qty * (h.price - h.avgCost)), 0);
  const startEmbedded = embedded(); const schedule = []; let y = startYear, totalWashed = 0, yearsToClear = null;
  for (let i = 0; i < years; i++) {
    const cfg = cfgFor(y); const pa = paFor(cfg.pa, income); const taxableIncome = Math.max(0, income - pa);
    const bandRoom = Math.max(0, cfg.basicLimit - taxableIncome); const rate = cfg.rates[cfg.rates.length - 1];
    const gainBudget = cfg.aea + (useBasicBand ? bandRoom : 0);
    let budgetLeft = gainBudget, realised = 0; const sells = [];
    const order = hs.map((h, idx) => ({ idx, gps: h.price - h.avgCost })).filter((o) => o.gps > 0).sort((a, b) => b.gps - a.gps);
    for (const { idx } of order) {
      if (budgetLeft <= 1e-6) break; const h = hs[idx], gps = h.price - h.avgCost;
      const takeGain = Math.min(h.qty * gps, budgetLeft); const shares = takeGain / gps;
      h.avgCost = ((h.qty - shares) * h.avgCost + shares * h.price) / h.qty;
      realised += takeGain; budgetLeft -= takeGain;
      sells.push({ ticker: h.ticker, shares: Math.round(shares * 1e4) / 1e4, gain: Math.round(takeGain * 100) / 100 });
    }
    const aeaUsed = Math.min(realised, cfg.aea); const bandGain = Math.max(0, realised - cfg.aea);
    const tax = Math.round((useBasicBand ? bandGain * rate.basic : 0) * 100) / 100; totalWashed += realised;
    const remaining = embedded();
    schedule.push({ year: y, aea: cfg.aea, gainBudget, gainRealised: Math.round(realised * 100) / 100, aeaUsed: Math.round(aeaUsed * 100) / 100, bandGain: Math.round(bandGain * 100) / 100, tax, sells, cumulativeWashed: Math.round(totalWashed * 100) / 100, remainingUnrealised: Math.round(remaining * 100) / 100 });
    if (remaining <= 1e-6 && yearsToClear == null) yearsToClear = i + 1;
    if (remaining <= 1e-6) break;
    if (growth) for (const h of hs) h.price *= 1 + growth; y = nextTaxYear(y);
  }
  return { schedule, yearsToClear, totalWashed: Math.round(totalWashed * 100) / 100, startEmbedded: Math.round(startEmbedded * 100) / 100, remainingAfter: Math.round(embedded() * 100) / 100 };
}

// Seed ISIN/domicile data for holdings, sourced from issuer fact sheets. Only
// covers tickers actually held (per the 2026-07 GIA ledger import); anything
// else is left blank for the user to fill in. `eri` flags offshore reporting
// funds (Irish/Lux-domiciled accumulating ETFs) that generate excess reportable
// income when held unsheltered; UK investment trusts are ordinary companies and
// pay ordinary dividends, not ERI, even though they accumulate/reinvest.
const SECURITY_SEED = {
  SJPA: { isin: "IE00B4L5YX21", name: "iShares Core MSCI Japan IMI UCITS ETF (Acc)", domicile: "IE", eri: true },
  CSP1: { isin: "IE00B5BMR087", name: "iShares Core S&P 500 UCITS ETF (Acc)", domicile: "IE", eri: true },
  EMIM: { isin: "IE00BKM4GZ66", name: "iShares Core MSCI EM IMI UCITS ETF (Acc)", domicile: "IE", eri: true },
  XNAQ: { isin: "IE00BMFKG444", name: "Xtrackers Nasdaq 100 UCITS ETF 1C", domicile: "IE", eri: true },
  UIFS: { isin: "IE00B4JNQZ49", name: "iShares S&P 500 Financials Sector UCITS ETF (Acc)", domicile: "IE", eri: true },
  HSTC: { isin: "IE00BMWXKN31", name: "HSBC Hang Seng TECH UCITS ETF", domicile: "IE", eri: true },
  IITU: { isin: "IE00B3WJKG14", name: "iShares S&P 500 Information Technology Sector UCITS ETF (Acc)", domicile: "IE", eri: true },
  SPXL: { isin: "IE000XZSV718", name: "SPDR S&P 500 UCITS ETF (Acc)", domicile: "IE", eri: true },
  SMT: { isin: "GB00BLDYK618", name: "Scottish Mortgage Investment Trust plc", domicile: "GB", eri: false },
  ATT: { isin: "GB00BNG2M159", name: "Allianz Technology Trust plc", domicile: "GB", eri: false },
  BNKR: { isin: "GB00BN4NDR39", name: "Bankers Investment Trust plc", domicile: "GB", eri: false },
  AIAG: { isin: "IE00BK5BCD43", name: "iShares AI & Automation Growth UCITS ETF", domicile: "IE", eri: true },
  CYSE: { isin: "IE00BLPK3577", name: "iShares Cybersecurity UCITS ETF", domicile: "IE", eri: true },
  DFEU: { isin: "IE000IAXNM41", name: "Defense Europe UCITS ETF", domicile: "IE", eri: true },
  RBTX: { isin: "IE00BYZK4552", name: "iShares Robotics & Automation UCITS ETF", domicile: "IE", eri: true },
  RENG: { isin: "IE00BK5BCH80", name: "iShares Renewable Energy UCITS ETF", domicile: "IE", eri: true },
};

/* ---- iShares/BlackRock "UK Reportable Income" workbook parser (one per fund
   umbrella — iShares Plc, iShares II-VII plc — one row per share class per
   accounting period, keyed by ISIN). Verified against 8 real files spanning
   2021-2025 across five umbrellas (see ishares.test.mjs, 29/29). Confirmed
   structure: Fund Umbrella Name | Fund Name | [Share Class Name] | ISIN |
   [HMRC share class reference] | Reporting Period | Currency | Statement
   Under Regulation 92(1)(e) | Excess of Reported Income per Unit | Fund
   Distribution Date | Meets definition of a Bond Fund for the period |
   Actual Distribution per Unit/Date - 1, 2, 3... (repeating, unused here).
   "Reporting Period" is a text range ("01 July 2024 to 30 June 2025" /
   "01.12.2024 to 30.11.2025" / "01/03/2020 to 28/02/2021" depending on
   report year) — period end is the second date. "Meets definition of a Bond
   Fund" is the authoritative dividend-vs-interest signal per HMRC's
   offshore-fund ERI treatment rule. Column detection stays keyword-based
   (not fixed-position) since column order shifts between umbrellas and
   report years. ---- */
const _isNorm = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9]/g, " ").replace(/\s+/g, " ").trim();
const _isHasAny = (h, ...words) => words.some((w) => h.includes(w));
const _IS_MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
const _pad2 = (n) => String(n).padStart(2, "0");
function _isParseDateText(s) {
  const t = String(s ?? "").trim(); if (!t) return "";
  let m = t.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/); // "31 May 2025"
  if (m) { const mo = _IS_MONTHS[m[2].slice(0, 3).toLowerCase()]; if (mo) return `${m[3]}-${_pad2(mo)}-${_pad2(+m[1])}`; }
  m = t.match(/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})$/); // "31/05/2025" or "31.05.2025" (dd/mm/yyyy)
  if (m) return `${m[3]}-${_pad2(+m[2])}-${_pad2(+m[1])}`;
  m = t.match(/^(\d{4})-(\d{2})-(\d{2})/); if (m) return `${m[1]}-${m[2]}-${m[3]}`; // "2025-05-31"
  return "";
}
function _isExcelSerialToISO(n) {
  const num = +n; if (!isFinite(num)) return "";
  const ms = Math.round((num - 25569) * 86400 * 1000);
  const d = new Date(ms); return isNaN(d) ? "" : d.toISOString().slice(0, 10);
}
function _isCellToISO(v) {
  if (v == null || v === "") return "";
  if (v instanceof Date) return isNaN(v) ? "" : v.toISOString().slice(0, 10);
  if (typeof v === "number") return _isExcelSerialToISO(v);
  return _isParseDateText(v);
}
function _isRangeEndToISO(v) {
  const t = String(v ?? ""); const m = t.split(/\bto\b/i);
  if (m.length < 2) return ""; return _isParseDateText(m[m.length - 1]);
}
const _ISIN_RE = /\b[A-Z]{2}[A-Z0-9]{9}\d\b/;
function findHeaderRow(aoa, maxScan = 15) {
  let best = -1, bestScore = -1;
  for (let r = 0; r < Math.min(maxScan, aoa.length); r++) {
    const row = aoa[r] || []; const cells = row.map((c) => _isNorm(c));
    const nonEmpty = cells.filter(Boolean).length; if (nonEmpty < 3) continue;
    const looksHeader = cells.some((c) => c === "isin" || c.includes("isin"));
    const score = nonEmpty + (looksHeader ? 10 : 0);
    if (score > bestScore) { bestScore = score; best = r; }
  }
  return best;
}
function guessColumnMap(headerRow) {
  const cells = (headerRow || []).map((c) => _isNorm(c));
  const find = (test) => { for (let i = 0; i < cells.length; i++) if (test(cells[i])) return i; return -1; };
  const findFundName = () => {
    let i = find((h) => _isHasAny(h, "fund name", "sub fund", "subfund"));
    if (i >= 0) return i;
    return find((h) => h.includes("fund") && h.includes("name"));
  };
  return {
    isin: find((h) => h === "isin" || h.includes("isin")),
    fundName: findFundName(),
    currency: find((h) => h === "currency" || h === "ccy" || h.includes("currency")),
    periodEnd: find((h) => _isHasAny(h, "period end", "accounting period end", "accounting date", "fund year end", "year end date", "distribution period end")),
    reportingPeriod: find((h) => h.includes("reporting period")),
    distributionDate: find((h) => h.includes("distribution") && h.includes("date") && !h.includes("actual")),
    eriPerUnit: find((h) => _isHasAny(h, "excess of reported income", "excess reportable income", "reportable income per", "eri per", "excess income per") || (h.includes("excess") && h.includes("income") && !h.includes("statement"))),
    bondFund: find((h) => h.includes("bond fund")),
    treatment: find((h) => _isHasAny(h, "income type", "type of income", "interest distribution", "dividend distribution", "excess type")),
  };
}
function extractISharesRows(aoa, headerRowIdx, colMap, holdingIsins) {
  const out = [];
  for (let r = headerRowIdx + 1; r < aoa.length; r++) {
    const row = aoa[r] || []; const get = (idx) => (idx >= 0 && idx < row.length ? row[idx] : "");
    let isin = String(get(colMap.isin) ?? "").trim().toUpperCase();
    if (!isin && colMap.isin < 0) { const m = row.map((c) => String(c ?? "")).join(" ").match(_ISIN_RE); if (m) isin = m[0]; }
    if (!_ISIN_RE.test(isin)) continue;
    if (holdingIsins && holdingIsins.size && !holdingIsins.has(isin)) continue;
    const eriRaw = get(colMap.eriPerUnit);
    const eri = typeof eriRaw === "number" ? eriRaw : parseFloat(String(eriRaw).replace(/,/g, ""));
    if (!isFinite(eri) || eri <= 0) continue; // zero/blank rows (fully distributed, or metadata) carry no tax impact

    const periodEndISO = colMap.periodEnd >= 0 ? _isCellToISO(get(colMap.periodEnd))
      : colMap.reportingPeriod >= 0 ? _isRangeEndToISO(get(colMap.reportingPeriod)) : "";
    let distributionDateISO = colMap.distributionDate >= 0 ? _isCellToISO(get(colMap.distributionDate)) : "";
    if (!distributionDateISO && periodEndISO) distributionDateISO = addMonthsISO(periodEndISO, 6);

    const currencyRaw = String(get(colMap.currency) ?? "").trim().toUpperCase();
    let treatment = "dividend";
    if (colMap.bondFund >= 0) treatment = _isNorm(get(colMap.bondFund)).startsWith("yes") ? "interest" : "dividend";
    else if (colMap.treatment >= 0) treatment = _isNorm(get(colMap.treatment)).includes("interest") ? "interest" : "dividend";

    out.push({ isin, fundName: String(get(colMap.fundName) ?? "").trim(), currency: currencyRaw === "GBX" || currencyRaw === "GBP PENCE" ? "GBp" : (currencyRaw || "GBP"), periodEnd: periodEndISO, distributionDate: distributionDateISO, perShare: Math.round(eri * 1e6) / 1e6, treatment });
  }
  return out;
}
function parseISharesWorkbook(sheets, holdingIsins) {
  return sheets.map((s) => {
    const headerRowIdx = findHeaderRow(s.aoa);
    if (headerRowIdx < 0) return { name: s.name, headerRowIdx: -1, colMap: null, headerCells: [], rows: [] };
    const headerCells = s.aoa[headerRowIdx] || [];
    const colMap = guessColumnMap(headerCells);
    const rows = colMap.eriPerUnit >= 0 || colMap.isin >= 0 ? extractISharesRows(s.aoa, headerRowIdx, colMap, holdingIsins) : [];
    return { name: s.name, headerRowIdx, colMap, headerCells, rows };
  });
}

/* ----------------------------- helpers ------------------------------ */
const round4 = (x) => Math.round(x * 1e4) / 1e4;
const gbp = (x) => (x < 0 ? "−£" : "£") + Math.abs(x).toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const num = (x, dp = 2) => (x ?? 0).toLocaleString("en-GB", { minimumFractionDigits: dp, maximumFractionDigits: dp });
const uid = () => Math.random().toString(36).slice(2, 9);
const todayISO = () => new Date().toISOString().slice(0, 10);

const SAMPLE = [
  { id: uid(), date: "2022-11-15", ticker: "WFC", side: "BUY", quantity: 120, nativeCurrency: "USD", nativeAmount: 5036, fxRate: 0.83, gbpAmount: 4180, note: "RSU vest" },
  { id: uid(), date: "2023-11-15", ticker: "WFC", side: "BUY", quantity: 140, nativeCurrency: "USD", nativeAmount: 6650, fxRate: 0.80, gbpAmount: 5320, note: "RSU vest" },
  { id: uid(), date: "2024-11-15", ticker: "WFC", side: "BUY", quantity: 160, nativeCurrency: "USD", nativeAmount: 11544, fxRate: 0.79, gbpAmount: 9120, note: "RSU vest" },
  { id: uid(), date: "2025-06-02", ticker: "WFC", side: "SELL", quantity: 200, nativeCurrency: "USD", nativeAmount: 18718, fxRate: 0.78, gbpAmount: 14600, note: "part sale" },
  { id: uid(), date: "2015-08-03", ticker: "AAPL", side: "BUY", quantity: 100, nativeCurrency: "USD", nativeAmount: 11500, fxRate: 0.64, gbpAmount: 7360, note: "" },
  { id: uid(), date: "2022-09-01", ticker: "AAPL", side: "SELL", quantity: 20, nativeCurrency: "USD", nativeAmount: 3200, fxRate: 0.86, gbpAmount: 2752, note: "" },
  { id: uid(), date: "2024-09-15", ticker: "AAPL", side: "SELL", quantity: 25, nativeCurrency: "USD", nativeAmount: 5500, fxRate: 0.76, gbpAmount: 4180, note: "pre-Budget" },
  { id: uid(), date: "2025-09-10", ticker: "AAPL", side: "SELL", quantity: 30, nativeCurrency: "USD", nativeAmount: 6900, fxRate: 0.74, gbpAmount: 5106, note: "" },
];

const METHOD = {
  SAME_DAY: { label: "Same-day", v: "--m-same" },
  THIRTY_DAY: { label: "30-day", v: "--m-bb" },
  SECTION_104: { label: "S104 pool", v: "--m-pool" },
};

/* ---- Alpha Vantage live prices (client-side; free tier: 25/day, 5/min) ----
   GLOBAL_QUOTE returns a price but NO currency, so the currency is set per
   ticker and every quote is normalised to GBP (GBp pence ÷100; USD/EUR via FX). */
const AV_URL = "https://www.alphavantage.co/query";
async function avQuote(symbol, key) {
  const res = await fetch(`${AV_URL}?function=GLOBAL_QUOTE&symbol=${encodeURIComponent(symbol)}&apikey=${encodeURIComponent(key)}`);
  const j = await res.json();
  if (j.Note || j.Information) throw new Error("Alpha Vantage limit hit (25/day, 5/min) — try again later.");
  const q = j["Global Quote"];
  const p = q && q["05. price"];
  if (p == null || p === "") throw new Error(`No quote for "${symbol}" (LSE symbols need .LON).`);
  return parseFloat(p);
}
async function fxToGBP(ccy) {
  if (ccy === "GBP" || ccy === "GBp") return 1;
  try { const r = await fetch(`https://api.frankfurter.dev/v1/latest?from=${ccy}&to=GBP`); const j = await r.json(); return j?.rates?.GBP ?? null; }
  catch { return null; }
}
// FX on a specific date (for imports that lack an FX-to-base rate).
async function fxHistorical(ccy, date) {
  if (ccy === "GBP" || ccy === "GBp") return 1;
  try { const r = await fetch(`https://api.frankfurter.dev/v1/${date}?from=${ccy}&to=GBP`); const j = await r.json(); return j?.rates?.GBP ?? null; }
  catch { return null; }
}
const toGBP = (raw, ccy, fx) => (ccy === "GBp" ? raw / 100 : ccy === "GBP" ? raw : fx ? raw * fx : null);
const avBudget = () => { const c = store.get("cgt.avcount", { date: "", n: 0 }); return c.date === todayISO() ? c : { date: todayISO(), n: 0 }; };
const avBump = () => { const c = avBudget(); c.n += 1; store.set("cgt.avcount", c); return c.n; };
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/* ============================== app =================================== */
export default function App() {
  const [dark, setDark] = useState(() => store.get("cgt.dark", true));
  const [txns, setTxns] = useState(() => store.get("cgt.txns", SAMPLE));
  const [tab, setTab] = useState("cgt");
  const [income, setIncome] = useState(() => store.get("cgt.income", 200000));
  const [carried, setCarried] = useState(() => store.get("cgt.carried", 0));
  const [incomeEntries, setIncomeEntries] = useState(() => store.get("cgt.incomeEntries", [])); // dividends/interest ledger
  const [eriEntries, setEriEntries] = useState(() => store.get("cgt.eriEntries", []));           // excess reportable income
  const [prices, setPrices] = useState(() => store.get("cgt.prices", {}));
  const [avKey, setAvKey] = useState(() => store.get("cgt.avkey", ""));
  const [avMeta, setAvMeta] = useState(() => store.get("cgt.avmeta", {}));       // { ticker: {symbol, currency} }
  const [priceMeta, setPriceMeta] = useState(() => store.get("cgt.pricemeta", {})); // { ticker: {asOf, raw, ccy} }
  const [secMeta, setSecMeta] = useState(() => ({ ...SECURITY_SEED, ...store.get("cgt.secmeta", {}) })); // { ticker: {isin, name, domicile, eri} }
  const [error, setError] = useState(null);

  // persist (guarded; no-ops in sandbox)
  React.useEffect(() => store.set("cgt.txns", txns), [txns]);
  React.useEffect(() => store.set("cgt.prices", prices), [prices]);
  React.useEffect(() => store.set("cgt.avkey", avKey), [avKey]);
  React.useEffect(() => store.set("cgt.avmeta", avMeta), [avMeta]);
  React.useEffect(() => store.set("cgt.pricemeta", priceMeta), [priceMeta]);
  React.useEffect(() => store.set("cgt.secmeta", secMeta), [secMeta]);
  React.useEffect(() => store.set("cgt.income", income), [income]);
  React.useEffect(() => store.set("cgt.carried", carried), [carried]);
  React.useEffect(() => store.set("cgt.incomeEntries", incomeEntries), [incomeEntries]);
  React.useEffect(() => store.set("cgt.eriEntries", eriEntries), [eriEntries]);
  React.useEffect(() => store.set("cgt.dark", dark), [dark]);

  // Only unsheltered (GIA) holdings are within scope for CGT and income tax;
  // ISA / SIPP / LISA are tax-free and excluded from every tax computation.
  const isGIA = (w) => (w || "GIA") === "GIA";
  const giaTxns = useMemo(() => txns.filter((t) => isGIA(t.wrapper)), [txns]);

  // Excess reportable income -> synthetic ERI txns (pool cost uplift) + income.
  const eriTxns = useMemo(() => eriEntries.map((e) => {
    const units = unitsHeldAt(giaTxns, e.periodEnd || "9999-12-31", e.ticker);
    const native = units * (+e.perShare || 0);
    const g = e.currency === "GBp" ? native / 100 : e.currency === "GBP" ? native : native * (+e.fxRate || 0);
    return { id: "eri-" + e.id, ticker: e.ticker, side: "ERI", date: e.distributionDate, quantity: 0, gbpAmount: Math.round((g || 0) * 100) / 100, _eri: e, _units: units, _gbp: g || 0 };
  }).filter((t) => t.ticker && t.date), [eriEntries, giaTxns]);

  const matched = useMemo(() => {
    try { setError(null); return matchPortfolio([...giaTxns, ...eriTxns]); }
    catch (e) { setError(e.message); return { disposals: [], pools: {} }; }
  }, [giaTxns, eriTxns]);

  const taxYears = useMemo(() => {
    const s = new Set(matched.disposals.map((d) => d.taxYear));
    return [...s].sort().reverse();
  }, [matched]);
  const [year, setYear] = useState(null);
  const activeYear = year && taxYears.includes(year) ? year : taxYears[0] || "2025/26";

  // Chain CGT losses across all tracked years (initial b/f losses = `carried`).
  const allYears = useMemo(() => {
    const yrs = [...new Set(matched.disposals.map((d) => d.taxYear))];
    const inc = Object.fromEntries(yrs.map((y) => [y, income]));
    return liabilityAllYears(matched.disposals, { incomeByYear: inc, initialCarried: carried });
  }, [matched, income, carried]);

  const yearDisposals = matched.disposals.filter((d) => d.taxYear === activeYear);
  const liab = allYears.results[activeYear] || liabilityForYear(yearDisposals, { income, carriedLosses: carried });

  // Aggregate dividends/interest per tax year (ledger + ERI-derived income).
  const incomeByYear = useMemo(() => {
    const m = {}; const add = (y, kind, amt) => { (m[y] ||= { dividends: 0, interest: 0 }); m[y][kind === "interest" ? "interest" : "dividends"] += amt; };
    for (const e of incomeEntries) if (e.date && e.amount && isGIA(e.wrapper)) add(ukTaxYear(e.date), e.kind, +e.amount);
    for (const t of eriTxns) if (t.date) add(ukTaxYear(t.date), t._eri.treatment, t._gbp);
    return m;
  }, [incomeEntries, eriTxns]);

  const fileRef = useRef(null);
  const [status, setStatus] = useState("");
  const flash = (msg) => { setStatus(msg); setTimeout(() => setStatus(""), 3500); };

  const exportJSON = async () => {
    const text = JSON.stringify(txns, null, 2);
    let downloaded = false;
    try {
      const blob = new Blob([text], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = "cgt-transactions.json";
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      downloaded = true;
    } catch { /* sandbox may block downloads */ }
    // Clipboard fallback so export never silently fails inside a sandboxed frame.
    try { await navigator.clipboard.writeText(text); flash(downloaded ? "Downloaded — also copied to clipboard." : "Download blocked here; JSON copied to clipboard instead."); }
    catch { flash(downloaded ? "Downloaded." : "Couldn't download or copy in this frame — try the deployed app."); }
  };
  const importJSON = (e) => {
    const f = e.target.files?.[0]; if (!f) { return; }
    const r = new FileReader();
    r.onload = () => {
      try {
        const d = JSON.parse(r.result);
        if (Array.isArray(d)) { setTxns(d.map((x) => ({ ...x, id: x.id || uid() }))); flash(`Imported ${d.length} transactions.`); }
        else setError("That file isn't a transaction array.");
      } catch { setError("Couldn't parse that JSON file."); }
    };
    r.readAsText(f);
    e.target.value = ""; // allow re-selecting the same file
  };

  return (
    <div className={dark ? "dark" : ""}>
      <style>{`
        .root{
          --bg:#f6f7f9;--panel:#ffffff;--panel2:#f1f3f6;--fg:#0f1729;--muted:#5b6677;
          --border:#e2e6ec;--accent:#4338ca;--accent-fg:#ffffff;--gain:#047857;--loss:#be123c;
          --m-same:#0369a1;--m-bb:#b45309;--m-pool:#4338ca;--chip:#eef1f6;
        }
        .dark .root{
          --bg:#080b12;--panel:#0f141d;--panel2:#151b26;--fg:#e8edf4;--muted:#8a97a8;
          --border:#222b38;--accent:#6366f1;--accent-fg:#ffffff;--gain:#34d399;--loss:#fb7185;
          --m-same:#38bdf8;--m-bb:#fbbf24;--m-pool:#a5b4fc;--chip:#1a2230;
        }
        .num{font-variant-numeric:tabular-nums;font-family:ui-monospace,"SF Mono",Menlo,Consolas,monospace;}
        @media print {
          body * { visibility: hidden !important; }
          .print-area, .print-area * { visibility: visible !important; }
          .print-area { position: absolute; left: 0; top: 0; width: 100%; padding: 24px;
            --bg:#fff; --panel:#fff; --panel2:#f6f7f9; --fg:#000; --muted:#444; --border:#ccc;
            --gain:#065f46; --loss:#9f1239; --accent:#1e293b; }
          .no-print { display: none !important; }
          table { page-break-inside: auto; } tr { page-break-inside: avoid; }
        }
      `}</style>
      <div className="root min-h-screen bg-[var(--bg)] text-[var(--fg)]" style={{ fontFamily: "ui-sans-serif,system-ui,-apple-system,Segoe UI,Roboto,sans-serif" }}>
        <div className="max-w-6xl mx-auto px-4 sm:px-6 py-6">
          {/* header */}
          <div className="flex items-center justify-between gap-3 flex-wrap">
            <div>
              <h1 className="text-xl font-semibold tracking-tight flex items-center gap-2">
                <Receipt size={20} className="text-[var(--accent)]" /> UK Capital Gains Ledger
              </h1>
              <p className="text-sm text-[var(--muted)] mt-0.5">Same-day · 30-day · Section 104 matching, with per-disposal audit trail. All figures GBP.</p>
            </div>
            <div className="flex items-center gap-2">
              {status && <span className="text-xs text-[var(--muted)] mr-1 max-w-[220px] text-right leading-tight">{status}</span>}
              <IconBtn onClick={exportJSON} title="Export / back up data (downloads a file, and copies to clipboard as a fallback)"><Download size={16} /></IconBtn>
              <IconBtn onClick={() => fileRef.current && fileRef.current.click()} title="Import a JSON file (replaces the ledger)"><Upload size={16} /></IconBtn>
              <input ref={fileRef} type="file" accept="application/json,.json" className="hidden" onChange={importJSON} />
              <IconBtn onClick={() => setDark((d) => !d)} title="Theme">{dark ? <Sun size={16} /> : <Moon size={16} />}</IconBtn>
            </div>
          </div>

          {/* tabs */}
          <div className="flex flex-wrap gap-1 mt-5 border-b border-[var(--border)]">
            {[["cgt", "CGT summary", TableProperties], ["holdings", "Holdings", Wallet], ["income", "Income", PoundSterling], ["planning", "Planning", TrendingUp], ["report", "Report", FileText], ["ledger", "Transactions", Receipt], ["whatif", "What-if sale", FlaskConical], ["import", "Import CSV", FileUp]].map(([k, label, Icon]) => (
              <button key={k} onClick={() => setTab(k)}
                className={"px-3 py-2 text-sm font-medium flex items-center gap-1.5 border-b-2 -mb-px transition " +
                  (tab === k ? "border-[var(--accent)] text-[var(--fg)]" : "border-transparent text-[var(--muted)] hover:text-[var(--fg)]")}>
                <Icon size={15} /> {label}
              </button>
            ))}
          </div>

          {error && (
            <div className="mt-4 flex items-start gap-2 text-sm rounded-lg px-3 py-2 text-[var(--loss)] border"
              style={{ background: "color-mix(in srgb, var(--loss) 12%, transparent)", borderColor: "color-mix(in srgb, var(--loss) 35%, transparent)" }}>
              <AlertTriangle size={16} className="mt-0.5 shrink-0" /> <span>{error}</span>
            </div>
          )}

          <div className="mt-5">
            {tab === "cgt" && <CgtTab {...{ taxYears, activeYear, setYear, yearDisposals, liab, income, setIncome, carried, setCarried, carryForward: allYears.carriedForward }} />}
            {tab === "income" && <IncomeTab {...{ incomeEntries, setIncomeEntries, eriEntries, setEriEntries, eriTxns, incomeByYear, income, setIncome, txns: giaTxns }} />}
            {tab === "holdings" && <HoldingsTab {...{ pools: matched.pools, prices, setPrices, avKey, setAvKey, avMeta, setAvMeta, priceMeta, setPriceMeta, txns: giaTxns, secMeta, setSecMeta }} />}
            {tab === "planning" && <PlanningTab {...{ pools: matched.pools, prices, setPrices, disposals: matched.disposals, txns: giaTxns, income }} />}
            {tab === "report" && <ReportTab {...{ taxYears, disposals: matched.disposals, income, carried }} />}
            {tab === "ledger" && <LedgerTab {...{ txns, setTxns }} />}
            {tab === "whatif" && <WhatIfTab {...{ pools: matched.pools, disposals: matched.disposals, income, carried, prices }} />}
            {tab === "import" && <ImportTab {...{ setTxns, setTab, setIncomeEntries, setEriEntries, secMeta }} />}
          </div>

          <p className="text-xs text-[var(--muted)] mt-8 leading-relaxed">
            Figures are an estimate to support your own filing, not tax advice. Verify before submitting to HMRC.
          </p>
        </div>
      </div>
    </div>
  );
}

/* ----------------------------- Income tab --------------------------- */
const addMonthsISO = (s, n) => {
  if (!s) return "";
  const [y, m, d] = s.split("-").map(Number);
  const target = new Date(Date.UTC(y, m - 1 + n, 1));
  const lastDayOfTarget = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();
  target.setUTCDate(Math.min(d, lastDayOfTarget));
  return target.toISOString().slice(0, 10);
};
const DIV_BLANK = () => ({ id: uid(), date: todayISO(), ticker: "", kind: "dividend", amount: "" });
const ERI_BLANK = () => ({ id: uid(), ticker: "", periodEnd: "", distributionDate: "", perShare: "", currency: "GBp", fxRate: 1, treatment: "dividend" });

function IncomeTab({ incomeEntries, setIncomeEntries, eriEntries, setEriEntries, eriTxns, incomeByYear, income, setIncome, txns }) {
  const [dv, setDv] = useState(DIV_BLANK());
  const [er, setEr] = useState(ERI_BLANK());
  const [fxBusy, setFxBusy] = useState(false);
  const years = Object.keys(incomeByYear).sort().reverse();

  const addDiv = () => { if (!dv.date || !(+dv.amount)) return; setIncomeEntries((p) => [...p, { ...dv, amount: +dv.amount }]); setDv(DIV_BLANK()); };
  const setEriF = (k, v) => setEr((e) => { const n = { ...e, [k]: v }; if (k === "periodEnd") n.distributionDate = addMonthsISO(v, 6); if (k === "currency" && (v === "GBP" || v === "GBp")) n.fxRate = 1; return n; });
  const addEri = () => { if (!er.ticker || !er.periodEnd || !er.distributionDate || !(+er.perShare)) return; setEriEntries((p) => [...p, { ...er, ticker: er.ticker.toUpperCase(), perShare: +er.perShare, fxRate: +er.fxRate || 0 }]); setEr(ERI_BLANK()); };
  const fetchEriFx = async () => {
    if (er.currency === "GBP" || er.currency === "GBp") return;
    setFxBusy(true);
    try { const fx = await fxToGBP(er.currency); if (fx) setEr((e) => ({ ...e, fxRate: +fx.toFixed(6) })); } catch { /* ignore */ }
    setFxBusy(false);
  };
  const eriPreview = (() => {
    const units = unitsHeldAt(txns, er.periodEnd || "9999-12-31", er.ticker);
    const native = units * (+er.perShare || 0);
    const g = er.currency === "GBp" ? native / 100 : er.currency === "GBP" ? native : native * (+er.fxRate || 0);
    return { units, g };
  })();

  return (
    <div className="space-y-6">
      <div className="flex items-end gap-3 flex-wrap">
        <Field label="Employment / other income (£)"><input type="number" value={income} onChange={(e) => setIncome(+e.target.value || 0)} className="input num w-48" /></Field>
        <p className="text-xs text-[var(--muted)] pb-2 max-w-md">Dividends and interest here are stacked on top of this income. Only taxable (GIA) income belongs here — anything inside an ISA or pension is tax-free and should be excluded.</p>
      </div>

      {/* Per-year income tax */}
      {years.length ? (
        <div className="space-y-2">
          <h3 className="font-semibold text-sm">Investment income tax by year</h3>
          <div className="rounded-xl border border-[var(--border)] overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-[var(--panel2)] text-[var(--muted)]">
                <tr>{["Tax year", "Dividends", "Interest", "Dividend tax", "Interest tax", "Total"].map((h, i) => <th key={i} className={"py-2 px-3 font-medium " + (i ? "text-right" : "text-left")}>{h}</th>)}</tr>
              </thead>
              <tbody className="divide-y divide-[var(--border)]">
                {years.map((y) => {
                  const d = incomeByYear[y], r = investmentIncomeTax({ salary: income, interest: d.interest, dividends: d.dividends, year: y });
                  return (
                    <tr key={y}>
                      <td className="py-2 px-3 font-medium">{y}{r.assumed ? " *" : ""}</td>
                      <td className="py-2 px-3 text-right num">{gbp(d.dividends)}</td>
                      <td className="py-2 px-3 text-right num">{gbp(d.interest)}</td>
                      <td className="py-2 px-3 text-right num">{gbp(r.dividendTax)}</td>
                      <td className="py-2 px-3 text-right num">{gbp(r.interestTax)}</td>
                      <td className="py-2 px-3 text-right num font-semibold text-[var(--loss)]">{gbp(r.tax)}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <p className="text-xs text-[var(--muted)]">Dividend allowance and Personal Savings Allowance are applied automatically by year and band. Figures marked * use assumed (latest) rates for years not in the table.</p>
        </div>
      ) : <Empty msg="No dividends, interest or ERI recorded yet. Add them below to see the income-tax position." />}

      {/* Dividend / interest ledger */}
      <div className="space-y-2">
        <h3 className="font-semibold text-sm">Dividends & interest</h3>
        <div className="flex items-end gap-2 flex-wrap rounded-xl border border-[var(--border)] bg-[var(--panel)] p-3">
          <Field label="Date"><input type="date" value={dv.date} onChange={(e) => setDv({ ...dv, date: e.target.value })} className="input num" /></Field>
          <Field label="Ticker (optional)"><input value={dv.ticker} onChange={(e) => setDv({ ...dv, ticker: e.target.value.toUpperCase() })} className="input num w-24" placeholder="—" /></Field>
          <Field label="Type"><select value={dv.kind} onChange={(e) => setDv({ ...dv, kind: e.target.value })} className="input"><option value="dividend">Dividend</option><option value="interest">Interest</option></select></Field>
          <Field label="Amount (£, GBP)"><input type="number" value={dv.amount} onChange={(e) => setDv({ ...dv, amount: e.target.value })} className="input num w-32" placeholder="0.00" /></Field>
          <button onClick={addDiv} className="btn-accent"><Plus size={15} /> Add</button>
        </div>
        {incomeEntries.length > 0 && (
          <div className="rounded-xl border border-[var(--border)] overflow-hidden">
            <table className="w-full text-sm">
              <tbody className="divide-y divide-[var(--border)]">
                {incomeEntries.slice().sort((a, b) => (a.date < b.date ? 1 : -1)).map((e) => (
                  <tr key={e.id}>
                    <td className="py-2 px-3 num text-[var(--muted)]">{e.date}</td>
                    <td className="py-2 px-3">{e.ticker || "—"}</td>
                    <td className="py-2 px-3 capitalize">{e.kind}</td>
                    <td className="py-2 px-3 num">{ukTaxYear(e.date)}</td>
                    <td className="py-2 px-3 text-right num">{gbp(+e.amount)}</td>
                    <td className="py-2 px-3 text-right"><button onClick={() => setIncomeEntries((p) => p.filter((x) => x.id !== e.id))} className="text-[var(--muted)] hover:text-[var(--loss)]"><Trash2 size={15} /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* ERI */}
      <div className="space-y-2">
        <h3 className="font-semibold text-sm">Excess reportable income (offshore reporting funds)</h3>
        <p className="text-xs text-[var(--muted)] max-w-3xl">For accumulating ETFs and other offshore reporting funds. Enter the reportable income per share from the fund's report and its reporting-period end. It's taxed on the fund distribution date (period end + 6 months), in that tax year, as dividend (equity funds) or interest (bond funds &gt;60% debt) — and the taxed amount is added to the Section 104 pool, lowering the gain on later disposals.</p>
        <div className="grid gap-2 rounded-xl border border-[var(--border)] bg-[var(--panel)] p-3" style={{ gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))" }}>
          <Field label="Ticker"><input value={er.ticker} onChange={(e) => setEriF("ticker", e.target.value.toUpperCase())} className="input num w-full" placeholder="e.g. XNAQ" /></Field>
          <Field label="Reporting period end"><input type="date" value={er.periodEnd} onChange={(e) => setEriF("periodEnd", e.target.value)} className="input num w-full" /></Field>
          <Field label="Fund distribution date"><input type="date" value={er.distributionDate} onChange={(e) => setEriF("distributionDate", e.target.value)} className="input num w-full" /></Field>
          <Field label="Reportable income / share"><input type="number" value={er.perShare} onChange={(e) => setEriF("perShare", e.target.value)} className="input num w-full" placeholder="0.00" /></Field>
          <Field label="Currency"><select value={er.currency} onChange={(e) => setEriF("currency", e.target.value)} className="input w-full">{["GBp", "GBP", "USD", "EUR"].map((c) => <option key={c} value={c}>{c}</option>)}</select></Field>
          <Field label="FX → GBP">
            <div className="flex gap-1"><input type="number" value={er.fxRate} onChange={(e) => setEriF("fxRate", e.target.value)} disabled={er.currency === "GBP" || er.currency === "GBp"} className="input num w-full disabled:opacity-50" />
              {er.currency !== "GBP" && er.currency !== "GBp" && <button onClick={fetchEriFx} disabled={fxBusy} className="text-[var(--accent)] px-1" title="Fetch latest FX">{fxBusy ? "…" : "↻"}</button>}</div>
          </Field>
          <Field label="Taxed as"><select value={er.treatment} onChange={(e) => setEriF("treatment", e.target.value)} className="input w-full"><option value="dividend">Dividend</option><option value="interest">Interest</option></select></Field>
          <div className="flex items-end"><button onClick={addEri} className="btn-accent w-full justify-center"><Plus size={15} /> Add</button></div>
        </div>
        {er.ticker && er.periodEnd && (
          <p className="text-xs text-[var(--muted)] num">Preview: {num(eriPreview.units, eriPreview.units % 1 ? 4 : 0)} units held at {er.periodEnd} → ERI {gbp(eriPreview.g || 0)} taxed in {er.distributionDate ? ukTaxYear(er.distributionDate) : "—"}.</p>
        )}
        {eriEntries.length > 0 && (
          <div className="rounded-xl border border-[var(--border)] overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-[var(--panel2)] text-[var(--muted)]">
                <tr>{["Fund", "Period end", "Dist. date", "Units", "ERI (GBP)", "Taxed as", "Tax year", ""].map((h, i) => <th key={i} className={"py-2 px-3 font-medium " + (i && i < 7 ? "text-right" : "text-left")}>{h}</th>)}</tr>
              </thead>
              <tbody className="divide-y divide-[var(--border)]">
                {eriEntries.slice().sort((a, b) => (a.distributionDate < b.distributionDate ? 1 : -1)).map((e) => {
                  const t = eriTxns.find((x) => x.id === "eri-" + e.id);
                  return (
                    <tr key={e.id}>
                      <td className="py-2 px-3 font-medium">{e.ticker}</td>
                      <td className="py-2 px-3 num text-[var(--muted)]">{e.periodEnd}</td>
                      <td className="py-2 px-3 num text-[var(--muted)]">{e.distributionDate}</td>
                      <td className="py-2 px-3 text-right num">{t ? num(t._units, t._units % 1 ? 4 : 0) : "—"}</td>
                      <td className="py-2 px-3 text-right num">{gbp(t ? t._gbp : 0)}</td>
                      <td className="py-2 px-3 capitalize">{e.treatment}</td>
                      <td className="py-2 px-3 num">{ukTaxYear(e.distributionDate)}</td>
                      <td className="py-2 px-3 text-right"><button onClick={() => setEriEntries((p) => p.filter((x) => x.id !== e.id))} className="text-[var(--muted)] hover:text-[var(--loss)]"><Trash2 size={15} /></button></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
        <p className="text-xs text-[var(--muted)]">The base-cost uplift shows up automatically in Holdings and the CGT summary. ERI is added to the pool on the distribution date, so it only reduces gains on disposals after that date.</p>
      </div>
    </div>
  );
}

/* ----------------------------- CGT tab ------------------------------ */
function CgtTab({ taxYears, activeYear, setYear, yearDisposals, liab, income, setIncome, carried, setCarried, carryForward }) {
  if (!taxYears.length) return <Empty msg="No disposals yet. Add or import transactions to see a CGT position." />;
  return (
    <div className="space-y-5">
      <div className="flex items-end gap-3 flex-wrap">
        <Field label="Tax year">
          <select value={activeYear} onChange={(e) => setYear(e.target.value)} className="input num">
            {taxYears.map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
        </Field>
        <Field label="Annual income before tax (£)"><input type="number" value={income} onChange={(e) => setIncome(+e.target.value || 0)} className="input num w-44" /></Field>
        <Field label="Losses b/f (before tracked years)"><input type="number" value={carried} onChange={(e) => setCarried(+e.target.value || 0)} className="input num w-52" /></Field>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <Stat label="Net gains" value={gbp(liab.net)} tone={liab.net >= 0 ? "gain" : "loss"} />
        <Stat label="Taxable after AEA" value={gbp(liab.taxable)} />
        <Stat label="CGT due" value={gbp(liab.tax)} tone="loss" big />
        <Stat label="Reporting" value={liab.reporting ? "Required" : "Not required"} sub={liab.reporting ? "tax due or proceeds over threshold" : "below thresholds"} />
      </div>

      <div className="text-xs text-[var(--muted)] num">
        Gains {gbp(liab.gains)} · losses {gbp(liab.losses)} · AEA {gbp(liab.aea)}{liab.usedCarried ? ` · carried losses used ${gbp(liab.usedCarried)}` : ""} ·
        {" "}{liab.breakdown.length ? liab.breakdown.map((b) => `${gbp(b.amount)} @ ${fmtRate(b.rate)}`).join(" + ") : "no taxable gain"} · proceeds {gbp(liab.proceeds)}
        {liab.assumed ? " · rates assumed (year not in table)" : ""}
      </div>
      <div className="text-xs text-[var(--muted)] num -mt-3">
        Income {gbp(income)} − personal allowance {gbp(liab.personalAllowance)} = taxable income {gbp(liab.taxableIncome)}; basic-rate band left for gains {gbp(Math.max(0, cfgFor(activeYear).basicLimit - liab.taxableIncome))}.
      </div>
      <div className="rounded-lg border border-[var(--border)] bg-[var(--panel2)] px-3 py-2 text-xs text-[var(--muted)] num -mt-1">
        <span className="font-medium text-[var(--fg)]">Loss pool</span> — brought into {activeYear} {gbp(liab.carriedInto ?? carried)}
        {liab.usedCarried ? ` · used ${gbp(liab.usedCarried)} (only down to the AEA)` : " · none used"}
        {liab.inYearNetLoss ? ` · net loss realised this year ${gbp(liab.inYearNetLoss)}` : ""} · carried out {gbp(liab.carriedOut ?? carried)}.
        {" "}Total unused losses across all tracked years: <span className="font-medium text-[var(--fg)]">{gbp(carryForward || 0)}</span>.
        {carryForward > 0 ? " Remember losses must be claimed within 4 years of the tax year they arose." : ""}
      </div>

      {/* audit trail — the signature element */}
      <div className="space-y-3">
        {yearDisposals.map((d) => (
          <div key={d.id} className="rounded-xl border border-[var(--border)] bg-[var(--panel)] overflow-hidden">
            <div className="flex items-center justify-between px-4 py-2.5 bg-[var(--panel2)] border-b border-[var(--border)]">
              <div className="flex items-baseline gap-3">
                <span className="font-semibold">{d.ticker}</span>
                <span className="text-sm text-[var(--muted)] num">{d.date} · sold {num(d.quantity, d.quantity % 1 ? 4 : 0)} · proceeds {gbp(d.proceeds)}</span>
              </div>
              <span className={"num font-semibold " + (d.gain >= 0 ? "text-[var(--gain)]" : "text-[var(--loss)]")}>{gbp(d.gain)}</span>
            </div>
            <div className="divide-y divide-[var(--border)]">
              {d.legs.map((l, i) => (
                <div key={i} className="grid grid-cols-12 gap-2 px-4 py-2 text-sm items-center">
                  <div className="col-span-3"><MethodChip m={l.method} /></div>
                  <div className="col-span-2 num text-[var(--muted)]">{num(l.quantity, l.quantity % 1 ? 4 : 0)} sh{l.matchedAcqDate ? "" : ""}</div>
                  <div className="col-span-3 num text-right">cost {gbp(l.cost)}</div>
                  <div className="col-span-2 num text-right text-[var(--muted)]">{gbp(l.proceeds)}</div>
                  <div className={"col-span-2 num text-right font-medium " + (l.gain >= 0 ? "text-[var(--gain)]" : "text-[var(--loss)]")}>{gbp(l.gain)}</div>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* --------------------------- Ledger tab ----------------------------- */
const BLANK = () => ({ id: uid(), date: todayISO(), ticker: "", side: "BUY", quantity: "", nativeCurrency: "GBP", nativeAmount: "", fxRate: 1, gbpAmount: "", wrapper: "GIA", note: "" });
const WRAPPERS = ["GIA", "ISA", "SIPP", "LISA"];
function LedgerTab({ txns, setTxns }) {
  const [draft, setDraft] = useState(BLANK());
  const [fxBusy, setFxBusy] = useState(false);

  const set = (k, v) => setDraft((d) => {
    const next = { ...d, [k]: v };
    if (["nativeAmount", "fxRate"].includes(k)) {
      const na = +next.nativeAmount || 0, fx = +next.fxRate || 0;
      if (na && fx) next.gbpAmount = +(na * fx).toFixed(2);
    }
    if (k === "nativeCurrency" && v === "GBP") { next.fxRate = 1; if (next.nativeAmount) next.gbpAmount = +next.nativeAmount; }
    return next;
  });

  const fetchFx = async () => {
    if (draft.nativeCurrency === "GBP") return;
    setFxBusy(true);
    try {
      const res = await fetch(`https://api.frankfurter.dev/v1/${draft.date}?from=${draft.nativeCurrency}&to=GBP`);
      const j = await res.json();
      const rate = j?.rates?.GBP;
      if (rate) set("fxRate", +rate.toFixed(6));
    } catch { /* offline / blocked — keep manual */ }
    setFxBusy(false);
  };

  const add = () => {
    if (!draft.ticker || !draft.date || !(+draft.quantity > 0)) return;
    const t = { ...draft, ticker: draft.ticker.toUpperCase().trim(), quantity: +draft.quantity, nativeAmount: +draft.nativeAmount || 0, fxRate: +draft.fxRate || 1, gbpAmount: +draft.gbpAmount || 0 };
    setTxns((p) => [...p, t]); setDraft(BLANK());
  };
  const rows = [...txns].sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));

  return (
    <div className="space-y-4">
      {/* add row */}
      <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] p-3">
        <div className="grid grid-cols-2 sm:grid-cols-9 gap-2 items-end">
          <Field label="Date"><input type="date" value={draft.date} onChange={(e) => set("date", e.target.value)} className="input num w-full" /></Field>
          <Field label="Ticker"><input value={draft.ticker} onChange={(e) => set("ticker", e.target.value)} placeholder="WFC" className="input w-full" /></Field>
          <Field label="Side">
            <select value={draft.side} onChange={(e) => set("side", e.target.value)} className="input w-full"><option>BUY</option><option>SELL</option></select>
          </Field>
          <Field label="Wrapper">
            <select value={draft.wrapper} onChange={(e) => set("wrapper", e.target.value)} className="input w-full">{WRAPPERS.map((w) => <option key={w}>{w}</option>)}</select>
          </Field>
          <Field label="Quantity"><input type="number" value={draft.quantity} onChange={(e) => set("quantity", e.target.value)} className="input num w-full" /></Field>
          <Field label="Ccy">
            <select value={draft.nativeCurrency} onChange={(e) => set("nativeCurrency", e.target.value)} className="input w-full">
              {["GBP", "USD", "EUR", "CHF"].map((c) => <option key={c}>{c}</option>)}
            </select>
          </Field>
          <Field label="Native amount"><input type="number" value={draft.nativeAmount} onChange={(e) => set("nativeAmount", e.target.value)} className="input num w-full" /></Field>
          <Field label={<span className="flex items-center gap-1">FX→GBP {draft.nativeCurrency !== "GBP" && <button onClick={fetchFx} title="Fetch ECB rate for date" className="text-[var(--accent)]">{fxBusy ? <RefreshCw size={12} className="animate-spin" /> : <Wand2 size={12} />}</button>}</span>}>
            <input type="number" value={draft.fxRate} onChange={(e) => set("fxRate", e.target.value)} disabled={draft.nativeCurrency === "GBP"} className="input num w-full disabled:opacity-50" />
          </Field>
          <Field label="GBP amount"><input type="number" value={draft.gbpAmount} onChange={(e) => set("gbpAmount", e.target.value)} className="input num w-full" /></Field>
        </div>
        <div className="flex items-center justify-between mt-2">
          <span className="text-xs text-[var(--muted)]">{draft.nativeCurrency !== "GBP" ? "GBP auto-computes from native × rate; both stay editable." : "GBP transaction — rate fixed at 1."}</span>
          <button onClick={add} className="btn-accent"><Plus size={15} /> Add transaction</button>
        </div>
      </div>

      {/* table */}
      <div className="rounded-xl border border-[var(--border)] overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-[var(--panel2)] text-[var(--muted)] text-xs uppercase tracking-wide">
            <tr>{["Date", "Ticker", "Side", "Wrapper", "Qty", "Native", "FX", "GBP", ""].map((h, i) => <th key={i} className={"px-3 py-2 font-medium " + (i >= 4 && i <= 7 ? "text-right" : "text-left")}>{h}</th>)}</tr>
          </thead>
          <tbody className="divide-y divide-[var(--border)] bg-[var(--panel)]">
            {rows.map((t) => (
              <tr key={t.id} className="hover:bg-[var(--panel2)]">
                <td className="px-3 py-2 num">{t.date}</td>
                <td className="px-3 py-2 font-medium">{t.ticker}</td>
                <td className="px-3 py-2"><span className={"text-xs font-semibold " + (t.side === "BUY" ? "text-[var(--gain)]" : "text-[var(--loss)]")}>{t.side}</span></td>
                <td className="px-3 py-2"><span className={"text-[10px] font-semibold px-1.5 py-0.5 rounded " + ((t.wrapper || "GIA") === "GIA" ? "bg-[var(--chip)] text-[var(--fg)]" : "bg-[color:color-mix(in_srgb,var(--gain)_18%,transparent)] text-[var(--gain)]")}>{t.wrapper || "GIA"}</span></td>
                <td className="px-3 py-2 num text-right">{num(t.quantity, t.quantity % 1 ? 4 : 0)}</td>
                <td className="px-3 py-2 num text-right text-[var(--muted)]">{t.nativeCurrency === "GBP" ? "—" : `${num(t.nativeAmount)} ${t.nativeCurrency}`}</td>
                <td className="px-3 py-2 num text-right text-[var(--muted)]">{t.nativeCurrency === "GBP" ? "—" : num(t.fxRate, 4)}</td>
                <td className="px-3 py-2 num text-right">{gbp(t.gbpAmount)}</td>
                <td className="px-3 py-2 text-right"><button onClick={() => setTxns((p) => p.filter((x) => x.id !== t.id))} className="text-[var(--muted)] hover:text-[var(--loss)]"><Trash2 size={15} /></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ----------------------- Live prices (Alpha Vantage) ---------------- */
function LivePricesPanel({ tickers, avKey, setAvKey, avMeta, setAvMeta, prices, setPrices, priceMeta, setPriceMeta, txns }) {
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [prog, setProg] = useState("");
  const [msg, setMsg] = useState("");

  const ledgerCcy = useMemo(() => {
    const m = {}; for (const t of txns) if (!m[t.ticker] && t.nativeCurrency) m[t.ticker] = t.nativeCurrency; return m;
  }, [txns]);
  const defYahoo = (tk) => (ledgerCcy[tk] === "GBP" ? `${tk}.L` : tk);   // Yahoo LSE suffix = .L
  const defAv = (tk) => (ledgerCcy[tk] === "GBP" ? `${tk}.LON` : tk);    // Alpha Vantage LSE suffix = .LON
  const defCcy = (tk) => (ledgerCcy[tk] === "USD" ? "USD" : ledgerCcy[tk] === "EUR" ? "EUR" : "GBp");
  const meta = (tk) => ({
    yahoo: avMeta[tk]?.yahoo ?? defYahoo(tk),
    av: avMeta[tk]?.av ?? avMeta[tk]?.symbol ?? defAv(tk),
    currency: avMeta[tk]?.currency ?? defCcy(tk),
  });
  const setMeta = (tk, patch) => setAvMeta((m) => ({ ...m, [tk]: { ...meta(tk), ...patch } }));
  const used = avBudget().n;

  const applyQuote = (tk, raw, ccy, fx, source) => {
    const g = toGBP(raw, ccy, fx);
    if (g == null) { setMsg(`${tk}: couldn't convert ${ccy} to GBP`); return false; }
    setPrices((p) => ({ ...p, [tk]: +g.toFixed(4) }));
    setPriceMeta((p) => ({ ...p, [tk]: { asOf: new Date().toISOString(), raw, ccy, source } }));
    return true;
  };
  const yahooFetch = async (syms) => {
    const r = await fetch(`/api/quotes?symbols=${encodeURIComponent(syms.join(","))}`);
    if (!r.ok) throw new Error(`function ${r.status}`);
    const j = await r.json();
    const by = {}; (j.quotes || []).forEach((q) => { by[q.symbol] = q; });
    return by;
  };

  const fetchOne = async (tk) => {
    setBusy(true); setProg(`Fetching ${tk}...`); setMsg("");
    const m = meta(tk);
    try {
      const q = (await yahooFetch([m.yahoo]))[m.yahoo];
      if (q && q.price != null) {
        const fx = await fxToGBP(q.currency);
        if (applyQuote(tk, q.price, q.currency, fx, "Yahoo")) { setMsg(`${tk}: ${num(q.price, 2)} ${q.currency} to ${gbp(toGBP(q.price, q.currency, fx))} (Yahoo)`); setBusy(false); setProg(""); return; }
      }
    } catch { /* fall through to AV */ }
    if (avKey && avBudget().n < 25) {
      try {
        const raw = await avQuote(m.av, avKey); avBump();
        const fx = await fxToGBP(m.currency);
        if (applyQuote(tk, raw, m.currency, fx, "AV")) { setMsg(`${tk}: ${num(raw, 2)} ${m.currency} to ${gbp(toGBP(raw, m.currency, fx))} (Alpha Vantage)`); setBusy(false); setProg(""); return; }
      } catch (e) { setMsg(`${tk}: ${e.message}`); setBusy(false); setProg(""); return; }
    }
    setMsg(`${tk}: no live price (deploy the Yahoo function${avKey ? "" : "; no AV key set"}) - enter manually.`);
    setBusy(false); setProg("");
  };

  const fetchAll = async () => {
    setBusy(true); setMsg(""); const done = {}; const fxCache = {};
    const getFx = async (ccy) => { if (ccy === "GBP" || ccy === "GBp") return 1; if (!(ccy in fxCache)) fxCache[ccy] = await fxToGBP(ccy); return fxCache[ccy]; };
    try {
      setProg("Fetching from Yahoo...");
      const by = await yahooFetch(tickers.map((tk) => meta(tk).yahoo));
      for (const tk of tickers) { const q = by[meta(tk).yahoo]; if (q && q.price != null) { const fx = await getFx(q.currency); if (applyQuote(tk, q.price, q.currency, fx, "Yahoo")) done[tk] = true; } }
    } catch { setMsg("Yahoo function unreachable - trying Alpha Vantage fallback."); }
    const rest = tickers.filter((tk) => !done[tk]);
    if (rest.length && avKey) {
      for (let i = 0; i < rest.length; i++) {
        if (avBudget().n >= 25) { setMsg("Alpha Vantage daily limit reached - enter the rest manually."); break; }
        const tk = rest[i], m = meta(tk); setProg(`Alpha Vantage fallback ${i + 1}/${rest.length}: ${tk}...`);
        try { const raw = await avQuote(m.av, avKey); avBump(); const fx = await getFx(m.currency); if (applyQuote(tk, raw, m.currency, fx, "AV")) done[tk] = true; }
        catch (e) { if (/limit/i.test(e.message)) { setMsg("Alpha Vantage limit reached - stopping."); break; } }
        if (i < rest.length - 1) { setProg("Waiting (AV 5/min)..."); await sleep(13000); }
      }
    }
    const got = Object.keys(done).length;
    setProg(""); setMsg(`Updated ${got}/${tickers.length} prices${got < tickers.length ? " - enter the rest manually." : "."}`);
    setBusy(false);
  };

  return (
    <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)]">
      <button onClick={() => setOpen((o) => !o)} className="w-full flex items-center justify-between px-4 py-2.5 text-sm">
        <span className="font-medium flex items-center gap-2"><RefreshCw size={14} className="text-[var(--accent)]" /> Live prices <span className="text-xs font-normal text-[var(--muted)]">- Yahoo then Alpha Vantage then manual</span></span>
        <span className="text-xs text-[var(--muted)]">{open ? "hide" : "set up"}</span>
      </button>
      {open && (
        <div className="px-4 pb-4 space-y-3 border-t border-[var(--border)] pt-3">
          <div className="flex items-end gap-2 flex-wrap">
            <button onClick={fetchAll} disabled={busy} className="btn-accent disabled:opacity-50"><RefreshCw size={15} className={busy ? "animate-spin" : ""} /> Fetch prices</button>
            {(prog || msg) && <span className="text-xs text-[var(--muted)] pb-2">{prog || msg}</span>}
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead className="text-[var(--muted)]">
                <tr>{["Ticker", "Yahoo symbol", "AV symbol", "Ccy (AV)", "", "Last quote", "Source", "As of"].map((h, i) => <th key={i} className="py-1 px-2 font-medium text-left">{h}</th>)}</tr>
              </thead>
              <tbody>
                {tickers.map((tk) => {
                  const m = meta(tk), pm = priceMeta[tk];
                  return (
                    <tr key={tk} className="border-t border-[var(--border)]">
                      <td className="py-1 px-2 font-medium">{tk}</td>
                      <td className="py-1 px-2"><input value={m.yahoo} onChange={(e) => setMeta(tk, { yahoo: e.target.value.trim() })} className="input num w-24 py-0.5" /></td>
                      <td className="py-1 px-2"><input value={m.av} onChange={(e) => setMeta(tk, { av: e.target.value.trim() })} className="input num w-24 py-0.5" /></td>
                      <td className="py-1 px-2">
                        <select value={m.currency} onChange={(e) => setMeta(tk, { currency: e.target.value })} className="input py-0.5">
                          {["GBp", "GBP", "USD", "EUR"].map((c) => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </td>
                      <td className="py-1 px-2"><button onClick={() => fetchOne(tk)} disabled={busy} className="text-[var(--accent)] disabled:opacity-40" title="Fetch this one">&#8635;</button></td>
                      <td className="py-1 px-2 num text-[var(--muted)]">{pm ? `${num(pm.raw, 2)} ${pm.ccy}` : "-"}</td>
                      <td className="py-1 px-2">{pm?.source ? <span className="text-[10px] px-1.5 py-0.5 rounded-full" style={{ color: pm.source === "Yahoo" ? "var(--m-pool)" : "var(--m-bb)", background: "var(--chip)" }}>{pm.source}</span> : <span className="text-[var(--muted)]">-</span>}</td>
                      <td className="py-1 px-2 num text-[var(--muted)]">{pm ? new Date(pm.asOf).toLocaleString("en-GB", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" }) : "-"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          <details className="text-xs">
            <summary className="cursor-pointer text-[var(--muted)]">Alpha Vantage fallback key ({used}/25 used today)</summary>
            <div className="mt-2">
              <Field label="Alpha Vantage key - used only if Yahoo fails (saved on this device)">
                <input type="password" value={avKey} onChange={(e) => setAvKey(e.target.value.trim())} placeholder="paste your Alpha Vantage key" className="input num w-64" />
              </Field>
            </div>
          </details>

          <p className="text-xs text-[var(--muted)] leading-relaxed">
            Yahoo is primary - it returns each quote's currency, so GBP normalisation is automatic (pence /100, USD/EUR via ECB rates) with no daily cap. It needs the <span className="font-medium">/api/quotes</span> serverless function deployed (LSE symbols use the <span className="font-medium">.L</span> suffix). If Yahoo is down or misses a symbol, Alpha Vantage fills in silently using the AV symbol (<span className="font-medium">.LON</span>) and the currency you set per line, capped at 25 calls/day. Anything neither can price, you enter by hand. Check "Last quote" against a price you know if a value looks off.
          </p>
        </div>
      )}
    </div>
  );
}

/* --------------------------- Holdings tab --------------------------- */
function HoldingsTab({ pools, prices, setPrices, avKey, setAvKey, avMeta, setAvMeta, priceMeta, setPriceMeta, txns, secMeta, setSecMeta }) {
  const tickers = Object.keys(pools).filter((t) => pools[t].qty > 1e-6).sort();
  if (!tickers.length) return <Empty msg="No open holdings yet. Add buy transactions to see your positions and unrealised gains." />;

  const setISIN = (tk, v) => setSecMeta((m) => ({ ...m, [tk]: { ...m[tk], isin: v.toUpperCase().trim() } }));

  const rows = tickers.map((tk) => {
    const { qty, cost } = pools[tk];
    const avg = qty ? cost / qty : 0;
    const price = prices[tk] ?? "";
    const hasP = price !== "" && !isNaN(+price);
    const value = hasP ? qty * +price : null;
    const unreal = hasP ? value - cost : null;
    return { tk, qty, cost, avg, price, value, unreal, pct: hasP && cost ? (unreal / cost) * 100 : null, sec: secMeta[tk] || {} };
  });
  const priced = rows.filter((r) => r.value != null);
  const totCost = priced.reduce((s, r) => s + r.cost, 0);
  const totValue = priced.reduce((s, r) => s + r.value, 0);
  const totUnreal = totValue - totCost;
  const missingIsin = rows.filter((r) => !r.sec.isin).length;

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <Stat label="Open pool cost" value={gbp(rows.reduce((s, r) => s + r.cost, 0))} />
        <Stat label="Market value (priced)" value={priced.length ? gbp(totValue) : "—"} sub={priced.length < rows.length ? `${priced.length}/${rows.length} priced` : "all priced"} />
        <Stat label="Unrealised gain" value={priced.length ? gbp(totUnreal) : "—"} tone={totUnreal >= 0 ? "gain" : "loss"} big />
        <Stat label="Unrealised %" value={priced.length && totCost ? `${totUnreal >= 0 ? "+" : ""}${num((totUnreal / totCost) * 100)}%` : "—"} tone={totUnreal >= 0 ? "gain" : "loss"} />
      </div>

      <LivePricesPanel {...{ tickers, avKey, setAvKey, avMeta, setAvMeta, prices, setPrices, priceMeta, setPriceMeta, txns }} />

      <div className="rounded-xl border border-[var(--border)] overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-[var(--panel2)] text-[var(--muted)] text-xs uppercase tracking-wide">
            <tr>{["Ticker", "ISIN", "Quantity", "Avg cost", "Pool cost", "Price now", "Market value", "Unrealised", "%"].map((h, i) => (
              <th key={i} className={"px-3 py-2 font-medium " + (i <= 1 ? "text-left" : "text-right")}>{h}</th>
            ))}</tr>
          </thead>
          <tbody className="divide-y divide-[var(--border)] bg-[var(--panel)]">
            {rows.map((r) => (
              <tr key={r.tk} className="hover:bg-[var(--panel2)]">
                <td className="px-3 py-2 font-medium">
                  {r.tk}
                  {r.sec.eri === true && <span title="Offshore reporting fund — generates excess reportable income (ERI) while held unsheltered" className="ml-1.5 text-[10px] font-semibold px-1.5 py-0.5 rounded bg-[color:color-mix(in_srgb,var(--m-bb)_18%,transparent)] text-[var(--m-bb)] align-middle">ERI</span>}
                </td>
                <td className="px-3 py-2">
                  <input value={r.sec.isin || ""} onChange={(e) => setISIN(r.tk, e.target.value)} placeholder="IE00…" className="input font-mono text-xs w-36 py-1" />
                </td>
                <td className="px-3 py-2 num text-right">{num(r.qty, r.qty % 1 ? 4 : 0)}</td>
                <td className="px-3 py-2 num text-right text-[var(--muted)]">{gbp(r.avg)}</td>
                <td className="px-3 py-2 num text-right">{gbp(r.cost)}</td>
                <td className="px-3 py-2 text-right">
                  <input type="number" value={r.price} placeholder="—"
                    onChange={(e) => setPrices((p) => ({ ...p, [r.tk]: e.target.value === "" ? undefined : +e.target.value }))}
                    className="input num w-24 text-right py-1" />
                </td>
                <td className="px-3 py-2 num text-right">{r.value != null ? gbp(r.value) : "—"}</td>
                <td className={"px-3 py-2 num text-right font-medium " + (r.unreal == null ? "text-[var(--muted)]" : r.unreal >= 0 ? "text-[var(--gain)]" : "text-[var(--loss)]")}>{r.unreal != null ? gbp(r.unreal) : "—"}</td>
                <td className={"px-3 py-2 num text-right " + (r.pct == null ? "text-[var(--muted)]" : r.pct >= 0 ? "text-[var(--gain)]" : "text-[var(--loss)]")}>{r.pct != null ? `${r.pct >= 0 ? "+" : ""}${num(r.pct)}%` : "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <p className="text-xs text-[var(--muted)]">
        Enter each holding's current price per share in GBP, or fetch live prices above. Prices save locally on your device.
        Unrealised gain = current value − Section 104 pool cost; it's an indicator, not a taxable event until you sell.
        {missingIsin > 0 && ` ISIN is set for ${rows.length - missingIsin}/${rows.length} holdings — it's the join key for matching issuer ERI reports, so fill in the rest when you get the chance.`}
        {" "}The <span className="text-[var(--m-bb)] font-semibold">ERI</span> badge flags offshore reporting funds; UK investment trusts pay ordinary dividends instead and won't show it.
      </p>
    </div>
  );
}

/* --------------------------- Planning tab --------------------------- */
function PlanningTab({ pools, prices, setPrices, disposals, txns, income }) {
  const yearNow = ukTaxYear(todayISO());
  const aea = aeaForYear(yearNow);
  const realised = disposals.filter((d) => d.taxYear === yearNow);
  const realisedNet = realised.reduce((s, d) => s + d.gain, 0);
  const headroom = Math.max(0, aea - realisedNet); // gains realisable tax-free this year
  const tickers = Object.keys(pools).filter((t) => pools[t].qty > 1e-6).sort();

  // 30-day forward warning: buys of the same ticker within the last 30 days.
  const today = new Date(todayISO());
  const recentBuys = {};
  for (const t of txns) {
    if (t.side !== "BUY") continue;
    const days = (today - new Date(t.date)) / 86400000;
    if (days >= 0 && days <= 30) recentBuys[t.ticker] = (recentBuys[t.ticker] || 0) + (+t.quantity);
  }
  // past disposals that were matched under the 30-day rule
  const pastBB = disposals.filter((d) => d.legs.some((l) => l.method === "THIRTY_DAY"));

  const rows = tickers.map((tk) => {
    const { qty, cost } = pools[tk];
    const avg = qty ? cost / qty : 0;
    const price = prices[tk];
    const hasP = price != null && price !== "" && !isNaN(+price);
    const perShare = hasP ? +price - avg : null;
    const maxShares = hasP && perShare > 0 ? Math.min(qty, Math.floor(headroom / perShare)) : null;
    const gainIf = maxShares != null ? maxShares * perShare : null;
    const unreal = hasP ? qty * +price - cost : null;
    return { tk, qty, avg, price: hasP ? price : "", perShare, maxShares, gainIf, unreal, recentBuy: recentBuys[tk] };
  });

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <Stat label={`AEA ${yearNow}`} value={gbp(aea).replace(".00", "")} />
        <Stat label="Net gains realised" value={gbp(realisedNet)} tone={realisedNet >= 0 ? "gain" : "loss"} />
        <Stat label="Tax-free headroom left" value={gbp(headroom)} tone="gain" big sub={realisedNet < 0 ? "AEA + realised losses" : "AEA − gains used"} />
        <Stat label="Holdings priced" value={`${rows.filter((r) => r.price !== "").length}/${rows.length}`} />
      </div>

      <div>
        <h3 className="text-sm font-semibold mb-2">Harvesting — sell within this year's allowance</h3>
        <div className="rounded-xl border border-[var(--border)] overflow-hidden">
          <table className="w-full text-sm">
            <thead className="bg-[var(--panel2)] text-[var(--muted)] text-xs uppercase tracking-wide">
              <tr>{["Ticker", "Avg cost", "Price now", "Gain / share", "Unrealised", "Max shares tax-free", "Gain realised"].map((h, i) => (
                <th key={i} className={"px-3 py-2 font-medium " + (i === 0 ? "text-left" : "text-right")}>{h}</th>
              ))}</tr>
            </thead>
            <tbody className="divide-y divide-[var(--border)] bg-[var(--panel)]">
              {rows.map((r) => (
                <tr key={r.tk} className="hover:bg-[var(--panel2)]">
                  <td className="px-3 py-2 font-medium">{r.tk}{r.recentBuy ? <AlertCircle size={13} className="inline ml-1 -mt-0.5 text-[var(--m-bb)]" /> : null}</td>
                  <td className="px-3 py-2 num text-right text-[var(--muted)]">{gbp(r.avg)}</td>
                  <td className="px-3 py-2 text-right">
                    <input type="number" value={r.price} placeholder="—"
                      onChange={(e) => setPrices((p) => ({ ...p, [r.tk]: e.target.value === "" ? undefined : +e.target.value }))}
                      className="input num w-24 text-right py-1" />
                  </td>
                  <td className={"px-3 py-2 num text-right " + (r.perShare == null ? "text-[var(--muted)]" : r.perShare >= 0 ? "text-[var(--gain)]" : "text-[var(--loss)]")}>{r.perShare != null ? gbp(r.perShare) : "—"}</td>
                  <td className={"px-3 py-2 num text-right " + (r.unreal == null ? "text-[var(--muted)]" : r.unreal >= 0 ? "text-[var(--gain)]" : "text-[var(--loss)]")}>{r.unreal != null ? gbp(r.unreal) : "—"}</td>
                  <td className="px-3 py-2 num text-right font-medium">{r.maxShares != null ? num(r.maxShares, 0) : (r.price === "" ? "—" : "no gain")}</td>
                  <td className="px-3 py-2 num text-right text-[var(--muted)]">{r.gainIf != null ? gbp(r.gainIf) : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p className="text-xs text-[var(--muted)] mt-2">
          "Max shares tax-free" assumes the whole remaining allowance is used on that one holding — the {gbp(headroom).replace(".00", "")} headroom is shared, so you can't stack it across several. Figures assume a clean sale with no repurchase within 30 days.
        </p>
      </div>

      {(Object.keys(recentBuys).length > 0 || pastBB.length > 0) && (
        <div className="rounded-xl border p-4 space-y-2"
          style={{ background: "color-mix(in srgb, var(--m-bb) 10%, transparent)", borderColor: "color-mix(in srgb, var(--m-bb) 35%, transparent)" }}>
          <h3 className="text-sm font-semibold flex items-center gap-2" style={{ color: "var(--m-bb)" }}><AlertCircle size={15} /> 30-day (bed &amp; breakfast) rule</h3>
          {Object.keys(recentBuys).length > 0 && (
            <div className="text-sm text-[var(--fg)]">
              You've bought within the last 30 days: {Object.entries(recentBuys).map(([t, q]) => `${num(q, q % 1 ? 2 : 0)} ${t}`).join(", ")}. A sale of the same holding now is matched to that purchase first — not your Section 104 pool — so it won't crystallise the pool gain you might be expecting.
            </div>
          )}
          {pastBB.length > 0 && (
            <div className="text-sm text-[var(--fg)]">
              Past disposals already matched under the 30-day rule: {pastBB.map((d) => `${d.ticker} ${d.date}`).join(", ")}.
            </div>
          )}
        </div>
      )}
      <p className="text-xs text-[var(--muted)]">
        The 30-day rule matches a disposal against any repurchase of the same security in the following 30 days before it touches the pool. To crystallise a pool gain (e.g. to use your allowance), avoid rebuying the same line within 30 days — buy a similar-but-not-identical fund, or repurchase inside an ISA/pension instead.
      </p>

      <MultiYearOptimiser pools={pools} prices={prices} income={income} />
    </div>
  );
}

/* Multi-year gain-harvesting: stagger disposals to soak up each year's AEA
   (and optionally basic-band room) and show how long an embedded gain takes to wash. */
function MultiYearOptimiser({ pools, prices, income }) {
  const yearNow = ukTaxYear(todayISO());
  const [startYear, setStartYear] = useState(yearNow);
  const [years, setYears] = useState(10);
  const [useBasicBand, setUseBasicBand] = useState(false);
  const [growth, setGrowth] = useState(0);

  const startOpts = useMemo(() => { const a = []; let y = yearNow; for (let i = 0; i < 4; i++) { a.push(y); y = nextTaxYear(y); } return a; }, [yearNow]);
  const holdings = useMemo(() => Object.keys(pools).filter((t) => pools[t].qty > 1e-6).map((t) => {
    const { qty, cost } = pools[t]; const p = prices[t];
    return { ticker: t, qty, cost, price: (p != null && p !== "" && !isNaN(+p)) ? +p : NaN };
  }).filter((h) => isFinite(h.price) && h.price > 0), [pools, prices]);

  const result = useMemo(() => {
    if (!holdings.length) return null;
    try { return optimiseDisposals({ holdings, startYear, years: Math.max(1, Math.min(40, +years || 1)), income: +income || 0, useBasicBand, growth: (+growth || 0) / 100 }); }
    catch { return null; }
  }, [holdings, startYear, years, income, useBasicBand, growth]);

  const priced = holdings.length;
  const totalTax = result ? result.schedule.reduce((s, r) => s + r.tax, 0) : 0;

  return (
    <div className="space-y-3 pt-2">
      <h3 className="text-sm font-semibold flex items-center gap-2"><Wand2 size={15} /> Multi-year disposal optimiser</h3>
      <p className="text-xs text-[var(--muted)]">
        Staggers sales across tax years to harvest gains up to each year's annual exempt amount (tax-free){useBasicBand ? ", plus basic-rate band room at 18%," : ""} then resets base cost by rebuying at market (bed-&amp;-ISA or bed-&amp;-spouse to sidestep the 30-day rule). Uses your {priced} priced GIA holding{priced === 1 ? "" : "s"}.
      </p>

      <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] p-3">
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 items-end">
          <Field label="Start tax year"><select value={startYear} onChange={(e) => setStartYear(e.target.value)} className="input w-full">{startOpts.map((y) => <option key={y}>{y}</option>)}</select></Field>
          <Field label="Horizon (years)"><input type="number" min="1" max="40" value={years} onChange={(e) => setYears(e.target.value)} className="input num w-full" /></Field>
          <Field label="Assumed growth %/yr"><input type="number" step="0.5" value={growth} onChange={(e) => setGrowth(e.target.value)} className="input num w-full" /></Field>
          <label className="flex items-center gap-2 text-sm cursor-pointer pb-2"><input type="checkbox" checked={useBasicBand} onChange={(e) => setUseBasicBand(e.target.checked)} className="accent-[var(--accent)]" /> Use basic-rate band (18%)</label>
        </div>
      </div>

      {!priced && <Empty msg="Set current prices on the holdings above (or the Holdings tab) to run the optimiser." />}

      {result && priced > 0 && (
        <>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <Stat label="Embedded gain now" value={gbp(result.startEmbedded)} tone={result.startEmbedded >= 0 ? "gain" : "loss"} />
            <Stat label={useBasicBand ? "Gain washed over horizon" : "Gain washed tax-free"} value={gbp(result.totalWashed)} big tone="gain" />
            <Stat label="Years to clear" value={result.yearsToClear ? `${result.yearsToClear}` : `>${years}`} sub={result.yearsToClear ? "" : "still embedded gain left"} />
            <Stat label="Tax over plan" value={gbp(totalTax)} tone={totalTax > 0 ? "loss" : undefined} sub={useBasicBand ? "basic-band 18%" : "within AEA"} />
          </div>

          <div className="rounded-xl border border-[var(--border)] overflow-hidden">
            <table className="w-full text-sm">
              <thead className="bg-[var(--panel2)] text-[var(--muted)] text-xs uppercase tracking-wide">
                <tr>{["Tax year", "Harvest", "AEA used", "Tax", "Sell", "Cumulative washed", "Gain still embedded"].map((h, i) => (
                  <th key={i} className={"px-3 py-2 font-medium " + (i === 0 || i === 4 ? "text-left" : "text-right")}>{h}</th>
                ))}</tr>
              </thead>
              <tbody className="divide-y divide-[var(--border)] bg-[var(--panel)]">
                {result.schedule.map((r) => (
                  <tr key={r.year} className="hover:bg-[var(--panel2)]">
                    <td className="px-3 py-2 num font-medium">{r.year}</td>
                    <td className="px-3 py-2 num text-right">{gbp(r.gainRealised)}</td>
                    <td className="px-3 py-2 num text-right text-[var(--muted)]">{gbp(r.aeaUsed).replace(".00", "")}</td>
                    <td className={"px-3 py-2 num text-right " + (r.tax > 0 ? "text-[var(--loss)]" : "text-[var(--muted)]")}>{r.tax > 0 ? gbp(r.tax) : "—"}</td>
                    <td className="px-3 py-2 text-[var(--muted)] text-xs">{r.sells.map((s) => `${num(s.shares, s.shares % 1 ? 2 : 0)} ${s.ticker}`).join(", ") || "—"}</td>
                    <td className="px-3 py-2 num text-right text-[var(--muted)]">{gbp(r.cumulativeWashed)}</td>
                    <td className="px-3 py-2 num text-right">{gbp(r.remainingUnrealised)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="text-xs text-[var(--muted)]">
            Each year's harvest sells the highest gain-per-share holdings first and assumes you rebuy at the same price to reset base cost. Growth compounds the price of unsold shares. This models the CGT wash only — dealing costs, spreads and stamp duty on rebuys aren't included, and a bed-&amp;-ISA rebuy also consumes ISA allowance.
          </p>
        </>
      )}
    </div>
  );
}

/* ---------------------------- Report tab ---------------------------- */
function ReportTab({ taxYears, disposals, income, carried }) {
  const [ry, setRy] = useState(taxYears[0] || "2025/26");
  const [msg, setMsg] = useState("");
  const yr = taxYears.includes(ry) ? ry : (taxYears[0] || "2025/26");
  const yd = disposals.filter((d) => d.taxYear === yr);
  const liab = liabilityForYear(yd, { income, carriedLosses: carried });
  const totalCost = yd.reduce((s, d) => s + d.cost, 0);
  const flash = (m) => { setMsg(m); setTimeout(() => setMsg(""), 3500); };

  const csvCell = (v) => {
    const s = String(v);
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const exportCSV = async () => {
    const rows = [["Tax year", "Disposal date", "Security", "Matching method", "Quantity", "Proceeds GBP", "Allowable cost GBP", "Gain/loss GBP"]];
    for (const d of yd) for (const l of d.legs) rows.push([yr, d.date, d.ticker, METHOD[l.method].label, l.quantity, l.proceeds.toFixed(2), l.cost.toFixed(2), l.gain.toFixed(2)]);
    rows.push([], ["Summary (SA108 Capital Gains — listed shares & securities)"]);
    rows.push(["Number of disposals", yd.length]);
    rows.push(["Disposal proceeds (box 24)", liab.proceeds.toFixed(2)]);
    rows.push(["Allowable costs (box 25)", totalCost.toFixed(2)]);
    rows.push(["Gains before losses (box 26)", liab.gains.toFixed(2)]);
    rows.push(["Losses in the year (box 27)", liab.losses.toFixed(2)]);
    rows.push(["Annual exempt amount", liab.aea.toFixed(2)]);
    rows.push(["Taxable gain", liab.taxable.toFixed(2)]);
    liab.breakdown.forEach((b) => rows.push([`Taxed at ${fmtRate(b.rate)}`, b.amount.toFixed(2), `tax ${b.tax.toFixed(2)}`]));
    rows.push(["CGT due", liab.tax.toFixed(2)]);
    rows.push(["Reporting required", liab.reporting ? "Yes" : "No"]);
    const text = rows.map((r) => r.map(csvCell).join(",")).join("\n");
    let dl = false;
    try {
      const url = URL.createObjectURL(new Blob([text], { type: "text/csv" }));
      const a = document.createElement("a"); a.href = url; a.download = `cgt-report-${yr.replace("/", "-")}.csv`;
      document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000); dl = true;
    } catch { /* sandbox */ }
    try { await navigator.clipboard.writeText(text); flash(dl ? "CSV downloaded (also copied)." : "Download blocked here — CSV copied to clipboard."); }
    catch { flash(dl ? "CSV downloaded." : "Couldn't export in this frame — use the deployed app."); }
  };

  if (!taxYears.length) return <Empty msg="No disposals to report. Add or import transactions first." />;
  return (
    <div className="space-y-4">
      <div className="flex items-end gap-3 flex-wrap no-print">
        <Field label="Tax year">
          <select value={yr} onChange={(e) => setRy(e.target.value)} className="input num">
            {taxYears.map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
        </Field>
        <button onClick={() => window.print()} className="btn-accent"><Printer size={15} /> Print / Save as PDF</button>
        <button onClick={exportCSV} className="inline-flex items-center gap-1.5 text-sm font-medium px-3 py-2 rounded-lg border border-[var(--border)] bg-[var(--panel)] hover:bg-[var(--panel2)]"><Download size={15} /> Download CSV</button>
        {msg && <span className="text-xs text-[var(--muted)]">{msg}</span>}
      </div>

      {/* printable report */}
      <div className="print-area rounded-xl border border-[var(--border)] bg-[var(--panel)] p-6 space-y-5">
        <div className="flex items-baseline justify-between border-b border-[var(--border)] pb-3">
          <div>
            <h2 className="text-lg font-semibold">Capital Gains Tax computation</h2>
            <p className="text-sm text-[var(--muted)]">Listed shares &amp; securities · Tax year {yr}</p>
          </div>
          <span className="text-xs text-[var(--muted)]">Generated {todayISO()}</span>
        </div>

        <div>
          <h3 className="text-sm font-semibold mb-2">Summary (SA108)</h3>
          <table className="w-full text-sm">
            <tbody className="num">
              {[
                ["Number of disposals", num(yd.length, 0)],
                ["Disposal proceeds — box 24", gbp(liab.proceeds)],
                ["Allowable costs — box 25", gbp(totalCost)],
                ["Gains in the year before losses — box 26", gbp(liab.gains)],
                ["Losses in the year — box 27", gbp(liab.losses)],
                ["Annual exempt amount", gbp(liab.aea)],
                ...(liab.usedCarried ? [["Losses brought forward used", gbp(liab.usedCarried)]] : []),
                ["Net taxable gain", gbp(liab.taxable)],
                ...liab.breakdown.map((b) => [`  taxed at ${fmtRate(b.rate)}`, `${gbp(b.amount)}  →  ${gbp(b.tax)}`]),
              ].map(([k, v], i) => (
                <tr key={i} className="border-b border-[var(--border)]">
                  <td className="py-1.5 font-sans text-[var(--muted)]">{k}</td>
                  <td className="py-1.5 text-right">{v}</td>
                </tr>
              ))}
              <tr className="border-t-2 border-[var(--border)]">
                <td className="py-2 font-sans font-semibold">CGT due</td>
                <td className="py-2 text-right font-semibold text-[var(--loss)]">{gbp(liab.tax)}</td>
              </tr>
            </tbody>
          </table>
          <p className="text-xs text-[var(--muted)] mt-2">Reporting {liab.reporting ? "required" : "not required"} for this year (tax due, or proceeds over the reporting threshold).</p>
        </div>

        <div>
          <h3 className="text-sm font-semibold mb-2">Disposal schedule &amp; matching</h3>
          <table className="w-full text-xs">
            <thead className="text-[var(--muted)] border-b border-[var(--border)]">
              <tr>{["Date", "Security", "Method", "Qty", "Proceeds", "Cost", "Gain/loss"].map((h, i) => (
                <th key={i} className={"py-1.5 font-medium " + (i < 3 ? "text-left" : "text-right")}>{h}</th>
              ))}</tr>
            </thead>
            <tbody className="num">
              {yd.map((d) => d.legs.map((l, li) => (
                <tr key={d.id + li} className="border-b border-[var(--border)]">
                  <td className="py-1.5">{li === 0 ? d.date : ""}</td>
                  <td className="py-1.5 font-sans">{li === 0 ? d.ticker : ""}</td>
                  <td className="py-1.5 font-sans">{METHOD[l.method].label}</td>
                  <td className="py-1.5 text-right">{num(l.quantity, l.quantity % 1 ? 4 : 0)}</td>
                  <td className="py-1.5 text-right">{gbp(l.proceeds)}</td>
                  <td className="py-1.5 text-right">{gbp(l.cost)}</td>
                  <td className={"py-1.5 text-right " + (l.gain >= 0 ? "text-[var(--gain)]" : "text-[var(--loss)]")}>{gbp(l.gain)}</td>
                </tr>
              )))}
            </tbody>
          </table>
        </div>
        <p className="text-[10px] text-[var(--muted)] pt-2 border-t border-[var(--border)]">
          Prepared as a computation to support a Self Assessment return. HMRC share-identification rules applied: same-day, then 30-day, then Section 104 pool. Not tax advice — verify before filing.
        </p>
      </div>
    </div>
  );
}

/* --------------------------- What-if tab ---------------------------- */
function WhatIfTab({ pools, disposals, income, carried, prices = {} }) {
  const tickers = Object.keys(pools).filter((t) => pools[t].qty > 1e-6);
  const [ticker, setTicker] = useState(tickers[0] || "");
  const tk = ticker && pools[ticker] ? ticker : tickers[0] || "";
  const pool = pools[tk] || { qty: 0, cost: 0 };
  const avg = pool.qty ? pool.cost / pool.qty : 0;

  const [priceEdited, setPriceEdited] = useState(false);
  const [priceRaw, setPriceRaw] = useState("");
  // default the price from the Holdings tab unless the user has typed their own
  const price = priceEdited ? priceRaw : (prices[tk] != null ? String(prices[tk]) : "");
  const setPrice = (v) => { setPriceEdited(true); setPriceRaw(v); };
  const [sellQty, setSellQty] = useState("");
  const yearNow = ukTaxYear(todayISO());
  const realisedThisYear = disposals.filter((d) => d.taxYear === yearNow);
  const base = liabilityForYear(realisedThisYear, { income, carriedLosses: carried });

  const p = +price || 0, q = Math.min(+sellQty || 0, pool.qty);
  const hypo = q > 0 && p > 0 ? { date: todayISO(), ticker: tk, quantity: q, proceeds: q * p, gain: q * p - avg * q, taxYear: yearNow, legs: [], cost: avg * q } : null;
  const withHypo = hypo ? liabilityForYear([...realisedThisYear, hypo], { income, carriedLosses: carried }) : base;
  const marginalTax = withHypo.tax - base.tax;

  const aeaHeadroom = Math.max(0, aeaForYear(yearNow) - base.net);
  const maxSharesAea = p > 0 ? sharesForTargetGain(pool.qty, pool.cost, p, aeaHeadroom) : 0;

  if (!tickers.length) return <Empty msg="No open holdings to model. Add buy transactions first." />;
  return (
    <div className="space-y-5">
      <div className="flex items-end gap-3 flex-wrap">
        <Field label="Holding">
          <select value={tk} onChange={(e) => { setTicker(e.target.value); setPriceEdited(false); setPriceRaw(""); }} className="input">{tickers.map((t) => <option key={t}>{t}</option>)}</select>
        </Field>
        <Field label="Price now (GBP/share)"><input type="number" value={price} onChange={(e) => setPrice(e.target.value)} placeholder="e.g. 60.00" className="input num w-36" /></Field>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
        <Stat label="Pool quantity" value={num(pool.qty, pool.qty % 1 ? 4 : 0)} />
        <Stat label="Pool cost" value={gbp(pool.cost)} />
        <Stat label="Average cost" value={gbp(avg)} />
        <Stat label={`Realised ${yearNow}`} value={gbp(base.net)} tone={base.net >= 0 ? "gain" : "loss"} />
      </div>

      <div className="grid sm:grid-cols-2 gap-4">
        {/* scenario A */}
        <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] p-4 space-y-3">
          <h3 className="font-semibold text-sm flex items-center gap-2"><FlaskConical size={15} className="text-[var(--accent)]" /> Sell a quantity</h3>
          <Field label="Shares to sell"><input type="number" value={sellQty} onChange={(e) => setSellQty(e.target.value)} className="input num w-full" /></Field>
          {hypo ? (
            <div className="text-sm space-y-1 num">
              <Row k="Proceeds" v={gbp(hypo.proceeds)} />
              <Row k="Cost (pool avg)" v={gbp(hypo.cost)} />
              <Row k="Gain on sale" v={gbp(hypo.gain)} tone={hypo.gain >= 0 ? "gain" : "loss"} />
              <div className="h-px bg-[var(--border)] my-1" />
              <Row k="CGT before" v={gbp(base.tax)} />
              <Row k="CGT after" v={gbp(withHypo.tax)} />
              <Row k="Marginal CGT" v={gbp(marginalTax)} tone="loss" bold />
            </div>
          ) : <p className="text-sm text-[var(--muted)]">Enter a price and quantity to model the disposal.</p>}
        </div>

        {/* scenario B */}
        <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] p-4 space-y-3">
          <h3 className="font-semibold text-sm flex items-center gap-2"><Check size={15} className="text-[var(--gain)]" /> Stay within the {gbp(aeaForYear(yearNow)).replace(".00", "")} allowance</h3>
          {p > 0 ? (
            <div className="text-sm space-y-1 num">
              <Row k="AEA headroom left" v={gbp(aeaHeadroom)} />
              <Row k="Gain per share" v={gbp(p - avg)} />
              <Row k="Max shares, tax-free" v={num(maxSharesAea, 0) + " sh"} tone="gain" bold />
              <p className="text-xs text-[var(--muted)] pt-1 font-sans">Clean sale, no repurchase within 30 days. Selling more triggers CGT on the excess at your marginal rate.</p>
            </div>
          ) : <p className="text-sm text-[var(--muted)]">Enter a current price to see how many shares fit inside this year's allowance.</p>}
        </div>
      </div>
    </div>
  );
}

/* --------------------------- Import tab ----------------------------- */
const FIELDS = ["date", "ticker", "side", "quantity", "nativeCurrency", "nativeAmount", "fxRate", "gbpAmount"];
function ImportTab({ setTxns, setTab, setIncomeEntries, setEriEntries, secMeta }) {
  const [mode, setMode] = useState("ibkr");
  const [wrapper, setWrapper] = useState("GIA");
  const [raw, setRaw] = useState("");
  const [parsed, setParsed] = useState(null);
  const [map, setMap] = useState({});
  const [ib, setIb] = useState(null);       // parseIBKR result
  const [importing, setImporting] = useState(false);
  const [note, setNote] = useState("");
  const [wb, setWb] = useState(null);        // parsed iShares workbook: { fileName, sheets: [{name, headerRowIdx, colMap, headerCells, rows}] }
  const [activeSheet, setActiveSheet] = useState(0);
  const [onlyHeld, setOnlyHeld] = useState(true);
  const [checked, setChecked] = useState({}); // isin -> bool
  const [wbBusy, setWbBusy] = useState(false);

  const readFile = (e, cb) => { const f = e.target.files?.[0]; if (!f) return; const r = new FileReader(); r.onload = () => cb(String(r.result)); r.readAsText(f); e.target.value = ""; };

  // ---- IBKR ----
  const parseIb = (text) => { const t = (text ?? raw).trim(); if (!t) return; setIb(parseIBKR(t, { defaultWrapper: wrapper })); };
  React.useEffect(() => { if (ib) setIb((r) => ({ ...r, trades: r.trades.map((t) => ({ ...t, wrapper })), income: r.income.map((t) => ({ ...t, wrapper })) })); }, [wrapper]); // eslint-disable-line
  const doImportIb = async () => {
    if (!ib) return; setImporting(true); setNote("");
    const cache = {};
    const resolve = async (row, gbpKey) => {
      if (!row.needsFx) return;
      const k = row.nativeCurrency + row.date;
      if (!(k in cache)) cache[k] = await fxHistorical(row.nativeCurrency, row.date);
      const fx = cache[k];
      if (fx) { row.fxRate = fx; row[gbpKey] = Math.round(row.nativeAmount * fx * 100) / 100; }
    };
    const trades = ib.trades.map((t) => ({ ...t })), income = ib.income.map((t) => ({ ...t }));
    for (const t of trades) await resolve(t, "gbpAmount");
    for (const t of income) await resolve(t, "amount");
    const newTxns = trades.filter((t) => t.gbpAmount != null).map((t) => ({ id: uid(), date: t.date, ticker: t.ticker, isin: t.isin, side: t.side, quantity: t.quantity, nativeCurrency: t.nativeCurrency, nativeAmount: t.nativeAmount, fxRate: t.fxRate || 1, gbpAmount: t.gbpAmount, wrapper: t.wrapper, note: "IBKR import" }));
    const newIncome = income.filter((t) => t.amount != null).map((t) => ({ id: uid(), date: t.date, ticker: t.ticker, kind: t.kind, amount: t.amount, wrapper: t.wrapper, note: "IBKR import" }));
    const skipped = (trades.length - newTxns.length) + (income.length - newIncome.length);
    setTxns((p) => [...p, ...newTxns]);
    if (newIncome.length) setIncomeEntries((p) => [...p, ...newIncome]);
    setImporting(false);
    if (skipped) { setNote(`Imported ${newTxns.length} trades and ${newIncome.length} income rows. ${skipped} row(s) skipped — FX could not be resolved; add them manually.`); }
    else setTab(newTxns.length ? "ledger" : "income");
  };

  // ---- generic ----
  const parse = () => {
    const res = Papa.parse(raw.trim(), { header: true, skipEmptyLines: true });
    if (!res.data?.length) return;
    const cols = res.meta.fields || [];
    const find = (re) => cols.find((c) => re.test(c));
    const guess = {};
    guess.date = find(/date|trade date|settl/i); guess.ticker = find(/ticker|symbol|instrument|stock/i);
    guess.side = find(/side|action|type|buy.?sell|b\/s/i); guess.quantity = find(/qty|quantity|shares|units/i);
    guess.nativeCurrency = find(/currency|ccy/i); guess.nativeAmount = find(/amount|proceeds|cost|value|consideration|net/i);
    guess.fxRate = find(/fx|rate|exchange/i); guess.gbpAmount = find(/gbp|sterling/i);
    setParsed(res.data); setMap(guess);
  };
  const normSide = (v) => /sell|^s$|sld|disp/i.test(v || "") ? "SELL" : "BUY";
  const preview = useMemo(() => (!parsed ? [] : parsed.slice(0, 5).map((r) => mapRow(r, map, normSide, wrapper))), [parsed, map, wrapper]);
  const doImport = () => {
    const rows = parsed.map((r) => mapRow(r, map, normSide, wrapper)).filter((t) => t.date && t.ticker && +t.quantity > 0);
    setTxns((p) => [...p, ...rows]); setTab("ledger");
  };

  // ---- iShares / issuer ERI workbook ----
  const heldIsins = useMemo(() => new Set(Object.values(secMeta || {}).map((s) => (s.isin || "").toUpperCase()).filter(Boolean)), [secMeta]);
  const isinToTicker = useMemo(() => {
    const m = {}; for (const [tk, s] of Object.entries(secMeta || {})) if (s.isin) m[s.isin.toUpperCase()] = tk; return m;
  }, [secMeta]);

  const readWorkbookFile = (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    setWbBusy(true); setWb(null); setChecked({});
    const r = new FileReader();
    r.onload = () => {
      try {
        const data = new Uint8Array(r.result);
        const book = XLSX.read(data, { type: "array", cellDates: true });
        const sheets = book.SheetNames.map((name) => ({ name, aoa: XLSX.utils.sheet_to_json(book.Sheets[name], { header: 1, raw: true, defval: "" }) }));
        const result = parseISharesWorkbook(sheets, null); // parse unfiltered; "only held" is a display filter below
        const bestSheet = result.reduce((best, s, i) => {
          const scoreOf = (idx) => result[idx].rows.filter((row) => heldIsins.has(row.isin)).length || result[idx].rows.length;
          return scoreOf(i) > scoreOf(best) ? i : best;
        }, 0);
        setWb({ fileName: f.name, sheets: result });
        setActiveSheet(bestSheet);
      } catch (err) {
        setWb({ fileName: f.name, sheets: [], error: err.message || "Could not read this file as a spreadsheet." });
      }
      setWbBusy(false);
    };
    r.readAsArrayBuffer(f);
    e.target.value = "";
  };

  const sheet = wb?.sheets?.[activeSheet];
  const allRows = sheet?.rows || [];
  const rows = useMemo(() => onlyHeld ? allRows.filter((r) => heldIsins.has(r.isin)) : allRows, [allRows, onlyHeld, heldIsins]);
  React.useEffect(() => {
    const c = {}; rows.forEach((r) => { c[r.isin] = true; }); setChecked(c);
  }, [sheet, onlyHeld]); // eslint-disable-line
  const toggleAll = (v) => { const c = {}; rows.forEach((r) => { c[r.isin] = v; }); setChecked(c); };
  const selectedCount = rows.filter((r) => checked[r.isin]).length;

  const doImportEri = async () => {
    const selected = rows.filter((r) => checked[r.isin]);
    const fxCache = {};
    const toAdd = [];
    for (const r of selected) {
      const ticker = isinToTicker[r.isin] || r.isin;
      let fxRate = r.currency === "GBP" || r.currency === "GBp" ? 1 : 0;
      if (fxRate === 0 && r.distributionDate) {
        const k = r.currency + r.distributionDate;
        if (!(k in fxCache)) fxCache[k] = await fxHistorical(r.currency, r.distributionDate);
        fxRate = fxCache[k] || 0;
      }
      const e = { id: uid(), ticker, periodEnd: r.periodEnd, distributionDate: r.distributionDate, perShare: +r.perShare || 0, currency: r.currency || "GBP", fxRate, treatment: r.treatment || "dividend" };
      if (e.ticker && e.periodEnd && e.distributionDate && e.perShare) toAdd.push(e);
    }
    if (!toAdd.length) return;
    setEriEntries((p) => [...p, ...toAdd]);
    const unresolvedFx = toAdd.filter((e) => e.currency !== "GBP" && e.currency !== "GBp" && !e.fxRate).length;
    if (unresolvedFx) setNote(`Imported ${toAdd.length} ERI entries. ${unresolvedFx} needed an FX rate that couldn't be fetched — set it manually on the Income tab.`);
    else setTab("income");
  };

  const Tab = ({ k, label }) => (
    <button onClick={() => setMode(k)} className={"px-3 py-1.5 text-sm rounded-lg border " + (mode === k ? "bg-[var(--accent)] text-[var(--accent-fg)] border-transparent" : "border-[var(--border)] text-[var(--muted)] hover:text-[var(--fg)]")}>{label}</button>
  );

  return (
    <div className="space-y-4">
      <div className="flex items-center gap-2 flex-wrap">
        <Tab k="ibkr" label="Interactive Brokers" /><Tab k="generic" label="Generic CSV" /><Tab k="ishares" label="iShares ERI" />
        <span className="ml-auto" />
        {mode !== "ishares" && <Field label="Import into wrapper"><select value={wrapper} onChange={(e) => setWrapper(e.target.value)} className="input">{WRAPPERS.map((w) => <option key={w}>{w}</option>)}</select></Field>}
      </div>

      {mode === "ibkr" && (
        <>
          <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] p-4 space-y-3">
            <p className="text-sm text-[var(--muted)]">Paste (or upload) an IBKR <strong>Flex Query</strong> CSV or an <strong>Activity Statement</strong> CSV. Trades and dividends/interest are both picked up. A Flex query carries an FX-to-base rate, so GBP conversion is automatic; Activity exports lack it, so non-GBP rows are converted by trade-date FX on import. {wrapper !== "GIA" && <span className="text-[var(--fg)]">Note: {wrapper} is tax-sheltered, so these rows won't affect CGT or income tax.</span>}</p>
            <textarea value={raw} onChange={(e) => setRaw(e.target.value)} rows={5} placeholder={'Symbol,ISIN,TradeDate,Buy/Sell,Quantity,TradePrice,Proceeds,IBCommission,CurrencyPrimary,FXRateToBase,AssetClass\nAAPL,US0378331005,20240115,BUY,10,180,-1800,-1,USD,0.79,STK'} className="input num w-full font-mono text-xs" />
            <div className="flex items-center gap-2">
              <button onClick={() => parseIb()} className="btn-accent"><Wand2 size={15} /> Parse</button>
              <label className="text-sm text-[var(--accent)] cursor-pointer flex items-center gap-1"><Upload size={14} /> Upload CSV<input type="file" accept=".csv,text/csv" className="hidden" onChange={(e) => readFile(e, (txt) => { setRaw(txt); parseIb(txt); })} /></label>
            </div>
          </div>

          {ib && (
            <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] p-4 space-y-3">
              <div className="flex items-center gap-4 text-sm flex-wrap">
                <span className="font-semibold">{ib.format === "activity" ? "Activity Statement" : "Flex Query"} detected</span>
                <span className="num">{ib.trades.length} trades</span>
                <span className="num">{ib.income.filter((i) => i.kind === "dividend").length} dividends</span>
                <span className="num">{ib.income.filter((i) => i.kind === "interest").length} interest</span>
              </div>
              {ib.warnings.map((w, i) => (
                <div key={i} className="flex items-start gap-2 text-xs rounded-lg px-3 py-2 text-[var(--loss)]" style={{ background: "color-mix(in srgb, var(--loss) 10%, transparent)" }}><AlertTriangle size={14} className="mt-0.5 shrink-0" />{w}</div>
              ))}
              {ib.trades.length > 0 && (
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead className="text-[var(--muted)]"><tr>{["date", "ticker", "side", "qty", "ccy", "native", "GBP"].map((h) => <th key={h} className="px-2 py-1 text-left">{h}</th>)}</tr></thead>
                    <tbody className="num">
                      {ib.trades.slice(0, 6).map((t, i) => (
                        <tr key={i} className="border-t border-[var(--border)]">
                          <td className="px-2 py-1">{t.date}</td><td className="px-2 py-1">{t.ticker}</td><td className="px-2 py-1">{t.side}</td>
                          <td className="px-2 py-1">{num(t.quantity, t.quantity % 1 ? 4 : 0)}</td><td className="px-2 py-1">{t.nativeCurrency}</td>
                          <td className="px-2 py-1">{num(t.nativeAmount)}</td><td className="px-2 py-1">{t.gbpAmount == null ? "FX on import" : gbp(t.gbpAmount)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {ib.trades.length > 6 && <p className="text-xs text-[var(--muted)] mt-1">+{ib.trades.length - 6} more…</p>}
                </div>
              )}
              {note && <div className="text-xs text-[var(--muted)]">{note}</div>}
              <div className="flex items-center justify-between">
                <span className="text-xs text-[var(--muted)]">Imports into <strong>{wrapper}</strong>. Trades → ledger, dividends/interest → Income tab.</span>
                <button onClick={doImportIb} disabled={importing} className="btn-accent">{importing ? <RefreshCw size={15} className="animate-spin" /> : <FileUp size={15} />} Import</button>
              </div>
            </div>
          )}
        </>
      )}

      {mode === "generic" && (
        <>
          <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] p-4 space-y-3">
            <p className="text-sm text-[var(--muted)]">Paste a CSV from any broker. Columns are auto-mapped — adjust below if needed. Rows import into <strong>{wrapper}</strong>.</p>
            <textarea value={raw} onChange={(e) => setRaw(e.target.value)} rows={5} placeholder="Date,Symbol,Action,Quantity,Currency,Amount,FXRate&#10;2025-06-02,WFC,SELL,200,USD,18718,0.78" className="input num w-full font-mono text-xs" />
            <button onClick={parse} className="btn-accent"><Wand2 size={15} /> Parse & map</button>
          </div>
          {parsed && (
            <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] p-4 space-y-3">
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                {FIELDS.map((f) => (
                  <Field key={f} label={f}>
                    <select value={map[f] || ""} onChange={(e) => setMap((m) => ({ ...m, [f]: e.target.value }))} className="input w-full text-xs">
                      <option value="">—</option>
                      {(Object.keys(parsed[0] || {})).map((c) => <option key={c}>{c}</option>)}
                    </select>
                  </Field>
                ))}
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs">
                  <thead className="text-[var(--muted)]"><tr>{["date", "ticker", "side", "qty", "ccy", "native", "fx", "gbp"].map((h) => <th key={h} className="px-2 py-1 text-left">{h}</th>)}</tr></thead>
                  <tbody className="num">
                    {preview.map((t, i) => (
                      <tr key={i} className="border-t border-[var(--border)]">
                        <td className="px-2 py-1">{t.date}</td><td className="px-2 py-1">{t.ticker}</td><td className="px-2 py-1">{t.side}</td>
                        <td className="px-2 py-1">{t.quantity}</td><td className="px-2 py-1">{t.nativeCurrency}</td><td className="px-2 py-1">{num(t.nativeAmount)}</td>
                        <td className="px-2 py-1">{num(t.fxRate, 4)}</td><td className="px-2 py-1">{gbp(t.gbpAmount)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <div className="flex items-center justify-between">
                <span className="text-xs text-[var(--muted)]">{parsed.length} rows ready. GBP fills from native × FX when GBP column is unmapped.</span>
                <button onClick={doImport} className="btn-accent"><FileUp size={15} /> Import {parsed.length} rows</button>
              </div>
            </div>
          )}
        </>
      )}

      {mode === "ishares" && (
        <>
          <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] p-4 space-y-3">
            <p className="text-sm text-[var(--muted)]">
              Upload an iShares/BlackRock <strong>"UK Reportable Income"</strong> workbook — one file per fund umbrella (iShares Plc, iShares III plc, iShares VII plc, etc.), downloaded from <span className="font-mono text-xs">ishares.com</span> → Literature → Tax Information. Rows are matched to your holdings by ISIN and added as excess reportable income entries — GIA only, since ISA/SIPP are exempt.
            </p>
            <label className="inline-flex items-center gap-2 text-sm text-[var(--accent)] cursor-pointer">
              <Upload size={14} /> Upload workbook (.xlsx/.xls)
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={readWorkbookFile} />
            </label>
            {wbBusy && <div className="flex items-center gap-2 text-xs text-[var(--muted)]"><RefreshCw size={13} className="animate-spin" /> Reading workbook…</div>}
            {wb?.error && (
              <div className="flex items-start gap-2 text-xs rounded-lg px-3 py-2 text-[var(--loss)]" style={{ background: "color-mix(in srgb, var(--loss) 10%, transparent)" }}><AlertTriangle size={14} className="mt-0.5 shrink-0" />{wb.error}</div>
            )}
          </div>

          {wb && !wb.error && (
            <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] p-4 space-y-3">
              <div className="flex items-center gap-3 flex-wrap text-sm">
                <span className="font-semibold truncate max-w-[16rem]" title={wb.fileName}>{wb.fileName}</span>
                {wb.sheets.length > 1 && (
                  <select value={activeSheet} onChange={(e) => setActiveSheet(+e.target.value)} className="input text-xs w-auto">
                    {wb.sheets.map((s, i) => <option key={i} value={i}>{s.name} ({s.rows.length} held match{s.rows.length === 1 ? "" : "es"})</option>)}
                  </select>
                )}
                <label className="flex items-center gap-1.5 text-xs cursor-pointer ml-auto">
                  <input type="checkbox" checked={onlyHeld} onChange={(e) => setOnlyHeld(e.target.checked)} className="accent-[var(--accent)]" /> Only show my holdings
                </label>
              </div>

              {sheet && sheet.headerRowIdx < 0 && (
                <div className="flex items-start gap-2 text-xs rounded-lg px-3 py-2 text-[var(--loss)]" style={{ background: "color-mix(in srgb, var(--loss) 10%, transparent)" }}>
                  <AlertTriangle size={14} className="mt-0.5 shrink-0" />Couldn't find a header row with an ISIN column in this sheet — it may not be a reportable-income report, or uses a layout this importer doesn't recognise yet.
                </div>
              )}

              {sheet && sheet.headerRowIdx >= 0 && rows.length === 0 && (
                <Empty msg={onlyHeld ? "No rows in this sheet match your current holdings' ISINs. Try unchecking \"Only show my holdings\", or check you've got the right umbrella file." : "No ERI rows found in this sheet (all excess income was zero, or no data rows present)."} />
              )}

              {rows.length > 0 && (
                <>
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead className="text-[var(--muted)]">
                        <tr>
                          <th className="px-2 py-1 text-left"><input type="checkbox" checked={selectedCount === rows.length} onChange={(e) => toggleAll(e.target.checked)} className="accent-[var(--accent)]" /></th>
                          {["Fund", "Ticker", "ISIN", "Period end", "Distribution date", "ERI/unit", "Ccy", "Taxed as"].map((h) => <th key={h} className="px-2 py-1 text-left">{h}</th>)}
                        </tr>
                      </thead>
                      <tbody className="num">
                        {rows.map((r) => {
                          const ticker = isinToTicker[r.isin];
                          return (
                            <tr key={r.isin + r.periodEnd} className="border-t border-[var(--border)]">
                              <td className="px-2 py-1"><input type="checkbox" checked={!!checked[r.isin]} onChange={(e) => setChecked((c) => ({ ...c, [r.isin]: e.target.checked }))} className="accent-[var(--accent)]" /></td>
                              <td className="px-2 py-1 max-w-[14rem] truncate" title={r.fundName}>{r.fundName}</td>
                              <td className="px-2 py-1 font-medium">{ticker || <span className="text-[var(--loss)]" title="No ticker in your holdings has this ISIN — add it on the Holdings tab first">unmatched</span>}</td>
                              <td className="px-2 py-1 font-mono text-[var(--muted)]">{r.isin}</td>
                              <td className="px-2 py-1">{r.periodEnd}</td>
                              <td className="px-2 py-1">{r.distributionDate}</td>
                              <td className="px-2 py-1">{r.perShare}</td>
                              <td className="px-2 py-1">{r.currency}</td>
                              <td className="px-2 py-1 capitalize">{r.treatment}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  <p className="text-xs text-[var(--muted)]">
                    "Taxed as" comes straight from the report's own "Meets definition of a Bond Fund" flag — bond funds are taxed as interest, everything else as dividend. Rows marked <span className="text-[var(--loss)]">unmatched</span> don't have a ticker with that ISIN on your Holdings tab; add the ISIN there first if you want to import them. Non-GBP amounts have their FX rate fetched automatically for the distribution date on import.
                  </p>
                  {note && <div className="text-xs text-[var(--muted)]">{note}</div>}
                  <div className="flex items-center justify-between">
                    <span className="text-xs text-[var(--muted)]">{selectedCount}/{rows.length} selected.</span>
                    <button onClick={doImportEri} disabled={!selectedCount} className="btn-accent disabled:opacity-50"><FileUp size={15} /> Import {selectedCount} ERI entr{selectedCount === 1 ? "y" : "ies"}</button>
                  </div>
                </>
              )}
            </div>
          )}
        </>
      )}
    </div>
  );
}
function mapRow(r, map, normSide, wrapper) {
  const g = (f) => (map[f] ? r[map[f]] : "");
  const ccy = (g("nativeCurrency") || "GBP").toUpperCase().trim();
  const native = parseFloat(String(g("nativeAmount")).replace(/[^0-9.\-]/g, "")) || 0;
  let fx = parseFloat(g("fxRate")) || (ccy === "GBP" ? 1 : 0);
  let gbpA = parseFloat(String(g("gbpAmount")).replace(/[^0-9.\-]/g, "")) || 0;
  if (!gbpA && native && fx) gbpA = +(native * fx).toFixed(2);
  if (!fx && gbpA && native) fx = +(gbpA / native).toFixed(6);
  return {
    id: uid(), date: (g("date") || "").slice(0, 10), ticker: (g("ticker") || "").toUpperCase().trim(),
    side: normSide(g("side")), quantity: Math.abs(parseFloat(g("quantity")) || 0),
    nativeCurrency: ccy, nativeAmount: native, fxRate: fx || 1, gbpAmount: gbpA, wrapper: wrapper || "GIA", note: "imported",
  };
}

/* ----------------------------- atoms -------------------------------- */
function IconBtn({ children, as = "button", ...p }) {
  const C = as; return <C {...p} className="inline-flex items-center justify-center w-9 h-9 rounded-lg border border-[var(--border)] bg-[var(--panel)] hover:bg-[var(--panel2)] text-[var(--fg)] cursor-pointer">{children}</C>;
}
function Field({ label, children }) {
  return <label className="flex flex-col gap-1"><span className="text-xs text-[var(--muted)]">{label}</span>{children}</label>;
}
function Stat({ label, value, sub, tone, big }) {
  const c = tone === "gain" ? "text-[var(--gain)]" : tone === "loss" ? "text-[var(--loss)]" : "text-[var(--fg)]";
  return (
    <div className="rounded-xl border border-[var(--border)] bg-[var(--panel)] px-4 py-3">
      <div className="text-xs text-[var(--muted)]">{label}</div>
      <div className={`num font-semibold ${big ? "text-2xl" : "text-lg"} ${c} mt-0.5`}>{value}</div>
      {sub && <div className="text-xs text-[var(--muted)] mt-0.5">{sub}</div>}
    </div>
  );
}
function Row({ k, v, tone, bold }) {
  const c = tone === "gain" ? "text-[var(--gain)]" : tone === "loss" ? "text-[var(--loss)]" : "";
  return <div className="flex justify-between"><span className="text-[var(--muted)] font-sans">{k}</span><span className={`${c} ${bold ? "font-semibold" : ""}`}>{v}</span></div>;
}
function MethodChip({ m }) {
  const d = METHOD[m];
  return <span className="text-xs font-medium px-2 py-0.5 rounded-full" style={{ color: `var(${d.v})`, background: "var(--chip)" }}>{d.label}</span>;
}
function Empty({ msg }) {
  return <div className="rounded-xl border border-dashed border-[var(--border)] py-12 text-center text-sm text-[var(--muted)]">{msg}</div>;
}

/* inline utility classes used above */
const _style = document.createElement("style");
_style.textContent = `
  .input{background:var(--panel2);border:1px solid var(--border);border-radius:.5rem;padding:.4rem .6rem;font-size:.875rem;color:var(--fg);outline:none}
  .input:focus{border-color:var(--accent)}
  .btn-accent{display:inline-flex;align-items:center;gap:.4rem;background:var(--accent);color:var(--accent-fg);font-size:.875rem;font-weight:600;padding:.45rem .8rem;border-radius:.5rem;cursor:pointer}
  .btn-accent:hover{opacity:.92}
`;
if (typeof document !== "undefined" && !document.getElementById("cgt-util")) { _style.id = "cgt-util"; document.head.appendChild(_style); }
