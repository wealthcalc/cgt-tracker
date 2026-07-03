// Vercel serverless function: Yahoo Finance historical FX rate proxy.
//
// Second-tier FX fallback behind Frankfurter (frankfurter.dev is the
// primary source; this covers the gap when Frankfurter is down or lacks a
// rate for an older/less common date or currency). Alpha Vantage remains
// the last resort, since its FX calls share the same 25/day budget as
// equity price lookups.
//
// Yahoo doesn't expose an FX rate directly for "1 unit of CCY in GBP" — it
// lists currency-pair tickers like "USDGBP=X" (1 USD in GBP) and "GBPUSD=X"
// (1 GBP in USD). This tries the direct cross first (CCYGBP=X), then falls
// back to the inverse of the reverse cross (1 / GBPCCY=X) if the direct
// listing has no data, which happens for some minor-currency pairs.
//
// GET /api/fx?ccy=USD&date=2024-05-31 -> { ccy, date, rate, source }
// `rate` is GBP per 1 unit of `ccy`, matching Frankfurter's convention so the
// client can swap providers with no change to how the number is used.

import YahooFinance from "yahoo-finance2";

const yf = new YahooFinance({ suppressNotices: ["yahooSurvey"] });

async function closeOnOrBefore(ticker, dateStr) {
  const target = new Date(dateStr + "T00:00:00Z");
  const period2 = new Date(target); period2.setUTCDate(period2.getUTCDate() + 1);
  const period1 = new Date(target); period1.setUTCDate(period1.getUTCDate() - 9); // covers weekends + a holiday or two
  const result = await yf.chart(ticker, { period1, period2, interval: "1d" });
  const quotes = (result?.quotes || []).filter((q) => q.close != null);
  if (!quotes.length) return null;
  const targetMs = target.getTime() + 86400000 - 1; // end of target day
  const onOrBefore = quotes.filter((q) => new Date(q.date).getTime() <= targetMs);
  const pick = (onOrBefore.length ? onOrBefore : quotes).slice(-1)[0];
  return pick && isFinite(pick.close) ? pick.close : null;
}

export default async function handler(req, res) {
  const ccy = (req.query?.ccy ?? "").toString().trim().toUpperCase();
  const date = (req.query?.date ?? "").toString().trim();
  if (!/^[A-Z]{3}$/.test(ccy) || !/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    res.status(400).json({ error: "Pass ?ccy=USD&date=YYYY-MM-DD" });
    return;
  }
  if (ccy === "GBP") { res.status(200).json({ ccy, date, rate: 1, source: "identity" }); return; }

  try {
    let rate = await closeOnOrBefore(`${ccy}GBP=X`, date);
    let source = `${ccy}GBP=X`;
    if (rate == null) {
      const inv = await closeOnOrBefore(`GBP${ccy}=X`, date);
      if (inv) { rate = 1 / inv; source = `1/GBP${ccy}=X`; }
    }
    if (rate == null) { res.status(404).json({ error: `No Yahoo FX data for ${ccy} around ${date}.` }); return; }
    res.setHeader("Cache-Control", "s-maxage=3600, stale-while-revalidate=86400");
    res.status(200).json({ ccy, date, rate, source });
  } catch (e) {
    res.status(502).json({ error: (e && e.message) || "fetch failed" });
  }
}
