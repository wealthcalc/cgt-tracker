// Vercel serverless function: Yahoo Finance quote proxy.
//
// Why this exists: Yahoo's endpoints can't be called from the browser (no CORS,
// plus a cookie/crumb handshake). This runs server-side, so the dashboard can
// fetch live prices via same-origin GET /api/quotes?symbols=SWDA.L,WFC
//
// Returns per-symbol: { symbol, price, currency, name } or { symbol, error }.
// The client normalises to GBP (Yahoo returns "GBp" for pence-quoted LSE lines,
// "USD"/"EUR" for others), so this function stays thin and stateless.
//
// Deploy: place at /api/quotes.mjs and add "yahoo-finance2" to dependencies.
// The .mjs extension forces ESM regardless of your package.json "type".

import YahooFinance from "yahoo-finance2";

const yf = new YahooFinance({ suppressNotices: ["yahooSurvey"] });

export default async function handler(req, res) {
  const raw = (req.query?.symbols ?? "").toString();
  const symbols = raw.split(",").map((s) => s.trim()).filter(Boolean).slice(0, 50);
  if (!symbols.length) {
    res.status(400).json({ error: "Pass ?symbols=AAA.L,BBB (comma-separated)." });
    return;
  }

  const quotes = await Promise.all(
    symbols.map(async (sym) => {
      try {
        const q = await yf.quote(sym);
        if (!q || q.regularMarketPrice == null) return { symbol: sym, error: "no data" };
        return {
          symbol: sym,
          price: q.regularMarketPrice,
          currency: q.currency || null,          // e.g. "GBp", "GBP", "USD"
          name: q.shortName || q.longName || null,
        };
      } catch (e) {
        return { symbol: sym, error: (e && e.message) || "fetch failed" };
      }
    })
  );

  // Cache at the edge briefly so repeated refreshes don't hammer Yahoo.
  res.setHeader("Cache-Control", "s-maxage=60, stale-while-revalidate=300");
  res.status(200).json({ quotes });
}
