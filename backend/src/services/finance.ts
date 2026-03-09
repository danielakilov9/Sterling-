// finance.ts — Data service: Yahoo Finance (stock data) + FRED (macro data)
//
// Design note on "as any":
//   yahoo-finance2 exposes complex internal union types for QuoteSummaryResult
//   and Quote. Every field is optional, and the module-specific subfields are
//   only present at runtime (not statically provable). Since we are piping
//   this data straight to the AI as JSON — not doing typed computation on it —
//   casting at the library boundary with "as any" is the correct engineering
//   tradeoff. It eliminates all type errors without compromising safety.

import yahooFinance from 'yahoo-finance2';
import axios from 'axios';
import dotenv from 'dotenv';

dotenv.config();

// ─── Yahoo Finance ─────────────────────────────────────────────────────────────

export async function getStockData(ticker: string, dataType: string): Promise<object> {
  try {
    // Cast to any immediately after the API call.
    // yahoo-finance2's Quote type marks every field as optional (number | undefined).
    // Using "as any" here means we access .regularMarketPrice, .marketCap, etc.
    // without TypeScript complaining about possible undefined on every property —
    // those fields ARE present at runtime for any valid ticker.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const quote = await yahooFinance.quote(ticker) as any;

    // ── 'quote' — lightweight real-time snapshot ──────────────────────────────
    if (dataType === 'quote') {
      return {
        ticker,
        price:     quote.regularMarketPrice ?? null,
        marketCap: quote.marketCap          ?? null,
        peRatio:   quote.trailingPE         ?? null,
        high52w:   quote.fiftyTwoWeekHigh   ?? null,
        low52w:    quote.fiftyTwoWeekLow    ?? null,
      };
    }

    // ── 'all' — full fundamentals for model building ──────────────────────────
    if (dataType === 'all') {
      // All four calls use quoteSummary → same return type → no type mismatch.
      // Promise.allSettled means one failed module does not abort the others.
      const [summaryResult, incomeResult, balanceResult, cashflowResult] =
        await Promise.allSettled([
          yahooFinance.quoteSummary(ticker, {
            modules: ['financialData', 'defaultKeyStatistics'],
          }),
          yahooFinance.quoteSummary(ticker, {
            modules: ['incomeStatementHistory'],
          }),
          yahooFinance.quoteSummary(ticker, {
            modules: ['balanceSheetHistory'],
          }),
          yahooFinance.quoteSummary(ticker, {
            modules: ['cashflowStatementHistory'],
          }),
        ]);

      // Narrow each PromiseSettledResult to its value with "as any".
      // After the ternary status check the value IS the QuoteSummaryResult, but
      // TypeScript still requires a cast to access module-specific sub-fields
      // (financialData, defaultKeyStatistics, etc.) because they are all
      // declared optional on QuoteSummaryResult — present only when requested.
      // "as any" on each resolved value is the correct fix for all four blocks.

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const summary  = summaryResult.status  === 'fulfilled' ? summaryResult.value  as any : null;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const income   = incomeResult.status   === 'fulfilled' ? incomeResult.value   as any : null;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const balance  = balanceResult.status  === 'fulfilled' ? balanceResult.value  as any : null;
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const cashflow = cashflowResult.status === 'fulfilled' ? cashflowResult.value as any : null;

      return {
        ticker,
        quote: {
          price:     quote.regularMarketPrice ?? null,
          marketCap: quote.marketCap          ?? null,
          peRatio:   quote.trailingPE         ?? null,
          high52w:   quote.fiftyTwoWeekHigh   ?? null,
          low52w:    quote.fiftyTwoWeekLow    ?? null,
          // priceToBook and sharesOutstanding live in defaultKeyStatistics,
          // NOT on the Quote object — this was error #1 and #2 in earlier versions.
          priceToBook:       summary?.defaultKeyStatistics?.priceToBook       ?? null,
          sharesOutstanding: summary?.defaultKeyStatistics?.sharesOutstanding ?? null,
        },
        financialData:   summary?.financialData     ?? null,
        keyStats:        summary?.defaultKeyStatistics ?? null,
        // Income statement was missing entirely from earlier versions (errors #4 & #5).
        incomeStatement: income?.incomeStatementHistory?.incomeStatementHistory?.slice(0, 4) ?? null,
        balanceSheet:    balance?.balanceSheetHistory?.balanceSheetStatements?.slice(0, 4)   ?? null,
        cashflow:        cashflow?.cashflowStatementHistory?.cashflowStatements?.slice(0, 4) ?? null,
      };
    }

    // ── Single-module fetches ─────────────────────────────────────────────────
    // We call quoteSummary directly in each branch rather than building a
    // dynamic array. The old "const modules = [] as Parameters<...>[1]['modules']"
    // pattern was the root cause of errors #1-#9 in the previous version:
    //   Parameters<...>[1] is optional → type is QuoteSummaryOptions | undefined
    //   Indexing (T | undefined)['modules'] → TS error #1
    //   Every .push() call → errors #2-#7
    //   .length → error #8
    //   passing { modules } to quoteSummary → error #9

    if (dataType === 'income') {
      const result = await yahooFinance.quoteSummary(ticker, {
        modules: ['incomeStatementHistory'],
      });
      return { ticker, data: result };
    }

    if (dataType === 'balance') {
      const result = await yahooFinance.quoteSummary(ticker, {
        modules: ['balanceSheetHistory'],
      });
      return { ticker, data: result };
    }

    if (dataType === 'cashflow') {
      const result = await yahooFinance.quoteSummary(ticker, {
        modules: ['cashflowStatementHistory'],
      });
      return { ticker, data: result };
    }

    return {
      ticker,
      error: `Unknown dataType: "${dataType}". Valid: quote | all | income | balance | cashflow`,
    };

  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(`Yahoo Finance error for ${ticker}:`, message);
    return {
      ticker,
      error:    message,
      fallback: true,
      note:     'Could not fetch live data. AI will use reasonable placeholder assumptions.',
    };
  }
}

// ─── FRED (Federal Reserve Economic Data) ─────────────────────────────────────
// Docs: https://fred.stlouisfed.org/docs/api/fred/
//
// Common series IDs:
//   FEDFUNDS  — Federal Funds Effective Rate
//   CPIAUCSL  — CPI (inflation proxy)
//   GDP       — US GDP
//   UNRATE    — Unemployment Rate
//   DGS10     — 10-Year Treasury Yield (risk-free rate for CAPM in DCF)
//   BAA       — Moody's Baa Corporate Bond Yield (proxy for cost of debt)

export async function getMacroData(seriesId: string): Promise<object> {
  const apiKey = process.env.FRED_API_KEY;

  if (!apiKey) {
    console.warn('FRED_API_KEY not set in .env — returning placeholder data');
    return getPlaceholderMacroData(seriesId);
  }

  try {
    const response = await axios.get<{
      observations: Array<{ date: string; value: string }>;
    }>('https://api.stlouisfed.org/fred/series/observations', {
      params: {
        series_id:         seriesId,
        api_key:           apiKey,
        file_type:         'json',
        sort_order:        'desc',
        limit:             12,
        observation_start: getOneYearAgo(),
      },
    });

    // FRED uses '.' as a sentinel for missing observations — filter those out.
    const valid = response.data.observations
      .filter((o) => o.value !== '.')
      .map((o)    => ({ date: o.date, value: parseFloat(o.value) }));

    const latest = valid[0];

    return {
      seriesId,
      latestValue: latest?.value ?? null,
      latestDate:  latest?.date  ?? null,
      history:     valid.slice(0, 6),
      unit:        getSeriesUnit(seriesId),
    };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(`FRED API error for ${seriesId}:`, message);
    return {
      seriesId,
      error:    message,
      fallback: true,
      ...getPlaceholderMacroData(seriesId),
    };
  }
}

// ─── Private helpers ───────────────────────────────────────────────────────────

function getOneYearAgo(): string {
  const d = new Date();
  d.setFullYear(d.getFullYear() - 1);
  return d.toISOString().split('T')[0]; // "YYYY-MM-DD"
}

function getSeriesUnit(seriesId: string): string {
  const units: Record<string, string> = {
    FEDFUNDS: 'Percent',
    CPIAUCSL: 'Index',
    GDP:      'Billions of Dollars',
    UNRATE:   'Percent',
    DGS10:    'Percent',
    BAA:      'Percent',
  };
  return units[seriesId] ?? 'See FRED for units';
}

// Reasonable real-world defaults so the AI can build models before FRED key setup.
function getPlaceholderMacroData(seriesId: string): object {
  const defaults: Record<string, { latestValue: number; note: string }> = {
    FEDFUNDS: { latestValue: 5.33,  note: 'Placeholder: Fed Funds Rate ~5.33%' },
    DGS10:    { latestValue: 4.25,  note: 'Placeholder: 10Y Treasury ~4.25% (use as risk-free rate in CAPM)' },
    CPIAUCSL: { latestValue: 314.0, note: 'Placeholder: CPI index ~314' },
    GDP:      { latestValue: 27500, note: 'Placeholder: US GDP ~$27.5T' },
    UNRATE:   { latestValue: 4.1,   note: 'Placeholder: Unemployment ~4.1%' },
    BAA:      { latestValue: 5.8,   note: 'Placeholder: Baa yield ~5.8% (cost of debt proxy)' },
  };
  return (
    defaults[seriesId] ?? {
      latestValue: null,
      note: `No placeholder for "${seriesId}". Register a free key at fred.stlouisfed.org`,
    }
  );
}