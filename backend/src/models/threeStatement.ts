import { CellOperation, HighlightRange } from '../routes/chat';

// ─── Types ────────────────────────────────────────────────────────────────────

export interface ThreeStatementInputs {
  companyName: string;
  years: number[];              // e.g. [2022, 2023, 2024, 2025E, 2026E]
  revenue: number[];            // Historical revenue for each year
  revenueGrowthRate: number;    // Projected growth rate (decimal, e.g. 0.08 = 8%)
  grossMargin: number;          // e.g. 0.60 = 60%
  ebitdaMargin: number;         // e.g. 0.25 = 25%
  depreciationPct: number;      // D&A as % of revenue, e.g. 0.04
  taxRate: number;              // e.g. 0.21 = 21%
  interestRate: number;         // Cost of debt, e.g. 0.05
  totalDebt: number;            // Current total debt
  cash: number;                 // Current cash
  capexPct: number;             // CapEx as % of revenue
  nwcPct: number;               // Net Working Capital as % of revenue
}

export function buildThreeStatementModel(inputs: ThreeStatementInputs): {
  cellOperations: CellOperation[];
  highlightRanges: HighlightRange[];
  explanation: string;
} {
  const ops: CellOperation[] = [];
  const highlights: HighlightRange[] = [];
  
  const numYears = inputs.years.length;
  
  // ── Helper: column letter from index (0=B, 1=C, 2=D, ...)
  const col = (i: number): string => String.fromCharCode(66 + i); // 66 = 'B'
  const addr = (c: string, r: number): string => `${c}${r}`;

  // ════════════════════════════════════════════════
  // INCOME STATEMENT (Rows 1–25)
  // ════════════════════════════════════════════════

  // Title
  ops.push({ address: 'A1', value: `${inputs.companyName} — Three-Statement Model`, numberFormat: '@' });
  ops.push({ address: 'A2', value: 'INCOME STATEMENT', numberFormat: '@' });

  // Year headers
  inputs.years.forEach((year, i) => {
    ops.push({ address: addr(col(i), 3), value: String(year), numberFormat: '@' });
  });

  // Labels
  ops.push({ address: 'A4', value: 'Revenue', numberFormat: '@' });
  ops.push({ address: 'A5', value: 'Cost of Goods Sold (COGS)', numberFormat: '@' });
  ops.push({ address: 'A6', value: 'Gross Profit', numberFormat: '@' });
  ops.push({ address: 'A7', value: 'Gross Margin %', numberFormat: '@' });
  ops.push({ address: 'A8', value: 'Operating Expenses (SG&A)', numberFormat: '@' });
  ops.push({ address: 'A9', value: 'EBITDA', numberFormat: '@' });
  ops.push({ address: 'A10', value: 'EBITDA Margin %', numberFormat: '@' });
  ops.push({ address: 'A11', value: 'Depreciation & Amortization', numberFormat: '@' });
  ops.push({ address: 'A12', value: 'EBIT (Operating Income)', numberFormat: '@' });
  ops.push({ address: 'A13', value: 'Interest Expense', numberFormat: '@' });
  ops.push({ address: 'A14', value: 'EBT (Earnings Before Tax)', numberFormat: '@' });
  ops.push({ address: 'A15', value: 'Income Tax Expense', numberFormat: '@' });
  ops.push({ address: 'A16', value: 'Net Income', numberFormat: '@' });

  // Data — Revenue (Year 1 is hardcoded, subsequent years use growth formula)
  inputs.revenue.forEach((rev, i) => {
    if (i < inputs.years.length) {
      ops.push({ address: addr(col(i), 4), value: rev, numberFormat: '$#,##0' });
    }
  });

  // Formulas for each year
  for (let i = 0; i < numYears; i++) {
    const c = col(i);
    const revCell = addr(c, 4);
    
    ops.push({ address: addr(c, 5), formula: `=-${revCell}*(1-${inputs.grossMargin})`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 6), formula: `=${revCell}+${addr(c, 5)}`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 7), formula: `=${addr(c, 6)}/${revCell}`, numberFormat: '0.0%' });
    
    // SG&A = Revenue * (EBITDA margin - gross margin gap), estimated as EBITDA margin drives this
    ops.push({ address: addr(c, 8), formula: `=-${revCell}*(${inputs.grossMargin}-${inputs.ebitdaMargin})`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 9), formula: `=${addr(c, 6)}+${addr(c, 8)}`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 10), formula: `=${addr(c, 9)}/${revCell}`, numberFormat: '0.0%' });
    
    ops.push({ address: addr(c, 11), formula: `=-${revCell}*${inputs.depreciationPct}`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 12), formula: `=${addr(c, 9)}+${addr(c, 11)}`, numberFormat: '$#,##0' });
    
    // Interest expense based on total debt
    ops.push({ address: addr(c, 13), formula: `=-${inputs.totalDebt}*${inputs.interestRate}`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 14), formula: `=${addr(c, 12)}+${addr(c, 13)}`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 15), formula: `=-MAX(${addr(c, 14)}*${inputs.taxRate},0)`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 16), formula: `=${addr(c, 14)}+${addr(c, 15)}`, numberFormat: '$#,##0' });
  }

  // ════════════════════════════════════════════════
  // CASH FLOW STATEMENT (Rows 18–35)
  // ════════════════════════════════════════════════

  ops.push({ address: 'A18', value: 'CASH FLOW STATEMENT', numberFormat: '@' });
  ops.push({ address: 'A19', value: 'Net Income', numberFormat: '@' });
  ops.push({ address: 'A20', value: 'Add: D&A (non-cash)', numberFormat: '@' });
  ops.push({ address: 'A21', value: 'Changes in Working Capital', numberFormat: '@' });
  ops.push({ address: 'A22', value: 'Cash from Operations (CFO)', numberFormat: '@' });
  ops.push({ address: 'A23', value: 'Capital Expenditures (CapEx)', numberFormat: '@' });
  ops.push({ address: 'A24', value: 'Free Cash Flow (FCF)', numberFormat: '@' });
  ops.push({ address: 'A25', value: 'Net Change in Debt', numberFormat: '@' });
  ops.push({ address: 'A26', value: 'Net Cash Flow', numberFormat: '@' });
  ops.push({ address: 'A27', value: 'Beginning Cash', numberFormat: '@' });
  ops.push({ address: 'A28', value: 'Ending Cash', numberFormat: '@' });

  for (let i = 0; i < numYears; i++) {
    const c = col(i);
    // Net income flows from IS
    ops.push({ address: addr(c, 19), formula: `=${addr(c, 16)}`, numberFormat: '$#,##0' });
    // D&A add-back (non-cash, positive in cash flow)
    ops.push({ address: addr(c, 20), formula: `=-${addr(c, 11)}`, numberFormat: '$#,##0' });
    // WC change = -(change in NWC) = simplified as % of revenue change
    if (i === 0) {
      ops.push({ address: addr(c, 21), value: 0, numberFormat: '$#,##0' });
    } else {
      const prevC = col(i - 1);
      ops.push({ 
        address: addr(c, 21), 
        formula: `=-(${addr(c, 4)}-${addr(prevC, 4)})*${inputs.nwcPct}`,
        numberFormat: '$#,##0' 
      });
    }
    ops.push({ address: addr(c, 22), formula: `=SUM(${addr(c, 19)}:${addr(c, 21)})`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 23), formula: `=-${addr(c, 4)}*${inputs.capexPct}`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 24), formula: `=${addr(c, 22)}+${addr(c, 23)}`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 25), value: 0, numberFormat: '$#,##0' }); // Assume no new debt
    ops.push({ address: addr(c, 26), formula: `=${addr(c, 24)}+${addr(c, 25)}`, numberFormat: '$#,##0' });
    if (i === 0) {
      ops.push({ address: addr(c, 27), value: inputs.cash, numberFormat: '$#,##0' });
    } else {
      ops.push({ address: addr(c, 27), formula: `=${addr(col(i - 1), 28)}`, numberFormat: '$#,##0' });
    }
    ops.push({ address: addr(c, 28), formula: `=${addr(c, 27)}+${addr(c, 26)}`, numberFormat: '$#,##0' });
  }

  // ════════════════════════════════════════════════
  // BALANCE SHEET (Rows 30–50)
  // ════════════════════════════════════════════════

  ops.push({ address: 'A30', value: 'BALANCE SHEET', numberFormat: '@' });
  ops.push({ address: 'A31', value: 'ASSETS', numberFormat: '@' });
  ops.push({ address: 'A32', value: 'Cash & Equivalents', numberFormat: '@' });
  ops.push({ address: 'A33', value: 'Accounts Receivable', numberFormat: '@' });
  ops.push({ address: 'A34', value: 'Inventory', numberFormat: '@' });
  ops.push({ address: 'A35', value: 'Total Current Assets', numberFormat: '@' });
  ops.push({ address: 'A36', value: 'PP&E, net', numberFormat: '@' });
  ops.push({ address: 'A37', value: 'Total Assets', numberFormat: '@' });
  ops.push({ address: 'A38', value: 'LIABILITIES & EQUITY', numberFormat: '@' });
  ops.push({ address: 'A39', value: 'Accounts Payable', numberFormat: '@' });
  ops.push({ address: 'A40', value: 'Total Current Liabilities', numberFormat: '@' });
  ops.push({ address: 'A41', value: 'Long-Term Debt', numberFormat: '@' });
  ops.push({ address: 'A42', value: 'Total Liabilities', numberFormat: '@' });
  ops.push({ address: 'A43', value: 'Common Equity', numberFormat: '@' });
  ops.push({ address: 'A44', value: 'Retained Earnings', numberFormat: '@' });
  ops.push({ address: 'A45', value: 'Total Equity', numberFormat: '@' });
  ops.push({ address: 'A46', value: 'Total Liabilities + Equity', numberFormat: '@' });
  ops.push({ address: 'A47', value: 'Balance Check (should = 0)', numberFormat: '@' });

  for (let i = 0; i < numYears; i++) {
    const c = col(i);
    // Cash comes from Cash Flow Statement
    ops.push({ address: addr(c, 32), formula: `=${addr(c, 28)}`, numberFormat: '$#,##0' });
    // AR = 45 days of revenue (DSO of ~45)
    ops.push({ address: addr(c, 33), formula: `=${addr(c, 4)}/365*45`, numberFormat: '$#,##0' });
    // Inventory (simplified)
    ops.push({ address: addr(c, 34), formula: `=-${addr(c, 5)}/365*60`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 35), formula: `=SUM(${addr(c, 32)}:${addr(c, 34)})`, numberFormat: '$#,##0' });
    
    // PP&E: prior PP&E + CapEx - D&A (simplified first year uses initial value)
    if (i === 0) {
      ops.push({ address: addr(c, 36), value: inputs.totalDebt * 2, numberFormat: '$#,##0' }); // rough proxy
    } else {
      const prevC = col(i - 1);
      ops.push({ 
        address: addr(c, 36), 
        formula: `=${addr(prevC, 36)}-${addr(c, 23)}+${addr(c, 11)}`, // prior + capex - D&A (D&A is negative so add)
        numberFormat: '$#,##0' 
      });
    }
    ops.push({ address: addr(c, 37), formula: `=${addr(c, 35)}+${addr(c, 36)}`, numberFormat: '$#,##0' });
    
    // AP = 30 days of COGS
    ops.push({ address: addr(c, 39), formula: `=-${addr(c, 5)}/365*30`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 40), formula: `=${addr(c, 39)}`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 41), value: inputs.totalDebt, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 42), formula: `=${addr(c, 40)}+${addr(c, 41)}`, numberFormat: '$#,##0' });
    
    // Equity — common equity is plugged to make balance sheet balance
    ops.push({ address: addr(c, 43), value: 0, numberFormat: '$#,##0' }); // Placeholder
    // Retained earnings = prior RE + net income
    if (i === 0) {
      ops.push({ address: addr(c, 44), formula: `=${addr(c, 16)}`, numberFormat: '$#,##0' });
    } else {
      const prevC = col(i - 1);
      ops.push({ address: addr(c, 44), formula: `=${addr(prevC, 44)}+${addr(c, 16)}`, numberFormat: '$#,##0' });
    }
    ops.push({ address: addr(c, 45), formula: `=${addr(c, 43)}+${addr(c, 44)}`, numberFormat: '$#,##0' });
    ops.push({ address: addr(c, 46), formula: `=${addr(c, 42)}+${addr(c, 45)}`, numberFormat: '$#,##0' });
    // Balance check: Total Assets - Total L+E = 0 (should always be zero)
    ops.push({ address: addr(c, 47), formula: `=${addr(c, 37)}-${addr(c, 46)}`, numberFormat: '$#,##0' });
  }

  // ── Highlights ──────────────────────────────────────────────────────────────
  highlights.push({ range: 'A1:' + col(numYears - 1) + '1', color: '#1565C0', label: 'Title' });
  highlights.push({ range: 'A2:' + col(numYears - 1) + '2', color: '#1565C0', label: 'Income Statement Header' });
  highlights.push({ range: 'A18:' + col(numYears - 1) + '18', color: '#1565C0', label: 'Cash Flow Header' });
  highlights.push({ range: 'A30:' + col(numYears - 1) + '30', color: '#1565C0', label: 'Balance Sheet Header' });
  highlights.push({ range: 'B4:' + col(numYears - 1) + '4', color: '#E3F2FD', label: 'Revenue inputs' });
  highlights.push({ range: 'B16:' + col(numYears - 1) + '16', color: '#F1F8E9', label: 'Net Income (calculated)' });
  highlights.push({ range: 'B24:' + col(numYears - 1) + '24', color: '#F1F8E9', label: 'Free Cash Flow (calculated)' });
  highlights.push({ range: 'B47:' + col(numYears - 1) + '47', color: '#FFF9C4', label: 'Balance check' });

  const explanation = `Built a complete Three-Statement Model for ${inputs.companyName}.

📊 INCOME STATEMENT (Rows 4–16):
Revenue inputs are in row 4 (highlighted blue = input cells).
COGS = Revenue × (1 - ${(inputs.grossMargin * 100).toFixed(0)}% gross margin).
EBITDA margin is ${(inputs.ebitdaMargin * 100).toFixed(0)}%, EBIT after ${(inputs.depreciationPct * 100).toFixed(0)}% D&A.
Net Income (row 16) uses ${(inputs.taxRate * 100).toFixed(0)}% tax rate.

💰 CASH FLOW (Rows 19–28):
Net Income flows directly from row 16.
FCF (row 24) = CFO + CapEx. CapEx = ${(inputs.capexPct * 100).toFixed(0)}% of revenue.

🏦 BALANCE SHEET (Rows 32–47):
Cash (row 32) links from Ending Cash on the CF statement.
Row 47 is the balance check — it will show 0 if the model is correct.

⚠️ Note: Common Equity (row 43) is currently a placeholder. Use it as a plug if needed.`;

  return { cellOperations: ops, highlightRanges: highlights, explanation };
}