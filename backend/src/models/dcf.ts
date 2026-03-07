import { CellOperation, HighlightRange } from '../routes/chat';

export interface DCFInputs {
  companyName: string;
  freeCashFlows: number[];   // Historical/projected FCFs for 5 years
  wacc: number;              // Weighted Average Cost of Capital (decimal)
  terminalGrowthRate: number; // Long-term growth rate (decimal, e.g. 0.025)
  netDebt: number;           // Total Debt - Cash (can be negative if net cash)
  sharesOutstanding: number; // In millions
}

export function buildDCFModel(inputs: DCFInputs): {
  cellOperations: CellOperation[];
  highlightRanges: HighlightRange[];
  explanation: string;
} {
  const ops: CellOperation[] = [];
  const highlights: HighlightRange[] = [];
  const col = (i: number) => String.fromCharCode(66 + i);

  // ── Title & Headers ─────────────────────────────────────────────────────────
  ops.push({ address: 'A1', value: `${inputs.companyName} — DCF Valuation`, numberFormat: '@' });
  ops.push({ address: 'A3', value: 'ASSUMPTIONS', numberFormat: '@' });
  ops.push({ address: 'A4', value: 'WACC', numberFormat: '@' });
  ops.push({ address: 'B4', value: inputs.wacc, numberFormat: '0.0%' });
  ops.push({ address: 'A5', value: 'Terminal Growth Rate', numberFormat: '@' });
  ops.push({ address: 'B5', value: inputs.terminalGrowthRate, numberFormat: '0.0%' });
  ops.push({ address: 'A6', value: 'Net Debt ($M)', numberFormat: '@' });
  ops.push({ address: 'B6', value: inputs.netDebt, numberFormat: '$#,##0.0' });
  ops.push({ address: 'A7', value: 'Shares Outstanding (M)', numberFormat: '@' });
  ops.push({ address: 'B7', value: inputs.sharesOutstanding, numberFormat: '#,##0.0' });

  // ── FCF Projections ─────────────────────────────────────────────────────────
  ops.push({ address: 'A9', value: 'FREE CASH FLOW PROJECTIONS', numberFormat: '@' });
  const years = ['Year 1', 'Year 2', 'Year 3', 'Year 4', 'Year 5'];
  years.forEach((y, i) => {
    ops.push({ address: `${col(i)}10`, value: y, numberFormat: '@' });
  });

  ops.push({ address: 'A11', value: 'Free Cash Flow ($M)', numberFormat: '@' });
  inputs.freeCashFlows.forEach((fcf, i) => {
    ops.push({ address: `${col(i)}11`, value: fcf, numberFormat: '$#,##0.0' });
  });

  // ── Discount Factors ─────────────────────────────────────────────────────────
  // Discount factor = 1 / (1 + WACC)^n
  // We reference B4 (WACC) with $ anchors so all 5 columns use the same WACC
  ops.push({ address: 'A12', value: 'Discount Factor (1/(1+WACC)^n)', numberFormat: '@' });
  ops.push({ address: 'A13', value: 'PV of FCF ($M)', numberFormat: '@' });
  
  for (let i = 0; i < 5; i++) {
    const c = col(i);
    const n = i + 1;
    // Formula: 1 / (1 + WACC)^n — WACC is in B4
    ops.push({ address: `${c}12`, formula: `=1/(1+$B$4)^${n}`, numberFormat: '0.0000' });
    ops.push({ address: `${c}13`, formula: `=${c}11*${c}12`, numberFormat: '$#,##0.0' });
  }

  // ── Terminal Value ───────────────────────────────────────────────────────────
  // Terminal Value = FCF_Year5 * (1 + g) / (WACC - g)
  // This is the Gordon Growth Model formula. It captures all value beyond year 5.
  ops.push({ address: 'A15', value: 'TERMINAL VALUE', numberFormat: '@' });
  ops.push({ address: 'A16', value: 'Terminal Value ($M)', numberFormat: '@' });
  // F11 = Year 5 FCF, B5 = terminal growth rate, B4 = WACC
  ops.push({ address: 'B16', formula: `=F11*(1+$B$5)/($B$4-$B$5)`, numberFormat: '$#,##0.0' });
  ops.push({ address: 'A17', value: 'PV of Terminal Value ($M)', numberFormat: '@' });
  // Discount TV back 5 years
  ops.push({ address: 'B17', formula: `=B16/(1+$B$4)^5`, numberFormat: '$#,##0.0' });

  // ── Valuation Summary ────────────────────────────────────────────────────────
  ops.push({ address: 'A19', value: 'VALUATION SUMMARY', numberFormat: '@' });
  ops.push({ address: 'A20', value: 'Sum of PV of FCFs ($M)', numberFormat: '@' });
  ops.push({ address: 'B20', formula: `=SUM(B13:F13)`, numberFormat: '$#,##0.0' });
  ops.push({ address: 'A21', value: 'PV of Terminal Value ($M)', numberFormat: '@' });
  ops.push({ address: 'B21', formula: `=B17`, numberFormat: '$#,##0.0' });
  ops.push({ address: 'A22', value: 'Enterprise Value ($M)', numberFormat: '@' });
  ops.push({ address: 'B22', formula: `=B20+B21`, numberFormat: '$#,##0.0' });
  ops.push({ address: 'A23', value: 'Less: Net Debt ($M)', numberFormat: '@' });
  ops.push({ address: 'B23', formula: `=B6`, numberFormat: '$#,##0.0' });
  ops.push({ address: 'A24', value: 'Equity Value ($M)', numberFormat: '@' });
  ops.push({ address: 'B24', formula: `=B22-B23`, numberFormat: '$#,##0.0' });
  ops.push({ address: 'A25', value: 'Implied Share Price ($)', numberFormat: '@' });
  // Equity Value / Shares Outstanding — both in millions, result is $/share
  ops.push({ address: 'B25', formula: `=B24/B7`, numberFormat: '$#,##0.00' });

  // ── Sensitivity Table: Share Price vs WACC / Growth Rate ────────────────────
  // A 5×5 sensitivity table is standard in real investment banking
  ops.push({ address: 'A27', value: 'SENSITIVITY ANALYSIS — Share Price vs WACC (rows) / Terminal Growth (cols)', numberFormat: '@' });
  
  const waccRange = [-0.02, -0.01, 0, 0.01, 0.02]; // ±2% around base WACC
  const growthRange = [-0.01, -0.005, 0, 0.005, 0.01]; // ±1% around base growth

  // Column headers (terminal growth rates)
  growthRange.forEach((delta, i) => {
    const growthVal = inputs.terminalGrowthRate + delta;
    ops.push({ address: `${col(i + 1)}28`, value: `${(growthVal * 100).toFixed(1)}%`, numberFormat: '@' });
  });
  ops.push({ address: 'A28', value: 'WACC \\ Terminal g', numberFormat: '@' });

  // Row headers and sensitivity formulas
  waccRange.forEach((wDelta, wi) => {
    const waccVal = inputs.wacc + wDelta;
    ops.push({ address: `A${29 + wi}`, value: `${(waccVal * 100).toFixed(1)}%`, numberFormat: '@' });
    
    growthRange.forEach((gDelta, gi) => {
      const growthVal = inputs.terminalGrowthRate + gDelta;
      // Re-calculate DCF inline for each combination
      // TV = FCF_5 * (1+g) / (WACC-g)  →  PV of TV = TV / (1+WACC)^5
      // EV = sum PV of FCFs + PV of TV
      // Note: FCFs are fixed regardless of WACC/growth changes in this sensitivity
      const waccFixed = waccVal;
      const pvFCFs = inputs.freeCashFlows.reduce((sum, fcf, n) => {
        return sum + fcf / Math.pow(1 + waccFixed, n + 1);
      }, 0);
      const tv = inputs.freeCashFlows[4] * (1 + growthVal) / (waccFixed - growthVal);
      const pvTV = tv / Math.pow(1 + waccFixed, 5);
      const ev = pvFCFs + pvTV;
      const price = (ev - inputs.netDebt) / inputs.sharesOutstanding;
      
      ops.push({ 
        address: `${col(gi + 1)}${29 + wi}`, 
        value: Math.max(0, price),
        numberFormat: '$#,##0.00'
      });
    });
  });

  // ── Highlights ────────────────────────────────────────────────────────────────
  highlights.push({ range: 'A1:F1', color: '#1565C0', label: 'Title' });
  highlights.push({ range: 'A3:B3', color: '#1565C0', label: 'Assumptions header' });
  highlights.push({ range: 'B4:B7', color: '#E3F2FD', label: 'Input assumptions' });
  highlights.push({ range: 'B11:F11', color: '#E3F2FD', label: 'FCF inputs' });
  highlights.push({ range: 'B13:F13', color: '#F1F8E9', label: 'PV of FCFs' });
  highlights.push({ range: 'B16:B17', color: '#F1F8E9', label: 'Terminal Value' });
  highlights.push({ range: 'A19:B19', color: '#1565C0', label: 'Valuation summary header' });
  highlights.push({ range: 'B25', color: '#FFECB3', label: '⭐ Implied Share Price' });
  highlights.push({ range: 'A27:F33', color: '#F3E5F5', label: 'Sensitivity analysis' });

  const tvPct = Math.round(
    (inputs.freeCashFlows[4] * (1 + inputs.terminalGrowthRate) / (inputs.wacc - inputs.terminalGrowthRate)) /
    (Math.pow(1 + inputs.wacc, 5)) /
    (inputs.freeCashFlows.reduce((s, f, n) => s + f / Math.pow(1 + inputs.wacc, n + 1), 0) +
     inputs.freeCashFlows[4] * (1 + inputs.terminalGrowthRate) / (inputs.wacc - inputs.terminalGrowthRate) / Math.pow(1 + inputs.wacc, 5)) * 100
  );

  const explanation = `Built a DCF Valuation for ${inputs.companyName}.

📐 MATH EXPLAINED:
- Discount factors (row 12): = 1/(1+WACC)^n using B4=${(inputs.wacc*100).toFixed(1)}% WACC
- PV of FCF (row 13): Each FCF × its discount factor
- Terminal Value (B16): = FCF_Year5 × (1+${(inputs.terminalGrowthRate*100).toFixed(1)}%) / (${(inputs.wacc*100).toFixed(1)}% - ${(inputs.terminalGrowthRate*100).toFixed(1)}%)
  This is the Gordon Growth Model — it values ALL cash flows beyond Year 5
- Enterprise Value (B22) = PV of FCFs + PV of Terminal Value
- Equity Value (B24) = Enterprise Value - Net Debt ($${inputs.netDebt}M)
- Implied Price (B25) = Equity Value / ${inputs.sharesOutstanding}M shares

⚠️ Terminal value represents ~${tvPct}% of enterprise value — this is typical but sensitive to your WACC and growth assumptions.

📊 The sensitivity table (rows 28-33) shows implied price for ±2% WACC and ±1% growth rate scenarios.`;

  return { cellOperations: ops, highlightRanges: highlights, explanation };
}