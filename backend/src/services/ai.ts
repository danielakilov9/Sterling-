import OpenAI from 'openai';
import { ChatResponse, CellOperation, HighlightRange, CellReference } from '../routes/chat';
import { getStockData, getMacroData } from './finance';

// ─── Client Setup ─────────────────────────────────────────────────────────────
// We use the OpenAI SDK for BOTH Ollama and Groq — they both speak the OpenAI API format.
// This means one codebase works with either backend.

function getAIClient(preferLocal: boolean = true): { client: OpenAI, model: string } {
  if (preferLocal) {
    // Ollama runs locally on port 11434
    return {
      client: new OpenAI({
        baseURL: 'http://localhost:11434/v1',
        apiKey: 'ollama', // Ollama doesn't need a real key, but the SDK requires something
      }),
      model: 'llama3.1'
    };
  } else {
    // Groq cloud
    return {
      client: new OpenAI({
        baseURL: 'https://api.groq.com/openai/v1',
        apiKey: process.env.GROQ_API_KEY!,
      }),
      model: 'llama-3.1-70b-versatile'
    };
  }
}

// ─── System Prompt ────────────────────────────────────────────────────────────
// This is the most critical part of the AI setup.
// It instructs the AI how to respond and when to use tools.

const SYSTEM_PROMPT = `You are a professional financial analyst and Excel modeling expert embedded
inside Microsoft Excel. Your job is to help users build accurate financial models.

CRITICAL RULES:
1. When building or modifying spreadsheet models, you MUST respond with a JSON object.
2. Every cell you reference in your explanation MUST be included in the references array.
3. All formulas MUST use standard Excel syntax (e.g. =SUM(B2:B10), =B2/C2).
4. Financial models follow these standard layouts:
   - Row 1: Title/header
   - Row 2: Blank separator
   - Row 3+: Data begins
   - Column A: Labels/descriptions
   - Column B onward: Years or categories

COLOR CODING (use these exact hex values for highlight colors):
- Headers/titles: #1565C0 (dark blue) with white text
- Input cells (hard-coded numbers): #E3F2FD (light blue)
- Calculated cells (formulas): #F1F8E9 (light green)  
- Warning/assumption cells: #FFF9C4 (light yellow)
- Negative/cost items: #FFEBEE (light red)

RESPONSE FORMAT:
Always return a JSON object with this exact structure:
{
  "message": "Plain English explanation of what you built and why",
  "cellOperations": [
    { "address": "A1", "value": "Revenue Model", "numberFormat": "@" },
    { "address": "B3", "formula": "=B2*1.05", "numberFormat": "$#,##0" }
  ],
  "highlightRanges": [
    { "range": "A1:E1", "color": "#1565C0", "label": "Header row" },
    { "range": "B3:E10", "color": "#F1F8E9", "label": "Calculated cells" }
  ],
  "references": [
    { "address": "B3", "description": "Revenue Year 1" },
    { "address": "C3", "description": "Revenue Year 2 (=B3 × 1.05 growth)" }
  ]
}

If the user asks a question without needing to write to cells, return:
{
  "message": "Your answer here",
  "cellOperations": [],
  "highlightRanges": [],
  "references": []
}

FINANCIAL MODEL ACCURACY REQUIREMENTS:
- DCF: Use WACC as discount rate. Terminal value = FCF_final × (1+g) / (WACC-g). 
  Enterprise Value = PV of FCFs + PV of Terminal Value.
  Equity Value = Enterprise Value - Net Debt.
- 3-Statement: Income Statement drives Balance Sheet and Cash Flow Statement.
  Net Income flows to Retained Earnings on Balance Sheet.
  Cash Flow from Operations reconciles Net Income to cash.
  Balance Sheet MUST balance: Assets = Liabilities + Equity.
- LBO: Entry equity = Enterprise Value × (1 - Debt%) 
  IRR = (Exit Equity / Entry Equity)^(1/years) - 1
  MOIC = Exit Equity / Entry Equity`;

// ─── Tool Definitions ─────────────────────────────────────────────────────────
// These are functions the AI can DECIDE to call when it needs financial data.

const tools: OpenAI.Chat.ChatCompletionTool[] = [
  {
    type: 'function',
    function: {
      name: 'get_stock_data',
      description: 'Get current stock price, financial statements, and key metrics for a publicly traded company. Use this when the user mentions a stock ticker or company name.',
      parameters: {
        type: 'object',
        properties: {
          ticker: {
            type: 'string',
            description: 'Stock ticker symbol, e.g. AAPL, MSFT, TSLA'
          },
          dataType: {
            type: 'string',
            enum: ['quote', 'income', 'balance', 'cashflow', 'all'],
            description: 'What financial data to retrieve'
          }
        },
        required: ['ticker', 'dataType']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_macro_data',
      description: 'Get macroeconomic indicators from the Federal Reserve (FRED). Use for interest rates, GDP growth, inflation when needed for DCF assumptions.',
      parameters: {
        type: 'object',
        properties: {
          seriesId: {
            type: 'string',
            description: 'FRED series ID. Common ones: FEDFUNDS (Fed Funds Rate), CPIAUCSL (CPI), GDP, UNRATE (unemployment)'
          }
        },
        required: ['seriesId']
      }
    }
  }
];

// ─── Main Agent Function ───────────────────────────────────────────────────────

interface AgentInput {
  message: string;
  conversationHistory: Array<{ role: 'user' | 'assistant'; content: string }>;
  spreadsheetContext?: {
    activeCell: string;
    selectedRange: string;
    sheetName: string;
    existingData?: Array<{ address: string; value: string | number | null; formula?: string }>;
  };
}

export async function runAIAgent(input: AgentInput): Promise<ChatResponse> {
  // Try local Ollama first, fall back to Groq if it fails
  let useLocal = true;
  try {
    const response = await fetch('http://localhost:11434/api/tags');
    if (!response.ok) useLocal = false;
  } catch {
    useLocal = false; // Ollama not running, use Groq
  }

  const { client, model } = getAIClient(useLocal);
  console.log(`Using AI backend: ${useLocal ? 'Ollama (local)' : 'Groq (cloud)'}`);

  // Build the context message if the spreadsheet has existing data
  let contextMessage = '';
  if (input.spreadsheetContext) {
    const ctx = input.spreadsheetContext;
    contextMessage = `\n\nCurrent spreadsheet context:
- Sheet name: ${ctx.sheetName}
- Active cell: ${ctx.activeCell}
- Selected range: ${ctx.selectedRange}`;
    
    if (ctx.existingData && ctx.existingData.length > 0) {
      const dataPreview = ctx.existingData.slice(0, 30) // Limit to 30 cells
        .map(c => `${c.address}: ${c.formula || c.value}`)
        .join('\n');
      contextMessage += `\n- Existing cell data:\n${dataPreview}`;
    }
  }

  // Build messages array for the AI
  const messages: OpenAI.Chat.ChatCompletionMessageParam[] = [
    { role: 'system', content: SYSTEM_PROMPT },
    ...input.conversationHistory.map(m => ({
      role: m.role as 'user' | 'assistant',
      content: m.content
    })),
    {
      role: 'user',
      content: input.message + contextMessage
    }
  ];

  // ── Agentic Loop ──────────────────────────────────────────────────────────
  // The AI may call tools multiple times before giving a final answer.
  // Example: User asks "build a DCF for Apple" →
  //   Turn 1: AI calls get_stock_data('AAPL', 'all')
  //   Turn 2: AI calls get_macro_data('FEDFUNDS')  
  //   Turn 3: AI returns final JSON with all cell operations
  
  const MAX_ITERATIONS = 5; // Prevent infinite loops
  let iterations = 0;

  while (iterations < MAX_ITERATIONS) {
    iterations++;

    const completion = await client.chat.completions.create({
      model,
      messages,
      tools,
      tool_choice: 'auto',
      temperature: 0.1, // Low temperature = more deterministic financial calculations
      max_tokens: 4000,
    });

    const responseMessage = completion.choices[0].message;
    messages.push(responseMessage); // Add AI response to history

    // If AI wants to call tools, execute them and loop
    if (responseMessage.tool_calls && responseMessage.tool_calls.length > 0) {
      for (const toolCall of responseMessage.tool_calls) {
        if (toolCall.type === 'function' && toolCall.function) {
            const args = JSON.parse(toolCall.function.arguments);
            let toolResult = '';

            try {
                if (toolCall.function.name === 'get_stock_data') {
                    const data = await getStockData(args.ticker, args.dataType);
                    toolResult = JSON.stringify(data);
                } else if (toolCall.function.name === 'get_macro_data') {
                    const data = await getMacroData(args.seriesId);
                    toolResult = JSON.stringify(data);
                }
            } catch (e) {
                toolResult = JSON.stringify({ error: `Failed to fetch data: ${e}` });
            }

            // Add the tool result back into the conversation
            messages.push({
                role: 'tool',
                tool_call_id: toolCall.id,
                content: toolResult
            });
        } 
      }
      continue; // Loop again with tool results added
    }

    // No more tool calls — this is the final response
    const content = responseMessage.content || '';

    // Extract JSON from the response
    try {
      // The AI might wrap its JSON in markdown code blocks — strip them
      const jsonMatch = content.match(/```(?:json)?\s*([\s\S]*?)\s*```/) || 
                        content.match(/(\{[\s\S]*\})/);
      
      if (jsonMatch) {
        const parsed = JSON.parse(jsonMatch[1]);
        return {
          message: parsed.message || content,
          cellOperations: parsed.cellOperations || [],
          highlightRanges: parsed.highlightRanges || [],
          references: parsed.references || []
        };
      }
    } catch {
      // If JSON parsing fails, return the message as plain text (no cell operations)
      console.warn('AI returned non-JSON response, treating as plain text');
    }

    return {
      message: content,
      cellOperations: [],
      highlightRanges: [],
      references: []
    };
  }

  return {
    message: 'I hit a processing limit. Please try a simpler request.',
    cellOperations: [],
    highlightRanges: [],
    references: []
  };
}