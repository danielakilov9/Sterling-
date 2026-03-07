import { Router, Request, Response } from 'express';
import multer from 'multer';
import pdfParse from 'pdf-parse';
import { runAIAgent } from '../services/ai';

const router = Router();
const upload = multer({ storage: multer.memoryStorage() });

// ─── Types ────────────────────────────────────────────────────────────────────

// What the Excel add-in sends us
interface ChatRequest {
  message: string;               // User's natural language request
  conversationHistory: Message[]; // Previous turns (for context)
  spreadsheetContext?: {         // Current state of the active sheet
    activeCell: string;          // e.g. "B5"
    selectedRange: string;       // e.g. "A1:D20"
    sheetName: string;
    existingData?: CellData[];   // Cells that already have data
  };
}

interface Message {
  role: 'user' | 'assistant';
  content: string;
}

interface CellData {
  address: string;  // e.g. "A1"
  value: string | number | null;
  formula?: string; // e.g. "=B2*C2"
}

// What we send back to the Excel add-in
export interface ChatResponse {
  message: string;          // AI's explanation in plain English
  cellOperations?: CellOperation[]; // Cells to write (if any)
  highlightRanges?: HighlightRange[]; // Ranges to color
  references?: CellReference[];      // Cells mentioned in explanation
}

export interface CellOperation {
  address: string;           // e.g. "C5"
  value?: string | number;   // Static value
  formula?: string;          // Excel formula, e.g. "=B2-B3"
  numberFormat?: string;     // e.g. "$#,##0.00" or "0.0%"
}

export interface HighlightRange {
  range: string;    // e.g. "A1:D1"
  color: string;    // Hex color, e.g. "#E8F5E9"
  label?: string;   // Tooltip label
}

export interface CellReference {
  address: string;
  description: string; // e.g. "Revenue (Year 1)"
}

// ─── POST /api/chat ───────────────────────────────────────────────────────────
router.post('/', async (req: Request, res: Response) => {
  try {
    const body: ChatRequest = req.body;

    if (!body.message) {
      return res.status(400).json({ error: 'message is required' });
    }

    const result = await runAIAgent({
      message: body.message,
      conversationHistory: body.conversationHistory || [],
      spreadsheetContext: body.spreadsheetContext
    });

    return res.json(result);
  } catch (error) {
    console.error('Chat error:', error);
    return res.status(500).json({ error: 'Internal server error' });
  }
});

// ─── POST /api/chat/upload-pdf ────────────────────────────────────────────────
// Accepts a 10-K PDF and extracts its text for the AI to analyze
router.post('/upload-pdf', upload.single('pdf'), async (req: Request, res: Response) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No PDF uploaded' });
    }

    const pdfData = await pdfParse(req.file.buffer);
    // Truncate to first 15,000 characters to stay within context limits
    const extractedText = pdfData.text.slice(0, 15000);

    return res.json({
      success: true,
      pageCount: pdfData.numpages,
      text: extractedText
    });
  } catch (error) {
    console.error('PDF parse error:', error);
    return res.status(500).json({ error: 'Failed to parse PDF' });
  }
});

export { router as chatRouter };