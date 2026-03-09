/* global Office, Excel */

// ─── State ────────────────────────────────────────────────────────────────────
const BACKEND_URL = 'http://localhost:3001/api';
// During local dev, use: 'http://localhost:3001/api'

let conversationHistory = [];
let sendContextWithNext = false;

// ─── Office Initialization ────────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById('statusDot').style.color = '#00c851';
    document.getElementById('statusDot').title = 'Connected to Excel';
    
    document.getElementById('send-btn').addEventListener('click', handleSend);
    document.getElementById('ctx-btn').addEventListener('click', toggleContext);
    document.getElementById('user-input').addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        handleSend();
      }
    });
  }
});

// ─── Context Toggle ───────────────────────────────────────────────────────────
function toggleContext() {
  sendContextWithNext = !sendContextWithNext;
  const btn = document.getElementById('ctx-btn');
  btn.style.background = sendContextWithNext ? '#1565C0' : '';
  btn.title = sendContextWithNext 
    ? 'Sheet context ENABLED — will send selected cells' 
    : 'Click to include sheet data in next message';
}

// ─── Get Spreadsheet Context ─────────────────────────────────────────────────
// This reads the user's current selection from Excel via Office.js
async function getSpreadsheetContext() {
  return new Promise((resolve) => {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const selection = context.workbook.getSelectedRange();
      
      sheet.load('name');
      selection.load(['address', 'values', 'formulas', 'cellCount']);
      
      await context.sync();

      const existingData = [];
      if (selection.cellCount <= 200) { // Only read small selections
        // Parse the address to get the range bounds
        const values = selection.values;
        const formulas = selection.formulas;
        
        // Get address parts: "Sheet1!A1:C5" → "A1:C5"
        const rangePart = selection.address.includes('!') 
          ? selection.address.split('!')[1] 
          : selection.address;
        
        // Extract starting cell column/row
        const startMatch = rangePart.match(/([A-Z]+)(\d+)/);
        if (startMatch) {
          const startCol = startMatch[1].charCodeAt(0) - 65; // A=0, B=1, etc.
          const startRow = parseInt(startMatch[2]);
          
          for (let r = 0; r < values.length; r++) {
            for (let c = 0; c < values[r].length; c++) {
              const colLetter = String.fromCharCode(65 + startCol + c);
              const cellAddr = `${colLetter}${startRow + r}`;
              const formula = formulas[r][c];
              const value = values[r][c];
              
              if (value !== '' && value !== null) {
                existingData.push({
                  address: cellAddr,
                  value: typeof value === 'number' ? value : String(value),
                  formula: formula !== value ? String(formula) : undefined
                });
              }
            }
          }
        }
      }

      resolve({
        activeCell: selection.address.split(':')[0].replace(/.*!/, ''),
        selectedRange: selection.address.replace(/.*!/, ''),
        sheetName: sheet.name,
        existingData
      });
    }).catch(() => resolve(null));
  });
}

// ─── Send Message ─────────────────────────────────────────────────────────────
async function handleSend() {
  const input = document.getElementById('user-input');
  const message = input.value.trim();
  if (!message) return;

  input.value = '';
  appendMessage('user', message);

  // Optionally gather spreadsheet context
  let spreadsheetContext = null;
  if (sendContextWithNext) {
    spreadsheetContext = await getSpreadsheetContext();
    sendContextWithNext = false;
    document.getElementById('ctx-btn').style.background = '';
  }

  // Show loading
  document.getElementById('loading').style.display = 'flex';
  document.getElementById('send-btn').disabled = true;

  try {
    const response = await fetch(`${BACKEND_URL}/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message,
        conversationHistory,
        spreadsheetContext
      })
    });

    if (!response.ok) throw new Error(`Server error: ${response.status}`);

    const data = await response.json();

    // Add to conversation history for multi-turn context
    conversationHistory.push(
      { role: 'user', content: message },
      { role: 'assistant', content: data.message }
    );
    // Keep last 10 turns to avoid token overflow
    if (conversationHistory.length > 20) {
      conversationHistory = conversationHistory.slice(-20);
    }

    // Display the AI's message
    appendMessage('assistant', data.message);

    // If there are cell operations, write them to Excel
    if (data.cellOperations && data.cellOperations.length > 0) {
      await writeCellsToExcel(data.cellOperations, data.highlightRanges || []);
      showCellReferences(data.references || []);
    }

  } catch (error) {
    appendMessage('assistant', `⚠️ Error: ${error.message}. Is the backend server running?`);
  } finally {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('send-btn').disabled = false;
  }
}

// ─── Write Cells to Excel ─────────────────────────────────────────────────────
// This is the core Office.js function that builds the financial models in the spreadsheet.
async function writeCellsToExcel(cellOperations, highlightRanges) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // ── Step 1: Write all values and formulas ──────────────────────────────────
    for (const op of cellOperations) {
      const range = sheet.getRange(op.address);
      
      if (op.formula) {
        range.formulas = [[op.formula]];
      } else if (op.value !== undefined) {
        range.values = [[op.value]];
      }
      
      if (op.numberFormat) {
        range.numberFormat = [[op.numberFormat]];
      }
    }

    // ── Step 2: Apply highlight colors ────────────────────────────────────────
    // We use Excel's interior fill color to visually mark cell types.
    for (const hl of highlightRanges) {
      try {
        const range = sheet.getRange(hl.range);
        range.format.fill.color = hl.color;

        // For dark backgrounds (headers), make text white and bold
        if (hl.color === '#1565C0' || hl.color.startsWith('#0')) {
          range.format.font.color = '#FFFFFF';
          range.format.font.bold = true;
        }
      } catch (e) {
        console.warn(`Could not highlight range ${hl.range}:`, e);
      }
    }

    // ── Step 3: Auto-fit columns ───────────────────────────────────────────────
    // Make all columns wide enough to show their content
    sheet.getUsedRange().format.autofitColumns();

    await context.sync();

    // ── Step 4: Scroll to show the model ──────────────────────────────────────
    if (cellOperations.length > 0) {
      const firstCell = sheet.getRange(cellOperations[0].address);
      firstCell.select();
    }
  });
}

// ─── Show Cell References in Sidebar ─────────────────────────────────────────
function showCellReferences(references) {
  if (!references || references.length === 0) return;

  const panel = document.getElementById('cell-refs');
  const list = document.getElementById('refs-list');
  
  panel.style.display = 'block';
  list.innerHTML = references.map(ref => `
    <div class="ref-item" onclick="selectCell('${ref.address}')">
      <span class="ref-addr">${ref.address}</span>
      <span class="ref-desc">${ref.description}</span>
    </div>
  `).join('');
}

// ─── Navigate to Cell on Click ────────────────────────────────────────────────
function selectCell(address) {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange(address);
    range.select();
    await context.sync();
  });
}

// ─── Append Chat Message ──────────────────────────────────────────────────────
function appendMessage(role, content) {
  const chatWindow = document.getElementById('chat-window');
  const div = document.createElement('div');
  div.className = `message ${role}`;
  
  // Convert **bold** markdown to <strong> for readability
  const formatted = content
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/\n/g, '<br>');
  
  div.innerHTML = `<div class="msg-content">${formatted}</div>`;
  chatWindow.appendChild(div);
  chatWindow.scrollTop = chatWindow.scrollHeight;
}