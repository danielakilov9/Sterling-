/* global Office, Excel */

// ─── Config ───────────────────────────────────────────────────────────────────
const BACKEND_URL = 'https://localhost:3001/api';

// ─── State ────────────────────────────────────────────────────────────────────
let conversationHistory  = [];
let sendContextWithNext  = false;
let activeModel          = 'dcf';

// ─── Office Init ──────────────────────────────────────────────────────────────
// Guard flag prevents renderWelcome() and bindEvents() from running more than
// once if Office.onReady fires multiple times (known behavior in some versions
// of the Office.js loader when the add-in reloads inside the task pane).
let _initialized = false;

Office.onReady((info) => {
  if (_initialized) return;   // ← exit immediately on any duplicate call
  _initialized = true;

  if (info.host === Office.HostType.Excel) {
    setStatus('connected', 'ready');
    renderWelcome();
    bindEvents();
  } else {
    setStatus('error', 'not excel');
  }
});

// ─── Status helper ────────────────────────────────────────────────────────────
function setStatus(state, label) {
  const dot   = document.getElementById('statusDot');
  const lbl   = document.getElementById('statusLabel');
  dot.className = 'status-dot ' + state;
  lbl.textContent = label;
}

// ─── Bind all events ──────────────────────────────────────────────────────────
// Remove then re-add every listener so hot-reloads never stack duplicates.
// cloneNode(true) replaces the element with a fresh copy that has no listeners.
function bindEvents() {
  ['send-btn', 'ctx-btn', 'refsClose'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.replaceWith(el.cloneNode(true)); // wipe any existing listeners
  });
  document.getElementById('send-btn').addEventListener('click', handleSend);
  document.getElementById('ctx-btn').addEventListener('click', toggleContext);
  document.getElementById('refsClose').addEventListener('click', () => {
    document.getElementById('cell-refs').style.display = 'none';
  });

  // Auto-resize textarea + Enter to send
  const ta = document.getElementById('user-input');
  ta.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  });
  ta.addEventListener('input', () => {
    ta.style.height = 'auto';
    ta.style.height = Math.min(ta.scrollHeight, 100) + 'px';
  });

  // Model chips
  document.querySelectorAll('.model-chip').forEach((chip) => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('.model-chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      activeModel = chip.dataset.model;
    });
  });
}

// ─── Welcome message ──────────────────────────────────────────────────────────
// Guard against hot-reload double-render: check the DOM directly instead of a
// JS variable. On webpack hot-reload the script re-runs but the DOM persists,
// so a JS flag resets to false while the welcome card is already in the DOM.
function renderWelcome() {
  const chatWindow = document.getElementById('chat-window');
  if (chatWindow.querySelector('.welcome-msg')) return; // already rendered
  const div = document.createElement('div');
  div.className = 'message assistant';
  div.innerHTML = `
    <div class="msg-role">Sterling</div>
    <div class="msg-content welcome-msg">
      <div class="welcome-title">// READY</div>
      <div class="welcome-item">
        <span class="tag">DCF</span>
        <span>Discounted cash flow with sensitivity tables</span>
      </div>
      <div class="welcome-item">
        <span class="tag">3-STMT</span>
        <span>Linked income, balance sheet &amp; cash flow</span>
      </div>
      <div class="welcome-item">
        <span class="tag">LBO</span>
        <span>Leveraged buyout with IRR &amp; MOIC returns</span>
      </div>
      <div class="welcome-item">
        <span class="tag">LIVE</span>
        <span>Pulls real-time data from Yahoo Finance &amp; SEC</span>
      </div>
      <div class="welcome-hint">
        Try: <em>"Build a DCF for Apple"</em><br>
        or:  <em>"Create a 3-statement model with $500M revenue"</em>
      </div>
    </div>`;
  chatWindow.appendChild(div);
}

// ─── Context toggle ───────────────────────────────────────────────────────────
function toggleContext() {
  sendContextWithNext = !sendContextWithNext;
  const btn    = document.getElementById('ctx-btn');
  const status = document.getElementById('ctx-status');
  if (sendContextWithNext) {
    btn.classList.add('active');
    status.textContent = '⊞ context ON';
    status.classList.add('active');
  } else {
    btn.classList.remove('active');
    status.textContent = '⊡ no context';
    status.classList.remove('active');
  }
}

// ─── Read spreadsheet context ─────────────────────────────────────────────────
async function getSpreadsheetContext() {
  return new Promise((resolve) => {
    Excel.run(async (context) => {
      const sheet     = context.workbook.worksheets.getActiveWorksheet();
      const selection = context.workbook.getSelectedRange();
      sheet.load('name');
      selection.load(['address', 'values', 'formulas', 'cellCount']);
      await context.sync();

      const existingData = [];
      if (selection.cellCount <= 200) {
        const values   = selection.values;
        const formulas = selection.formulas;
        const rangePart = selection.address.includes('!')
          ? selection.address.split('!')[1]
          : selection.address;
        const startMatch = rangePart.match(/([A-Z]+)(\d+)/);
        if (startMatch) {
          const startCol = startMatch[1].charCodeAt(0) - 65;
          const startRow = parseInt(startMatch[2]);
          for (let r = 0; r < values.length; r++) {
            for (let c = 0; c < values[r].length; c++) {
              const addr    = `${String.fromCharCode(65 + startCol + c)}${startRow + r}`;
              const formula = formulas[r][c];
              const value   = values[r][c];
              if (value !== '' && value !== null) {
                existingData.push({
                  address: addr,
                  value: typeof value === 'number' ? value : String(value),
                  formula: formula !== value ? String(formula) : undefined,
                });
              }
            }
          }
        }
      }

      resolve({
        activeCell:    selection.address.split(':')[0].replace(/.*!/, ''),
        selectedRange: selection.address.replace(/.*!/, ''),
        sheetName:     sheet.name,
        existingData,
      });
    }).catch(() => resolve(null));
  });
}

// ─── Send message ─────────────────────────────────────────────────────────────
async function handleSend() {
  const ta      = document.getElementById('user-input');
  const message = ta.value.trim();
  if (!message) return;

  ta.value = '';
  ta.style.height = 'auto';
  appendMessage('user', message);

  // Gather context if toggled
  let spreadsheetContext = null;
  if (sendContextWithNext) {
    spreadsheetContext = await getSpreadsheetContext();
    // Reset context toggle after use
    sendContextWithNext = false;
    document.getElementById('ctx-btn').classList.remove('active');
    document.getElementById('ctx-status').textContent = '⊡ no context';
    document.getElementById('ctx-status').classList.remove('active');
  }

  // Show typing indicator
  setTyping(true);
  setStatus('connected', 'thinking');
  document.getElementById('send-btn').disabled = true;

  try {
    const response = await fetch(`${BACKEND_URL}/chat`, {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message,
        conversationHistory,
        spreadsheetContext,
        preferredModel: activeModel,
      }),
    });

    if (!response.ok) throw new Error(`Server error ${response.status}`);

    const data = await response.json();

    // Update conversation history (keep last 20 turns)
    conversationHistory.push(
      { role: 'user',      content: message },
      { role: 'assistant', content: data.message }
    );
    if (conversationHistory.length > 20) {
      conversationHistory = conversationHistory.slice(-20);
    }

    appendMessage('assistant', data.message, data.cellOperations?.length);

    if (data.cellOperations && data.cellOperations.length > 0) {
      await writeCellsToExcel(data.cellOperations, data.highlightRanges || []);
      showCellReferences(data.references || []);
    }

    setStatus('connected', 'ready');
  } catch (err) {
    appendMessage('assistant', `⚠️ Error: ${err.message}. Is the backend server running?`, 0, true);
    setStatus('error', 'offline');
  } finally {
    setTyping(false);
    document.getElementById('send-btn').disabled = false;
    document.getElementById('user-input').focus();
  }
}

// ─── Write cells to Excel ─────────────────────────────────────────────────────
async function writeCellsToExcel(cellOperations, highlightRanges) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    for (const op of cellOperations) {
      const range = sheet.getRange(op.address);
      if (op.formula)          range.formulas     = [[op.formula]];
      else if (op.value !== undefined) range.values = [[op.value]];
      if (op.numberFormat)     range.numberFormat = [[op.numberFormat]];
    }

    for (const hl of highlightRanges) {
      try {
        const range = sheet.getRange(hl.range);
        range.format.fill.color = hl.color;
        if (hl.color === '#1565C0' || hl.color.toLowerCase().startsWith('#0d') || hl.color.toLowerCase().startsWith('#1a')) {
          range.format.font.color = '#FFFFFF';
          range.format.font.bold  = true;
        }
      } catch (e) {
        console.warn(`Could not highlight ${hl.range}:`, e);
      }
    }

    sheet.getUsedRange().format.autofitColumns();
    await context.sync();

    // Scroll to first written cell
    if (cellOperations.length > 0) {
      sheet.getRange(cellOperations[0].address).select();
    }
  });
}

// ─── Cell reference panel ─────────────────────────────────────────────────────
function showCellReferences(references) {
  if (!references || references.length === 0) return;
  const panel = document.getElementById('cell-refs');
  const list  = document.getElementById('refs-list');
  panel.style.display = 'block';
  list.innerHTML = references.map(ref => `
    <div class="ref-item" onclick="selectCell('${ref.address}')">
      <span class="ref-addr">${ref.address}</span>
      <span class="ref-desc">${ref.description}</span>
    </div>
  `).join('');
}

function selectCell(address) {
  Excel.run(async (ctx) => {
    ctx.workbook.worksheets.getActiveWorksheet().getRange(address).select();
    await ctx.sync();
  });
}

// ─── Append chat message ──────────────────────────────────────────────────────
function appendMessage(role, content, cellCount, isError) {
  const chatWindow = document.getElementById('chat-window');
  const div = document.createElement('div');
  div.className = 'message ' + role + (isError ? ' error' : '');

  const roleLabel = role === 'user' ? 'YOU' : 'Sterling';

  // Convert **bold**, `code`, and newlines
  const formatted = content
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\n/g, '<br>');

  let badge = '';
  if (cellCount && cellCount > 0) {
    badge = `<div class="cells-badge">◈ ${cellCount} cell${cellCount > 1 ? 's' : ''} written to sheet</div>`;
  }

  div.innerHTML = `
    <div class="msg-role">${roleLabel}</div>
    <div class="msg-content">${formatted}${badge}</div>`;

  chatWindow.appendChild(div);
  chatWindow.scrollTop = chatWindow.scrollHeight;
}

// ─── Typing indicator ─────────────────────────────────────────────────────────
function setTyping(show) {
  document.getElementById('typing-indicator').style.display = show ? 'block' : 'none';
  if (show) {
    const chatWindow = document.getElementById('chat-window');
    chatWindow.scrollTop = chatWindow.scrollHeight;
  }
}