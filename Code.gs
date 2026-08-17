/**
 * Groq Sheets Autocomplete — an open-source AI sentence-completion agent for Google Sheets.
 *
 * Everything runs in your own Apps Script project and talks directly to the Groq API
 * with your own key. No third-party server, no telemetry.
 *
 * Licence: MIT
 */

/* ------------------------------------------------------------------ *
 * Constants
 * ------------------------------------------------------------------ */

var GROQ_ENDPOINT = 'https://api.groq.com/openai/v1/chat/completions';

/**
 * Fallback list, used only until listModels() gets the live set from Groq.
 * Groq decommissions models on a rolling basis (llama-3.1-8b-instant and
 * llama-3.3-70b-versatile were both retired on 2026-08-16), so never trust a
 * hardcoded list — the sidebar refreshes this from /v1/models on every open.
 */
var MODELS = [
  { id: 'openai/gpt-oss-20b',          label: 'GPT-OSS 20B — fastest, good default' },
  { id: 'openai/gpt-oss-120b',         label: 'GPT-OSS 120B — smarter' },
  { id: 'qwen/qwen3.6-27b',            label: 'Qwen 3.6 27B' },
  { id: 'moonshotai/kimi-k2-instruct', label: 'Kimi K2 — long context' }
];

var MODELS_ENDPOINT = 'https://api.groq.com/openai/v1/models';

/** Substrings marking models that cannot serve chat completions. */
var NON_CHAT = /whisper|tts|orpheus|guard|embed|rerank/i;

var DEFAULTS = {
  model: 'openai/gpt-oss-20b',
  temperature: 0.4,
  maxTokens: 500,
  numSuggestions: 3,
  headerRow: 1,
  neighbourCount: 8,
  useHeader: true,
  useRow: true,
  useNeighbours: true,
  autoSuggest: false,
  instruction: '',
  projectContext: '',
  useSheetMap: true
};

/** UserProperties caps a single value at 9 KB — keep the brief well inside it. */
var MAX_CONTEXT_CHARS = 4000;

var PROP_KEY = 'GROQ_API_KEY';
var PROP_SETTINGS = 'GSA_SETTINGS';

/** Above this many selected cells we stop enumerating them on every poll. */
var SELECTION_CAP = 400;

/* ------------------------------------------------------------------ *
 * Menu / lifecycle
 * ------------------------------------------------------------------ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Autocomplete')
    .addItem('Open assistant', 'showSidebar')
    .addSeparator()
    .addItem('Set Groq API key…', 'promptForApiKey')
    .addItem('Fill selected range with AI', 'fillSelectedRange')
    .addSeparator()
    .addItem('Clear stored API key', 'clearApiKey')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('AI Autocomplete');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ------------------------------------------------------------------ *
 * Settings & key storage (per-user, never shared with collaborators)
 * ------------------------------------------------------------------ */

function props_() {
  return PropertiesService.getUserProperties();
}

function promptForApiKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt(
    'Groq API key',
    'Paste your key from https://console.groq.com/keys\n(stored privately in your own user properties):',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var key = (res.getResponseText() || '').trim();
  if (!key) {
    ui.alert('No key entered.');
    return;
  }
  props_().setProperty(PROP_KEY, key);
  ui.alert('Key saved. Open "AI Autocomplete → Open assistant" to start.');
}

function saveApiKey(key) {
  key = (key || '').trim();
  if (!key) throw new Error('Empty API key.');
  props_().setProperty(PROP_KEY, key);
  return true;
}

function clearApiKey() {
  props_().deleteProperty(PROP_KEY);
  SpreadsheetApp.getUi().alert('Groq API key removed from this account.');
}

function hasApiKey() {
  return !!props_().getProperty(PROP_KEY);
}

function getApiKey_() {
  var key = props_().getProperty(PROP_KEY);
  if (!key) throw new Error('No Groq API key set. Use "AI Autocomplete → Set Groq API key…".');
  return key;
}

function getSettings() {
  var raw = props_().getProperty(PROP_SETTINGS);
  var saved = {};
  if (raw) {
    try { saved = JSON.parse(raw); } catch (err) { saved = {}; }
  }
  var out = {};
  Object.keys(DEFAULTS).forEach(function (k) {
    out[k] = (saved[k] === undefined || saved[k] === null) ? DEFAULTS[k] : saved[k];
  });
  out.models = MODELS;
  out.hasKey = hasApiKey();
  return out;
}

function saveSettings(patch) {
  var current = getSettings();
  delete current.models;
  delete current.hasKey;
  Object.keys(patch || {}).forEach(function (k) {
    if (DEFAULTS.hasOwnProperty(k)) current[k] = patch[k];
  });
  if (current.projectContext) {
    current.projectContext = String(current.projectContext).slice(0, MAX_CONTEXT_CHARS);
  }
  props_().setProperty(PROP_SETTINGS, JSON.stringify(current));
  return getSettings();
}

/* ------------------------------------------------------------------ *
 * Context engine — what the model gets to see
 * ------------------------------------------------------------------ */

/**
 * Reads everything around the active cell that could inform a completion.
 * Called by the sidebar on a poll loop, so it stays cheap: at most
 * (2 * neighbourCount + 1) rows and one header row are touched.
 */
function contextFor_(sheet, row, col, settings) {
  var cell = sheet.getRange(row, col);
  var lastRow = Math.max(sheet.getLastRow(), row);
  var lastCol = Math.max(sheet.getLastColumn(), col);

  var ctx = {
    sheetName: sheet.getName(),
    a1: cell.getA1Notation(),
    row: row,
    col: col,
    text: String(cell.getDisplayValue() || ''),
    header: '',
    headers: [],
    rowPairs: [],
    neighbours: []
  };

  var headerRow = Math.max(1, Number(settings.headerRow) || 1);

  // Column header + the full header strip (used to label the row values).
  if (headerRow <= lastRow) {
    var headerVals = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
    ctx.headers = headerVals.map(function (v) { return String(v || ''); });
    ctx.header = ctx.headers[col - 1] || '';
  }

  // Whole row across columns, as header: value pairs, skipping the active cell.
  if (settings.useRow && row !== headerRow) {
    var rowVals = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
    for (var c = 0; c < rowVals.length; c++) {
      var val = String(rowVals[c] || '').trim();
      if (!val || (c + 1) === col) continue;
      ctx.rowPairs.push({
        header: (ctx.headers[c] || columnLetter_(c + 1)),
        value: truncate_(val, 300)
      });
    }
  }

  // Neighbouring rows in the same column — this is what teaches the model the pattern.
  if (settings.useNeighbours) {
    var span = Math.max(0, Number(settings.neighbourCount) || 0);
    if (span > 0) {
      var start = Math.max(headerRow + 1, row - span);
      var end = Math.min(lastRow, row + span);
      if (end >= start) {
        var colVals = sheet.getRange(start, col, end - start + 1, 1).getDisplayValues();
        for (var i = 0; i < colVals.length; i++) {
          var r = start + i;
          if (r === row) continue;
          var v = String(colVals[i][0] || '').trim();
          if (v) ctx.neighbours.push(truncate_(v, 300));
        }
      }
    }
  }

  return ctx;
}

/** Context for whatever cell the user is on, plus a summary of their selection. */
function getCellContext() {
  var settings = getSettings();
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  var ctx = contextFor_(sheet, cell.getRow(), cell.getColumn(), settings);
  ctx.selection = selectionInfo_(sheet);
  return ctx;
}

/**
 * Describes the selected block. Clamped to the used range so selecting a whole
 * column reports the ~200 rows that exist rather than a million empty ones.
 */
function selectionInfo_(sheet) {
  var range = sheet.getActiveRange();
  if (!range) return null;

  var usedRow = Math.max(1, sheet.getLastRow());
  var usedCol = Math.max(1, sheet.getLastColumn());
  var r0 = range.getRow();
  var c0 = range.getColumn();
  var r1 = Math.max(r0, Math.min(range.getLastRow(), usedRow));
  var c1 = Math.max(c0, Math.min(range.getLastColumn(), usedCol));
  var rows = r1 - r0 + 1;
  var cols = c1 - c0 + 1;

  var info = {
    a1: sheet.getRange(r0, c0, rows, cols).getA1Notation(),
    rows: rows,
    cols: cols,
    total: rows * cols,
    empty: 0,
    cells: [],
    truncated: false
  };
  if (info.total <= 1) return info;
  if (info.total > SELECTION_CAP) { info.truncated = true; return info; }

  var vals = sheet.getRange(r0, c0, rows, cols).getDisplayValues();
  for (var i = 0; i < rows; i++) {
    for (var j = 0; j < cols; j++) {
      var filled = String(vals[i][j] || '').trim().length > 0;
      if (!filled) info.empty++;
      info.cells.push({ a1: columnLetter_(c0 + j) + (r0 + i), filled: filled });
    }
  }
  return info;
}

/**
 * A compact map of the whole workbook: every tab and its column headers. Gives
 * the model the vocabulary of your sheet without shipping any row data.
 * Cached for 5 minutes because it is rebuilt on every suggestion.
 */
function sheetMap_() {
  var cache = CacheService.getUserCache();
  var hit = cache ? cache.get('GSA_SHEETMAP') : null;
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }

  var settings = getSettings();
  var headerRow = Math.max(1, Number(settings.headerRow) || 1);
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var map = [];

  for (var i = 0; i < sheets.length && i < 10; i++) {
    var sh = sheets[i];
    var lastCol = Math.min(sh.getLastColumn(), 30);
    if (lastCol < 1 || sh.getLastRow() < headerRow) continue;

    var vals = sh.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
    var headers = [];
    for (var c = 0; c < vals.length; c++) {
      var h = String(vals[c] || '').trim();
      if (h) headers.push(h);
    }
    if (headers.length) {
      map.push({ name: sh.getName(), rows: Math.max(0, sh.getLastRow() - headerRow), headers: headers });
    }
  }

  if (cache) cache.put('GSA_SHEETMAP', JSON.stringify(map), 300);
  return map;
}

function columnLetter_(n) {
  var s = '';
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - m - 1) / 26;
  }
  return s;
}

function truncate_(s, n) {
  s = String(s);
  return s.length > n ? s.slice(0, n) + '…' : s;
}

/* ------------------------------------------------------------------ *
 * Prompt construction
 * ------------------------------------------------------------------ */

function buildMessages_(ctx, settings) {
  var partial = String(ctx.text || '');
  var isContinuation = partial.trim().length > 0;
  var n = Math.min(5, Math.max(1, Number(settings.numSuggestions) || 3));

  var sys = [
    'You are an autocomplete engine embedded in a Google Sheets cell.',
    'You return short, plausible cell content — never commentary, never markdown, never quotes around the value.',
    'Match the tone, length, capitalisation and formatting of the existing entries in the same column exactly.',
    isContinuation
      ? 'The user has already typed part of the cell. Return ONLY the text that continues it, starting exactly where they stopped (include a leading space if the continuation begins a new word). Never repeat the text they already typed.'
      : 'The cell is empty. Return the complete value for the cell.',
    'Respond with strict JSON only, in this shape: {"completions": [' +
      new Array(n + 1).join('"…", ').replace(/, $/, '') + ']}',
    'Give exactly ' + n + ' distinct alternative' + (n === 1 ? '' : 's') + ', best first.'
  ].join(' ');

  var parts = [];

  // Background first: everything after it is read in light of this.
  var brief = String(settings.projectContext || '').trim();
  if (brief) {
    parts.push('ABOUT THIS PRODUCT AND SHEET (background — treat as fact):\n' + brief);
  }

  if (settings.useSheetMap && ctx.sheetMap && ctx.sheetMap.length) {
    parts.push(
      'WORKBOOK STRUCTURE:\n' +
      ctx.sheetMap.map(function (t) {
        return '- ' + t.name + ' (' + t.rows + ' rows): ' + t.headers.join(' | ');
      }).join('\n')
    );
  }

  parts.push('SHEET: ' + ctx.sheetName + '   CELL: ' + ctx.a1);

  if (settings.useHeader && ctx.header) {
    parts.push('COLUMN HEADER: ' + ctx.header);
  }

  if (settings.useRow && ctx.rowPairs && ctx.rowPairs.length) {
    parts.push(
      'OTHER VALUES IN THIS ROW (facts you may use):\n' +
      ctx.rowPairs.map(function (p) { return '- ' + p.header + ': ' + p.value; }).join('\n')
    );
  }

  if (settings.useNeighbours && ctx.neighbours && ctx.neighbours.length) {
    parts.push(
      'EXISTING ENTRIES IN THIS COLUMN (copy their style and length):\n' +
      ctx.neighbours.map(function (v) { return '- ' + v; }).join('\n')
    );
  }

  var instruction = String(settings.instruction || '').trim();
  if (instruction) {
    parts.push('USER STYLE INSTRUCTIONS (highest priority):\n' + instruction);
  }

  parts.push(
    isContinuation
      ? 'TEXT ALREADY TYPED IN THE CELL:\n"""' + partial + '"""\n\nContinue it.'
      : 'The cell is empty. Write its value.'
  );

  return [
    { role: 'system', content: sys },
    { role: 'user', content: parts.join('\n\n') }
  ];
}

/* ------------------------------------------------------------------ *
 * Groq call
 * ------------------------------------------------------------------ */

/**
 * Budget for the reply. Reasoning models (gpt-oss) spend part of the completion
 * allowance on hidden reasoning tokens before emitting a single visible
 * character, so a tight cap truncates the JSON mid-string and Groq rejects it
 * with a 400. Scale the ceiling with how many suggestions we asked for.
 */
function tokenBudget_(settings) {
  var base = Number(settings.maxTokens) || DEFAULTS.maxTokens;
  var n = Math.min(5, Math.max(1, Number(settings.numSuggestions) || 3));
  return Math.min(4000, base + (n - 1) * 150);
}

function isReasoningModel_(model) {
  return /gpt-oss|qwen3|reason/i.test(String(model || ''));
}

function postGroq_(body) {
  var res = UrlFetchApp.fetch(GROQ_ENDPOINT, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + getApiKey_() },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  return { code: res.getResponseCode(), text: res.getContentText() };
}

/** Shared error mapping for both the JSON and plain-text calls. */
function checkGroq_(r, model) {
  if (r.code === 401) throw new Error('Groq rejected the API key (401). Re-enter it from the menu.');
  if (r.code === 429) throw new Error('Groq rate limit hit (429). Wait a moment, or pick a model with a higher cap.');
  if (r.code === 404) {
    throw new Error('Model "' + model + '" is not available — Groq has most likely ' +
      'decommissioned it. Open Settings and click "Refresh models", then pick a current one.');
  }
  if (r.code < 200 || r.code >= 300) {
    var msg = r.text;
    try { msg = JSON.parse(r.text).error.message; } catch (e) {}
    throw new Error('Groq error ' + r.code + ': ' + truncate_(msg, 300));
  }
}

/** One-shot prose call — no JSON contract. Used to draft the project brief. */
function callGroqText_(messages, settings, maxTokens) {
  var model = settings.model || DEFAULTS.model;
  var body = {
    model: model,
    messages: messages,
    temperature: 0.3,
    max_tokens: maxTokens || 700,
    stream: false
  };
  if (isReasoningModel_(model)) body.reasoning_effort = 'low';

  var r = postGroq_(body);
  checkGroq_(r, model);
  var payload = JSON.parse(r.text);
  var choice = payload && payload.choices && payload.choices[0];
  return String((choice && choice.message && choice.message.content) || '').trim();
}

function callGroq_(messages, settings) {
  var model = settings.model || DEFAULTS.model;
  var body = {
    model: model,
    messages: messages,
    temperature: Number(settings.temperature),
    max_tokens: tokenBudget_(settings),
    response_format: { type: 'json_object' },
    stream: false
  };
  if (isNaN(body.temperature)) body.temperature = DEFAULTS.temperature;

  // Keep hidden reasoning short so the visible JSON fits inside the budget.
  if (isReasoningModel_(model)) body.reasoning_effort = 'low';

  var r = postGroq_(body);

  // Groq validates JSON mode server-side. A model that rambles, emits reasoning
  // prose, or gets truncated fails that check — retry once in plain-text mode,
  // where parseCompletions_ is tolerant enough to cope.
  if (r.code === 400 && /validate JSON|json_validate|failed_generation/i.test(r.text)) {
    delete body.response_format;
    body.max_tokens = Math.min(4000, body.max_tokens * 2);
    body.messages = messages.slice();
    body.messages[0] = {
      role: 'system',
      content: messages[0].content +
        ' Output the raw JSON object and nothing else: no preamble, no explanation, no code fences.'
    };
    r = postGroq_(body);
  }

  checkGroq_(r, model);

  var payload = JSON.parse(r.text);
  var choice = payload && payload.choices && payload.choices[0];
  var content = choice && choice.message && choice.message.content;

  // A reasoning model can burn the whole budget and return nothing visible.
  if (!content && choice && choice.finish_reason === 'length') {
    throw new Error('The model hit its token limit before answering. Raise "Response length" ' +
      'in Settings, or ask for fewer suggestions.');
  }

  return String(content || '');
}

/**
 * Tolerant parser: the JSON contract is enforced via response_format, but small
 * models occasionally wrap it in a fence or return a bare array/string.
 */
function parseCompletions_(raw) {
  var s = String(raw || '').trim();
  s = s.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();

  var out = [];
  try {
    var parsed = JSON.parse(s);
    if (Array.isArray(parsed)) {
      out = parsed;
    } else if (parsed && Array.isArray(parsed.completions)) {
      out = parsed.completions;
    } else if (parsed && typeof parsed === 'object') {
      // e.g. {"completion": "..."} or {"1": "...", "2": "..."}
      Object.keys(parsed).forEach(function (k) {
        var v = parsed[k];
        if (typeof v === 'string') out.push(v);
        else if (Array.isArray(v)) out = out.concat(v);
      });
    } else if (typeof parsed === 'string') {
      out = [parsed];
    }
  } catch (err) {
    out = salvage_(s);
  }

  var seen = {};
  return out
    .map(function (v) {
      return String(v === null || v === undefined ? '' : v)
        .replace(/^\s*[-*\d.)\]]+\s+/, '')
        .replace(/^["'`]|["'`]$/g, '')
        .replace(/\s+$/, '');
    })
    .filter(function (v) {
      if (!v.trim()) return false;
      var k = v.trim().toLowerCase();
      if (seen[k]) return false;
      seen[k] = true;
      return true;
    });
}

/**
 * Asks Groq which models this key can actually use, so a decommissioned ID
 * can never strand the add-on. Falls back to MODELS if the call fails.
 */
function listModels() {
  try {
    var res = UrlFetchApp.fetch(MODELS_ENDPOINT, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + getApiKey_() },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return MODELS;

    var data = JSON.parse(res.getContentText()).data || [];
    var out = [];
    for (var i = 0; i < data.length; i++) {
      var id = String(data[i].id || '');
      if (!id || NON_CHAT.test(id)) continue;
      var ctx = Number(data[i].context_window || 0);
      out.push({ id: id, label: id + (ctx ? '  ·  ' + Math.round(ctx / 1000) + 'k ctx' : '') });
    }
    out.sort(function (a, b) { return a.id < b.id ? -1 : (a.id > b.id ? 1 : 0); });
    return out.length ? out : MODELS;
  } catch (err) {
    return MODELS;
  }
}

/**
 * Last resort when JSON.parse fails — usually because the reply was cut off
 * mid-array. Pulls the completed string literals out and discards the rest,
 * so a truncated response still yields the suggestions that did arrive.
 */
function salvage_(s) {
  // Only mine for string literals if this actually looks like broken JSON.
  // A plain-text reply is better handled line by line.
  if (!/^[\[{]/.test(s.trim())) return s.split('\n');

  var body = s;
  var at = s.indexOf('"completions"');
  if (at > -1) {
    var open = s.indexOf('[', at);
    if (open > -1) body = s.slice(open + 1);
  }

  var strings = body.match(/"(?:[^"\\]|\\.)*"/g);
  if (strings && strings.length) {
    return strings.map(function (raw) {
      try { return JSON.parse(raw); } catch (e) { return raw.replace(/^"|"$/g, ''); }
    });
  }
  return s.split('\n');
}

/* ------------------------------------------------------------------ *
 * Sidebar entry points
 * ------------------------------------------------------------------ */

function mergeSettings_(overrides) {
  var settings = getSettings();
  Object.keys(overrides || {}).forEach(function (k) {
    if (DEFAULTS.hasOwnProperty(k)) settings[k] = overrides[k];
  });
  return settings;
}

/** Removes an accidental restatement of what the user already typed. */
function stripEcho_(completions, prefix) {
  var p = String(prefix || '').trim().toLowerCase();
  if (!p) return completions;
  return completions.map(function (c) {
    return c.toLowerCase().indexOf(p) === 0 ? c.slice(p.length) : c;
  }).filter(function (c) { return c.trim().length > 0; });
}

/**
 * Main call from the sidebar. Returns suggestions plus the context echo so the
 * UI can show what the model actually saw.
 */
function suggest(overrides) {
  var settings = mergeSettings_(overrides);
  var ctx = getCellContext();
  if (settings.useSheetMap) ctx.sheetMap = sheetMap_();
  var messages = buildMessages_(ctx, settings);
  var raw = callGroq_(messages, settings);
  var completions = parseCompletions_(raw);

  var isContinuation = String(ctx.text || '').trim().length > 0;
  if (isContinuation) completions = stripEcho_(completions, ctx.text);

  return {
    sheetName: ctx.sheetName,
    a1: ctx.a1,
    prefix: ctx.text,
    mode: isContinuation ? 'continue' : 'fill',
    completions: completions.slice(0, Math.max(1, Number(settings.numSuggestions) || 3)),
    contextUsed: {
      header: settings.useHeader ? ctx.header : '',
      rowFields: settings.useRow ? ctx.rowPairs.length : 0,
      examples: settings.useNeighbours ? ctx.neighbours.length : 0
    }
  };
}

/**
 * Writes a suggestion into the target cell.
 * @param {string} a1 Target cell, guards against the user having moved on.
 * @param {string} completion The chosen text.
 * @param {string} mode 'continue' appends to what is there, 'fill' replaces.
 * @param {string=} sheetName Guards against the user having switched sheets.
 */
function applyCompletion(a1, completion, mode, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sheetName ? ss.getSheetByName(sheetName) : SpreadsheetApp.getActiveSheet();
  if (!sheet) throw new Error('Sheet "' + sheetName + '" no longer exists.');
  var range = sheet.getRange(a1);
  var existing = String(range.getDisplayValue() || '');
  var value;

  if (mode === 'continue' && existing.trim()) {
    var startsWithPunct = /^[,.;:!?')\]]/.test(completion);
    var needsSpace = !/\s$/.test(existing) && !/^\s/.test(completion) && !startsWithPunct;
    value = existing + (needsSpace ? ' ' : '') + completion;
  } else {
    value = completion;
  }

  range.setValue(value);

  // Re-assert the selection only when the cursor is already here — otherwise a
  // batch run would drag the user's cursor across the sheet cell by cell.
  if (sheet.getSheetId() === ss.getActiveSheet().getSheetId() &&
      sheet.getActiveCell().getA1Notation() === range.getA1Notation()) {
    sheet.setActiveRange(range);
  }
  return value;
}

/**
 * Completes ONE named cell and writes the result. The sidebar calls this in a
 * loop, one cell per request, so the run shows live progress, can be stopped
 * mid-way, and never trips Apps Script's 6-minute execution ceiling the way a
 * server-side loop over a large range would.
 *
 * @param {string} sheetName Sheet the cell lives on.
 * @param {string} a1 The cell.
 * @param {Object} overrides Settings from the sidebar.
 * @param {boolean} skipFilled Leave cells that already have content alone.
 */
function completeCell(sheetName, a1, overrides, skipFilled) {
  var settings = mergeSettings_(overrides);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sheetName ? ss.getSheetByName(sheetName) : SpreadsheetApp.getActiveSheet();
  if (!sheet) throw new Error('Sheet "' + sheetName + '" no longer exists.');

  var range = sheet.getRange(a1);
  var ctx = contextFor_(sheet, range.getRow(), range.getColumn(), settings);
  if (settings.useSheetMap) ctx.sheetMap = sheetMap_();
  var hasText = String(ctx.text || '').trim().length > 0;

  if (skipFilled && hasText) {
    return { a1: a1, status: 'skipped', reason: 'already filled' };
  }

  var completions = parseCompletions_(callGroq_(buildMessages_(ctx, settings), settings));
  if (hasText) completions = stripEcho_(completions, ctx.text);
  var best = completions[0];
  if (!best) return { a1: a1, status: 'empty', reason: 'model returned nothing' };

  var value = applyCompletion(a1, best, hasText ? 'continue' : 'fill', sheet.getName());
  return { a1: a1, status: 'written', value: truncate_(value, 120) };
}

/**
 * Reads the workbook structure plus a handful of real rows and asks the model
 * to write a short brief describing what this sheet is for. The user edits and
 * saves it — it is a starting point, not an authority.
 */
function draftProjectContext() {
  var settings = getSettings();
  var sheet = SpreadsheetApp.getActiveSheet();
  var map = sheetMap_();
  var headerRow = Math.max(1, Number(settings.headerRow) || 1);

  var sample = '';
  var lastRow = sheet.getLastRow();
  var lastCol = Math.min(sheet.getLastColumn(), 12);
  if (lastRow > headerRow && lastCol > 0) {
    var n = Math.min(6, lastRow - headerRow);
    var rows = sheet.getRange(headerRow + 1, 1, n, lastCol).getDisplayValues();
    var heads = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
    sample = rows.map(function (row, i) {
      var pairs = [];
      for (var c = 0; c < row.length; c++) {
        var v = String(row[c] || '').trim();
        if (v) pairs.push((heads[c] || columnLetter_(c + 1)) + ': ' + truncate_(v, 160));
      }
      return 'Row ' + (i + 1) + ' — ' + pairs.join(' // ');
    }).join('\n');
  }

  var messages = [
    {
      role: 'system',
      content: 'You write short factual briefs that will be given to another AI as background ' +
        'context. Describe only what the data shows. Never invent a company name, product name, ' +
        'team, methodology or purpose that is not evident in the data. If something is unclear, ' +
        'say so plainly rather than guessing. Output 5-9 short lines, no headings, no markdown, ' +
        'no preamble.'
    },
    {
      role: 'user',
      content:
        'Describe what this spreadsheet is for, so an autocomplete assistant can write new cell ' +
        'values that fit it. Cover: what the workbook tracks, what one row represents, what each ' +
        'important column holds, the product or domain vocabulary in use, and the writing style ' +
        'of the existing entries.\n\n' +
        'WORKBOOK STRUCTURE:\n' +
        map.map(function (t) { return '- ' + t.name + ' (' + t.rows + ' rows): ' + t.headers.join(' | '); }).join('\n') +
        '\n\nSAMPLE ROWS FROM "' + sheet.getName() + '":\n' + (sample || '(no data rows yet)')
    }
  ];

  return callGroqText_(messages, settings, 700);
}

/* ------------------------------------------------------------------ *
 * Batch fill for a selected range (menu item)
 * ------------------------------------------------------------------ */

function fillSelectedRange() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var cells = range.getNumRows() * range.getNumColumns();

  if (cells > 60) {
    ui.alert('Select 60 cells or fewer — Apps Script caps a single run at ~6 minutes.');
    return;
  }

  var confirm = ui.alert(
    'Fill ' + cells + ' cell' + (cells === 1 ? '' : 's') + ' with AI?',
    'Each cell uses one Groq request. Empty cells get a full value; partly typed cells get continued.',
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  var startRow = range.getRow();
  var startCol = range.getColumn();
  var filled = 0;
  var failed = 0;

  for (var r = 0; r < range.getNumRows(); r++) {
    for (var c = 0; c < range.getNumColumns(); c++) {
      var a1 = columnLetter_(startCol + c) + (startRow + r);
      try {
        if (completeCell(sheet.getName(), a1, {}, false).status === 'written') filled++;
      } catch (err) {
        failed++;
        if (String(err.message).indexOf('429') > -1) Utilities.sleep(2000);
      }
      Utilities.sleep(150); // stay under 30 req/min on the free tier
    }
  }
  ui.alert('Done. Filled ' + filled + ' cell' + (filled === 1 ? '' : 's') +
    (failed ? ', ' + failed + ' failed.' : '.'));
}

/* ------------------------------------------------------------------ *
 * Connection test (sidebar settings panel)
 * ------------------------------------------------------------------ */

function testConnection(model) {
  var settings = getSettings();
  if (model) settings.model = model;
  var raw = callGroq_([
    { role: 'system', content: 'Reply with strict JSON: {"completions":["ok"]}' },
    { role: 'user', content: 'ping' }
  ], settings);
  var parsed = parseCompletions_(raw);
  return 'Connected to ' + settings.model + ' — replied "' + (parsed[0] || raw).slice(0, 40) + '"';
}
