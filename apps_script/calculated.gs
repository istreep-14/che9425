/**
 * Calculated sheet pipeline: classify mode/variant, compute per-mode running ratings for a user,
 * and enrich with opponent stats using Chess.com public API.
 */

const CALC_SETTINGS = {
  userPropertyKey: 'PRIMARY_USERNAME',
  cacheMinutes: 60
};

function setPrimaryUsername(username) {
  let name = String(username || '').trim();
  if (!name) {
    try {
      const ss = getSpreadsheet_();
      const cfgName = (typeof SHEET_NAMES === 'object' && SHEET_NAMES.CONFIG) ? SHEET_NAMES.CONFIG : 'Config';
      const configSheet = ss.getSheetByName(cfgName);
      if (configSheet && typeof readConfig === 'function') {
        const cfg = readConfig(configSheet);
        const fromConfig = String((cfg && cfg.username) || '').trim();
        if (fromConfig) name = fromConfig;
      }
    } catch (e) {}
  }
  if (!name) {
    try {
      const email = Session.getActiveUser && Session.getActiveUser();
      const addr = email && typeof email.getEmail === 'function' ? email.getEmail() : '';
      const local = addr ? String(addr).split('@')[0] : '';
      if (local) name = local;
    } catch (e) {}
  }
  if (!name) throw new Error('Username required. Pass setPrimaryUsername("yourname") or set Config → username.');
  PropertiesService.getDocumentProperties().setProperty(CALC_SETTINGS.userPropertyKey, String(name));
}

function getPrimaryUsername_() {
  return PropertiesService.getDocumentProperties().getProperty(CALC_SETTINGS.userPropertyKey) || '';
}

function classifyModeAndVariant_(rules, timeClass) {
  const rulesNorm = String(rules || 'chess').toLowerCase();
  const timeNorm = String(timeClass || '').toLowerCase();
  // Mode: bullet/blitz/rapid/daily when rules == chess; otherwise variant name
  // Chess960 also has daily960 as a separate combined label
  if (rulesNorm === 'chess') {
    if (['bullet','blitz','rapid','daily'].includes(timeNorm)) return { mode: timeNorm, variant: 'standard' };
    return { mode: 'unknown', variant: 'standard' };
  }
  if (rulesNorm === 'chess960') {
    if (timeNorm === 'daily') return { mode: 'daily960', variant: 'chess960' };
    if (['bullet','blitz','rapid'].includes(timeNorm)) return { mode: 'chess960', variant: 'chess960' };
    return { mode: 'chess960', variant: 'chess960' };
  }
  // For other variants, mode is simply the variant name
  return { mode: rulesNorm, variant: rulesNorm };
}

function resultToScore_(winnerColor, perspectiveColor) {
  if (!winnerColor) return 0.5; // draw or unknown
  if (winnerColor === 'draw') return 0.5;
  return winnerColor === perspectiveColor ? 1 : 0;
}

function winnerColorFromResult_(resultRaw) {
  const r = String(resultRaw || '').trim();
  if (r === '1-0') return 'white';
  if (r === '0-1') return 'black';
  if (r === '1/2-1/2') return 'draw';
  return '';
}

function fetchOpponentStats_(username) {
  if (!username) return null;
  const cache = CacheService.getDocumentCache();
  const key = `oppstats:${username.toLowerCase()}`;
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }
  const url = `https://api.chess.com/pub/player/${encodeURIComponent(username)}/stats`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'Accept-Encoding': 'gzip' } });
  if (res.getResponseCode() !== 200) return null;
  const data = JSON.parse(res.getContentText());
  cache.put(key, JSON.stringify(data), CALC_SETTINGS.cacheMinutes * 60);
  return data;
}

function extractOpponentSummary_(stats) {
  if (!stats || typeof stats !== 'object') return {};
  function getRating(path) {
    const obj = path.reduce((acc, k) => (acc && acc[k] ? acc[k] : null), stats);
    if (!obj) return '';
    // prefer last.rating else best.rating else highest.rating
    return (obj.last && obj.last.rating) || (obj.best && obj.best.rating) || (obj.highest && obj.highest.rating) || '';
  }
  return {
    opp_bullet_rating: getRating(['chess_bullet']),
    opp_blitz_rating: getRating(['chess_blitz']),
    opp_rapid_rating: getRating(['chess_rapid']),
    opp_daily_rating: getRating(['chess_daily']),
    opp_daily960_rating: getRating(['chess_daily960']) || '',
    opp_chess960_rating: getRating(['chess960']) || ''
  };
}

function buildCalculatedFromRaw() {
  const username = getPrimaryUsername_();
  if (!username) throw new Error('Set your username with setPrimaryUsername("yourname") first.');
  const ss = getSpreadsheet_();
  const raw = ss.getSheetByName(HEADER_CONFIG.rawDataSheetName);
  if (!raw) throw new Error('Missing sheet: ' + HEADER_CONFIG.rawDataSheetName);
  let calc = ss.getSheetByName(HEADER_CONFIG.calcDataSheetName);
  if (!calc) calc = ss.insertSheet(HEADER_CONFIG.calcDataSheetName);

  const rawHeaders = readHeaderRow_(raw);
  if (!rawHeaders.length) throw new Error('Raw sheet has no headers.');

  const idx = new Map(rawHeaders.map((h, i) => [h, i]));

  // Define calculated headers
  const calcHeaders = [
    'url', 'date', 'time', 'white.username', 'black.username', 'white.rating', 'black.rating', 'result', 'termination',
    'rules', 'time_class', 'mode', 'variant',
    'my_color', 'my_score', 'my_result',
    'my_bullet_rating', 'my_blitz_rating', 'my_rapid_rating', 'my_daily_rating', 'my_chess960_rating', 'my_daily960_rating',
    'opp_bullet_rating', 'opp_blitz_rating', 'opp_rapid_rating', 'opp_daily_rating', 'opp_chess960_rating', 'opp_daily960_rating'
  ];

  // Initialize output
  calc.clear();
  calc.getRange(1, 1, 1, calcHeaders.length).setValues([calcHeaders]);
  calc.setFrozenRows(1);

  const lastRow = raw.getLastRow();
  if (lastRow <= 1) return;
  const values = raw.getRange(2, 1, lastRow - 1, raw.getLastColumn()).getValues();

  let running = {
    bullet: null,
    blitz: null,
    rapid: null,
    daily: null,
    chess960: null,
    daily960: null
  };

  const out = [];
  for (const row of values) {
    function val(h) { const i = idx.get(h); return i != null ? row[i] : ''; }
    const whiteUser = String(val('white.username') || val('white') || '').trim();
    const blackUser = String(val('black.username') || val('black') || '').trim();
    const myColor = whiteUser.toLowerCase() === username.toLowerCase() ? 'white' : (blackUser.toLowerCase() === username.toLowerCase() ? 'black' : '');
    const rules = val('rules');
    const timeClass = val('time_class') || val('time-class') || '';
    const { mode, variant } = classifyModeAndVariant_(rules, timeClass);

    const resultRaw = val('result') || val('Result') || '';
    const winner = winnerColorFromResult_(resultRaw);
    const myScore = myColor ? resultToScore_(winner, myColor) : '';
    const myResult = myScore === '' ? '' : (myScore === 1 ? 'Win' : (myScore === 0.5 ? 'Draw' : 'Loss'));

    // Ratings
    const whiteRating = Number(val('white.rating') || val('WhiteElo') || '') || '';
    const blackRating = Number(val('black.rating') || val('BlackElo') || '') || '';

    // Update running ratings only for the mode being played; others carry forward previous value
    const modes = ['bullet','blitz','rapid','daily','chess960','daily960'];
    const runningNext = { ...running };
    let playedRating = '';
    if (myColor === 'white') playedRating = whiteRating; else if (myColor === 'black') playedRating = blackRating;
    if (modes.includes(mode)) {
      runningNext[mode] = playedRating !== '' ? Number(playedRating) : running[mode];
    }
    running = runningNext;

    // Opponent stats (cached per username)
    const oppUser = myColor === 'white' ? blackUser : (myColor === 'black' ? whiteUser : '');
    const oppStats = oppUser ? extractOpponentSummary_(fetchOpponentStats_(oppUser)) : {};

    out.push([
      val('url') || '',
      val('UTCDate') || val('Date') || '',
      val('UTCTime') || '',
      whiteUser || '',
      blackUser || '',
      whiteRating,
      blackRating,
      resultRaw || '',
      val('termination') || val('Termination') || '',
      rules || '',
      timeClass || '',
      mode,
      variant,
      myColor || '',
      myScore,
      myResult,
      running.bullet,
      running.blitz,
      running.rapid,
      running.daily,
      running.chess960,
      running.daily960,
      oppStats.opp_bullet_rating || '',
      oppStats.opp_blitz_rating || '',
      oppStats.opp_rapid_rating || '',
      oppStats.opp_daily_rating || '',
      oppStats.opp_chess960_rating || '',
      oppStats.opp_daily960_rating || ''
    ]);
  }

  if (out.length) calc.getRange(2, 1, out.length, calcHeaders.length).setValues(out);
}

/**
 * Provide a canonical list of calculated headers with implied formulas and metadata.
 * Formulas are expressed as templates; adapt to your column letters if implementing in-sheets.
 */
function getCalculatedHeaderDefinitions_() {
  // Columns referenced below assume Raw sheet has columns named as in headers and Calculated has the headers we output.
  // Use placeholders like ROW for the current row number.
  return [
    {
      name: 'url',
      formula: '=Raw!A{ROW}',
      description: 'Direct link to the game on Chess.com',
      example: 'https://www.chess.com/game/live/1234567890',
      fieldType: 'string',
      category: 'Meta',
      notes: 'Adjust cell reference to match your Raw column for url.'
    },
    {
      name: 'date',
      formula: '=IF(LEN(Raw!B{ROW}), Raw!B{ROW}, Raw!C{ROW})',
      description: 'Prefers UTCDate else Date from PGN',
      example: '2025.08.14',
      fieldType: 'date-string',
      category: 'Time',
      notes: 'Pick the Raw column letters for UTCDate/Date.'
    },
    {
      name: 'time',
      formula: '=Raw!D{ROW}',
      description: 'UTCTime',
      example: '18:42:12',
      fieldType: 'time-string',
      category: 'Time',
      notes: 'Use the UTCTime column from Raw.'
    },
    {
      name: 'white.username',
      formula: '=Raw!E{ROW}',
      description: 'White username',
      example: 'ians141',
      fieldType: 'string',
      category: 'Players',
      notes: ''
    },
    {
      name: 'black.username',
      formula: '=Raw!F{ROW}',
      description: 'Black username',
      example: 'OpponentUser',
      fieldType: 'string',
      category: 'Players',
      notes: ''
    },
    {
      name: 'white.rating',
      formula: '=IF(LEN(Raw!G{ROW}), Raw!G{ROW}, Raw!H{ROW})',
      description: 'White rating from JSON or PGN Elo',
      example: '2120',
      fieldType: 'number',
      category: 'Ratings',
      notes: 'Choose correct Raw columns (white.rating vs WhiteElo).'
    },
    {
      name: 'black.rating',
      formula: '=IF(LEN(Raw!I{ROW}), Raw!I{ROW}, Raw!J{ROW})',
      description: 'Black rating from JSON or PGN Elo',
      example: '2155',
      fieldType: 'number',
      category: 'Ratings',
      notes: 'Choose correct Raw columns (black.rating vs BlackElo).'
    },
    {
      name: 'result',
      formula: '=Raw!K{ROW}',
      description: 'Game result string',
      example: '1-0',
      fieldType: 'enum',
      category: 'Result',
      notes: 'Expected values: 1-0, 0-1, 1/2-1/2.'
    },
    {
      name: 'termination',
      formula: '=IF(LEN(Raw!L{ROW}), Raw!L{ROW}, Raw!M{ROW})',
      description: 'Termination reason',
      example: 'Normal',
      fieldType: 'string',
      category: 'Result',
      notes: 'Use either JSON termination or PGN Termination.'
    },
    {
      name: 'rules',
      formula: '=Raw!N{ROW}',
      description: 'Game rules/variant',
      example: 'chess',
      fieldType: 'enum',
      category: 'Variant',
      notes: 'Examples: chess, chess960, etc.'
    },
    {
      name: 'time_class',
      formula: '=Raw!O{ROW}',
      description: 'Bullet/Blitz/Rapid/Daily',
      example: 'rapid',
      fieldType: 'enum',
      category: 'Time',
      notes: ''
    },
    {
      name: 'mode',
      formula: '=IF(N{ROW}="chess", O{ROW}, IF(N{ROW}="chess960", IF(O{ROW}="daily", "daily960", "chess960"), N{ROW}))',
      description: 'Mode per rules/time_class',
      example: 'rapid',
      fieldType: 'enum',
      category: 'Variant',
      notes: 'Replicates classifyModeAndVariant_ logic.'
    },
    {
      name: 'variant',
      formula: '=IF(N{ROW}="chess", "standard", N{ROW})',
      description: 'Variant label',
      example: 'standard',
      fieldType: 'enum',
      category: 'Variant',
      notes: ''
    },
    {
      name: 'my_color',
      formula: '=IF(LOWER(E{ROW})=LOWER($B$1), "white", IF(LOWER(F{ROW})=LOWER($B$1), "black", ""))',
      description: 'Your color this game (requires PRIMARY_USERNAME in B1)',
      example: 'white',
      fieldType: 'enum',
      category: 'Players',
      notes: 'Put your username in Calculated!B1 or adapt reference.'
    },
    {
      name: 'my_score',
      formula: '=IF(LEN(P{ROW})=0, "", IF(K{ROW}="1-0", IF(P{ROW}="white",1,0), IF(K{ROW}="0-1", IF(P{ROW}="black",1,0), 0.5)))',
      description: 'Numeric score from result under your color',
      example: '1',
      fieldType: 'number',
      category: 'Result',
      notes: ''
    },
    {
      name: 'my_result',
      formula: '=IF(Q{ROW}="", "", IF(Q{ROW}=1, "Win", IF(Q{ROW}=0.5, "Draw", "Loss")))',
      description: 'Win/Draw/Loss from my_score',
      example: 'Win',
      fieldType: 'enum',
      category: 'Result',
      notes: ''
    },
    {
      name: 'my_bullet_rating',
      formula: '=IF(R{ROW}="bullet", IF(P{ROW}="white", G{ROW}, IF(P{ROW}="black", I{ROW}, R{ROW}-1)), R{ROW}-1)',
      description: 'Running bullet rating (carry-forward else current)',
      example: '2105',
      fieldType: 'number',
      category: 'Ratings',
      notes: 'Use previous row value in place of R{ROW}-1; spreadsheets need exact prev cell reference.'
    },
    {
      name: 'my_blitz_rating',
      formula: '=IF(R{ROW}="blitz", IF(P{ROW}="white", G{ROW}, IF(P{ROW}="black", I{ROW}, S{ROW}-1)), S{ROW}-1)',
      description: 'Running blitz rating',
      example: '2150',
      fieldType: 'number',
      category: 'Ratings',
      notes: 'Replace S{ROW}-1 with previous row cell for this column.'
    },
    {
      name: 'my_rapid_rating',
      formula: '=IF(R{ROW}="rapid", IF(P{ROW}="white", G{ROW}, IF(P{ROW}="black", I{ROW}, T{ROW}-1)), T{ROW}-1)',
      description: 'Running rapid rating',
      example: '2200',
      fieldType: 'number',
      category: 'Ratings',
      notes: 'Replace T{ROW}-1 with previous row cell for this column.'
    },
    {
      name: 'my_daily_rating',
      formula: '=IF(R{ROW}="daily", IF(P{ROW}="white", G{ROW}, IF(P{ROW}="black", I{ROW}, U{ROW}-1)), U{ROW}-1)',
      description: 'Running daily rating',
      example: '1800',
      fieldType: 'number',
      category: 'Ratings',
      notes: 'Replace U{ROW}-1 with previous row cell for this column.'
    },
    {
      name: 'my_chess960_rating',
      formula: '=IF(R{ROW}="chess960", IF(P{ROW}="white", G{ROW}, IF(P{ROW}="black", I{ROW}, V{ROW}-1)), V{ROW}-1)',
      description: 'Running Chess960 rating',
      example: '1750',
      fieldType: 'number',
      category: 'Ratings',
      notes: 'Replace V{ROW}-1 with previous row cell for this column.'
    },
    {
      name: 'my_daily960_rating',
      formula: '=IF(R{ROW}="daily960", IF(P{ROW}="white", G{ROW}, IF(P{ROW}="black", I{ROW}, W{ROW}-1)), W{ROW}-1)',
      description: 'Running Daily960 rating',
      example: '1700',
      fieldType: 'number',
      category: 'Ratings',
      notes: 'Replace W{ROW}-1 with previous row cell for this column.'
    },
    {
      name: 'opp_bullet_rating',
      formula: 'N/A (API-derived)',
      description: 'Opponent bullet rating (Chess.com stats)',
      example: '2350',
      fieldType: 'number',
      category: 'Opponent',
      notes: 'Only available via API; no pure formula.'
    },
    {
      name: 'opp_blitz_rating',
      formula: 'N/A (API-derived)',
      description: 'Opponent blitz rating',
      example: '2400',
      fieldType: 'number',
      category: 'Opponent',
      notes: ''
    },
    {
      name: 'opp_rapid_rating',
      formula: 'N/A (API-derived)',
      description: 'Opponent rapid rating',
      example: '2300',
      fieldType: 'number',
      category: 'Opponent',
      notes: ''
    },
    {
      name: 'opp_daily_rating',
      formula: 'N/A (API-derived)',
      description: 'Opponent daily rating',
      example: '2100',
      fieldType: 'number',
      category: 'Opponent',
      notes: ''
    },
    {
      name: 'opp_chess960_rating',
      formula: 'N/A (API-derived)',
      description: 'Opponent Chess960 rating',
      example: '2000',
      fieldType: 'number',
      category: 'Opponent',
      notes: ''
    },
    {
      name: 'opp_daily960_rating',
      formula: 'N/A (API-derived)',
      description: 'Opponent Daily960 rating',
      example: '1950',
      fieldType: 'number',
      category: 'Opponent',
      notes: ''
    }
  ];
}

function GenerateCalculatedDataFormulas() {
  const ss = getSpreadsheet_();
  const sheetName = 'Calculated  Data Formulas'; // Note: two spaces per request
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clear({ contentsOnly: true });
  const headers = ['Header', 'Formula (Implied)', 'Description', 'Example', 'Field Type', 'Category', 'Notes'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  const defs = getCalculatedHeaderDefinitions_();
  const rows = defs.map(d => [d.name, d.formula, d.description, d.example, d.fieldType, d.category, d.notes || '']);
  if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  // Color by category
  const catIdx = headers.indexOf('Category') + 1;
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const cats = sh.getRange(2, catIdx, lastRow - 1, 1).getValues().map(r => r[0]);
    const colors = cats.map(c => [colorForCategory_(c)]);
    sh.getRange(2, catIdx, lastRow - 1, 1).setBackgrounds(colors);
  }
  sh.autoResizeColumns(1, headers.length);
}

function InteractiveAddCalculatedHeaderWithAI() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getSpreadsheet_().getSheetByName('Calculated  Data Formulas') || getSpreadsheet_().insertSheet('Calculated  Data Formulas');
  if (sheet.getLastRow() === 0) GenerateCalculatedDataFormulas();

  const qName = ui.prompt('New Calculated Header', 'Header name (e.g., my_accuracy_gap):', ui.ButtonSet.OK_CANCEL);
  if (qName.getSelectedButton() !== ui.Button.OK) return;
  const qDesc = ui.prompt('Description', 'Describe what this field should represent:', ui.ButtonSet.OK_CANCEL);
  if (qDesc.getSelectedButton() !== ui.Button.OK) return;
  const qInputs = ui.prompt('Inputs', 'List raw/calculated fields you would use (comma separated):', ui.ButtonSet.OK_CANCEL);
  if (qInputs.getSelectedButton() !== ui.Button.OK) return;
  const qType = ui.prompt('Field Type', 'Type (number, string, enum, date, time):', ui.ButtonSet.OK_CANCEL);
  if (qType.getSelectedButton() !== ui.Button.OK) return;
  const qCategory = ui.prompt('Category', 'Category (Meta, Players, Ratings, Time, PGN, Result, Variant, Opponent, Derived):', ui.ButtonSet.OK_CANCEL);
  if (qCategory.getSelectedButton() !== ui.Button.OK) return;

  const payload = {
    name: qName.getResponseText().trim(),
    description: qDesc.getResponseText().trim(),
    inputs: qInputs.getResponseText().trim(),
    fieldType: qType.getResponseText().trim(),
    category: qCategory.getResponseText().trim()
  };

  const ai = generateWithGemini_(payload);
  const formula = ai.formula || 'N/A (script recommended)';
  const example = ai.example || '';
  const notes = ai.notes || '';

  const headerRow = ['Header', 'Formula (Implied)', 'Description', 'Example', 'Field Type', 'Category', 'Notes'];
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
  sheet.appendRow([payload.name, formula, payload.description, example, payload.fieldType, payload.category, notes]);
}

function generateWithGemini_(payload) {
  const apiKey = PropertiesService.getDocumentProperties().getProperty('GENAI_API_KEY');
  const prompt = `You are to propose a Google Sheets A1-style formula (or return 'N/A' if not feasible) to compute a calculated chess game field.\n` +
    `Field Name: ${payload.name}\n` +
    `Description: ${payload.description}\n` +
    `Inputs: ${payload.inputs}\n` +
    `Field Type: ${payload.fieldType}\n` +
    `Category: ${payload.category}\n` +
    `Assume a Raw sheet with common columns (url, UTCDate, UTCTime, white.username, black.username, white.rating, black.rating, result, termination, rules, time_class).` +
    `When formulas depend on previous row, indicate placeholder PREV_CELL. Provide JSON: {"formula":"...","example":"...","notes":"..."}`;

  if (!apiKey) {
    // Fallback heuristic without AI
    return {
      formula: 'N/A (set GENAI_API_KEY to enable AI formula synthesis)',
      example: '',
      notes: 'No API key configured.'
    };
  }
  try {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + encodeURIComponent(apiKey);
    const body = {
      contents: [{ parts: [{ text: prompt }] }]
    };
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return { formula: 'N/A', example: '', notes: 'AI call failed: ' + res.getResponseCode() };
    const data = JSON.parse(res.getContentText());
    const text = (((data.candidates || [])[0] || {}).content || {}).parts ? (((data.candidates[0].content.parts[0] || {}).text) || '') : '';
    let obj;
    try { obj = JSON.parse(text); } catch (e) { obj = { formula: text }; }
    return obj || { formula: 'N/A' };
  } catch (e) {
    return { formula: 'N/A', example: '', notes: String(e) };
  }
}

