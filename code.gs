// ===== ChessManager.gs (Part 1/5) =====
// Enhanced Chess.com Game Data Manager
// This script manages Chess.com game data across multiple sheets

// =================================================================
// SCRIPT CONFIGURATION
// =================================================================

const SHEETS = {
  CONFIG: 'Config',
  ARCHIVES: 'Archives',
  GAMES: 'Game Data',
  DAILY: 'Daily Data',
  STATS: 'Player Stats',
  PROFILE: 'Player Profile',
  LOGS: 'Execution Logs',
  OPENINGS_URLS: 'Opening URLS',
  ADD_OPENINGS: 'Add Openings',
  CHESS_ECO: 'Chess ECO',
  TRIGGERS: 'Triggers'
};

const HEADERS = {
  CONFIG: ['Setting', 'Value', 'Description'],
  ARCHIVES: ['Archive URL', 'Year-Month', 'Last Updated'],
  GAMES: [
    // Core game/meta
    'Game URL', 'Time Control', 'Base Time (min)', 'Increment (sec)', 'Rated',
    'Time Class', 'Rules', 'Format',
    // Timing (legacy single End Time + duration)
    'End Time', 'Game Duration (sec)',

    // Players/opponent/result (summary)
    'My Rating', 'My Color', 'Opponent', 'Opponent Rating', 'Result',
    'Termination', 'Winner',

    // PGN classic headers
    'Event', 'Site', 'Date', 'Round', 'Opening', 'ECO', 'ECO URL',

    // Opening placeholders (user-derived) + helper
    'ECO URL Tail', 'Opening Canonical', 'Opening Variation', 'Opening Subvariation',

    // UTC splits (date/time only)
    'UTC Start Date', 'UTC Start Time', 'UTC End Date', 'UTC End Time',

    // Local splits (date/time only)
    'Local Start Date', 'Local Start Time', 'Local End Date', 'Local End Time',

    // Timezone/context
    'PGN Timezone Label', 'Implied Local Timezone', 'Local − UTC Offset (hrs)',

    // PGN timing raw for reference (kept for compatibility)
    'UTC Date', 'UTC Time', 'PGN Start Time', 'PGN End Date', 'PGN End Time',

    // Positions and PGN payload
    'Current Position', 'Final FEN', 'Full PGN',

    // Moves (CSV) and quick counts
    'Moves', 'Times', 'Moves Per Side',

    // Identity/timezone (legacy)
    'My Timezone', 'Local Start Time (Live)', 'Hour Differential (hrs)',

    // Identity fields and rating delta
    'My Username', 'Rating Change', 'My Player ID', 'My UUID',
    'Opponent Color', 'Opponent Username', 'Opponent Player ID', 'Opponent UUID',

    // Accuracies
    'My Accuracy', 'Opponent Accuracy',

    // Opening classification (if ever used manually; kept separate)
    'Opening from URL', 'Opening from ECO',

    // Data quality
    'Data Warnings'
  ],
  DAILY: [
    'Date',
    'Bullet Win', 'Bullet Loss', 'Bullet Draw', 'Bullet Rating', 'Bullet Change', 'Bullet Time',
    'Blitz Win', 'Blitz Loss', 'Blitz Draw', 'Blitz Rating', 'Blitz Change', 'Blitz Time',
    'Rapid Win', 'Rapid Loss', 'Rapid Draw', 'Rapid Rating', 'Rapid Change', 'Rapid Time',
    'Total Games', 'Total Wins', 'Total Losses', 'Total Draws', 'Rating Sum', 'Total Rating Change', 'Total Time (sec)', 'Avg Game Duration (sec)'
  ],
  LOGS: ['Timestamp', 'Function', 'Username', 'Status', 'Execution Time (ms)', 'Notes'],
  STATS: ['Field', 'Value'],
  PROFILE: ['Field', 'Value']
};

// Functions allowed to be scheduled via Triggers sheet
const TRIGGERABLE_FUNCTIONS = [
  'quickUpdate',
  'completeUpdate',
  'refreshRecentData',
  'updateArchives',
  'updateDailyData'
];

const CONFIG_KEYS = {
  USERNAME: 'username',
  BASE_API: 'base_api_url'
};

const DEFAULT_CONFIG = [
  [CONFIG_KEYS.USERNAME, '', 'Your Chess.com username (lowercase)'],
  [CONFIG_KEYS.BASE_API, 'https://api.chess.com/pub', 'Base API URL']
];

// =================================================================
// MENU & UI
// =================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const dailyMenu = ui.createMenu('Daily Automation')
    .addItem('Run Daily Roll (now)', 'dailyRoll')
    .addItem('Install Midnight Trigger', 'installDailyTrigger')
    .addItem('Remove Midnight Triggers', 'removeDailyTriggers')
    .addItem('Setup Game Data Columns', 'ensureGameDataComputedColumns');

  const triggersMenu = ui.createMenu('Trigger Management')
    .addItem('Set up Triggers Sheet', 'setupTriggersSheet')
    .addItem('✅ Apply Trigger Settings', 'applyTriggers')
    .addItem('❌ Delete All Triggers', 'deleteAllTriggers');

  const individualMenu = ui.createMenu('Individual Actions')
    .addItem('Fetch & Process Current Month', 'refreshRecentData')
    .addItem('Update Archives List', 'updateArchives')
    .addItem('Fetch Current Month Games', 'fetchCurrentMonthGames')
    .addItem('Process Daily Data', 'updateDailyData')
    .addItem('Update Player Stats', 'updateStats')
    .addItem('Update Player Profile', 'updateProfile');

  ui.createMenu('Chess.com Manager')
    .addItem('▶️ Quick Update (Recommended)', 'quickUpdate')
    .addItem('🔄 Complete Update (All Games)', 'completeUpdate')
    .addSeparator()
    .addItem('📝 Categorize Openings from URL', 'categorizeOpeningsFromUrl')
    .addItem('📊 Categorize Openings from ECO', 'categorizeOpeningsFromEco')
    .addSeparator()
    .addSubMenu(dailyMenu)
    .addSeparator()
    .addSubMenu(triggersMenu)
    .addSeparator()
    .addSubMenu(individualMenu)
    .addSeparator()
    .addItem('Run Initial Setup (Once)', 'setupSheets')
    .addToUi();
}

function onInstall() {
  onOpen();
}

// =================================================================
// SETUP & CONFIG
// =================================================================

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheetWithHeaders(ss, SHEETS.CONFIG, HEADERS.CONFIG);
  ensureSheetWithHeaders(ss, SHEETS.ARCHIVES, HEADERS.ARCHIVES);
  ensureSheetWithHeaders(ss, SHEETS.GAMES, HEADERS.GAMES);
  ensureSheetWithHeaders(ss, SHEETS.DAILY, HEADERS.DAILY);
  ensureSheetWithHeaders(ss, SHEETS.STATS, HEADERS.STATS);
  ensureSheetWithHeaders(ss, SHEETS.PROFILE, HEADERS.PROFILE);
  ensureSheetWithHeaders(ss, SHEETS.LOGS, HEADERS.LOGS);
  ensureSheetWithHeaders(ss, SHEETS.OPENINGS_URLS, [['Family', 'Base URL']]);
  ensureSheetWithHeaders(ss, SHEETS.ADD_OPENINGS, [['URLs to Categorize']]);
  ensureSheetWithHeaders(ss, SHEETS.CHESS_ECO, [['ECO', 'Opening Name']]);

  const configSheet = ss.getSheetByName(SHEETS.CONFIG);
  if (configSheet.getLastRow() < 2) {
    configSheet.getRange(2, 1, DEFAULT_CONFIG.length, 3).setValues(
      DEFAULT_CONFIG.map(([k, v, d]) => [k, v, d])
    );
  }

  applyNumberFormatsForNewColumns();

  SpreadsheetApp.getActiveSpreadsheet().toast('Sheets initialized. Fill in Config → username.', 'Setup Complete', 8);
}

function ensureSheetWithHeaders(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const headerRow = Array.isArray(headers[0]) ? headers[0] : headers;
  sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function getUsername() {
  const username = getConfig(CONFIG_KEYS.USERNAME);
  if (!username) throw new Error('Config username is missing. Open Config sheet and set "username".');
  return username.toString().trim().toLowerCase();
}

function updateConfig(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG);
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value, '']);
}

function getConfig(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.CONFIG);
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === key) return values[i][1];
  }
  return '';
}

function logExecution(fnName, username, status, execMs, notes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS.LOGS);
    sheet.appendRow([
      new Date(),
      fnName,
      username || '',
      status,
      execMs,
      notes || ''
    ]);
  } catch (e) {
    console.log('Log error: ' + e.message);
  }
}

// =================================================================
// CHESS.COM API HELPERS
// =================================================================

function buildApiUrl(path) {
  const base = getConfig(CONFIG_KEYS.BASE_API) || 'https://api.chess.com/pub';
  return base.replace(/\/$/, '') + '/' + path.replace(/^\//, '');
}

function fetchJson(url) {
  const maxAttempts = 5;
  const baseSleepMs = 500;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, validateHttpsCertificates: true });
      const code = resp.getResponseCode();
      if (code >= 200 && code < 300) return JSON.parse(resp.getContentText());
      if (code === 429 || (code >= 500 && code < 600)) {
        const sleepMs = baseSleepMs * Math.pow(2, attempt - 1) + Math.floor(Math.random() * 250);
        Utilities.sleep(sleepMs);
        continue;
      }
      throw new Error('HTTP ' + code + ' for ' + url);
    } catch (e) {
      if (attempt === maxAttempts) throw e;
      Utilities.sleep(baseSleepMs * Math.pow(2, attempt - 1));
    }
  }
}

// =================================================================
// PARSERS & MAPPERS
// =================================================================

function parseTimeControl(tc) {
  if (!tc) return { baseMinutes: '', incrementSeconds: '' };
  // Daily controls like "1/259200": leave blank as requested
  if (tc.indexOf('/') >= 0) {
    return { baseMinutes: '', incrementSeconds: '' };
  }
  const [base, inc] = tc.split('+');
  const baseMinutes = base ? Math.round(parseInt(base, 10) / 60) : '';
  const incrementSeconds = inc ? parseInt(inc, 10) : '';
  return { baseMinutes, incrementSeconds };
}

function formatDateTime(epochSeconds) {
  if (!epochSeconds && epochSeconds !== 0) return '';
  const d = new Date(epochSeconds * 1000);
  return d;
}

function formatDate(epochSeconds) {
  if (!epochSeconds && epochSeconds !== 0) return '';
  return new Date(epochSeconds * 1000);
}

function computeGameDuration(startEpoch, endEpoch) {
  if (!startEpoch || !endEpoch) return '';
  const sec = Math.max(0, endEpoch - startEpoch);
  return sec;
}

function computeFormat(rules) {
  if (!rules) return '';
  return rules.toLowerCase() === 'chess960' ? 'Chess960' : 'Standard';
}

function getPlayerPerspective(game, username) {
  const whiteUsername = (game.white && game.white.username) ? game.white.username.toLowerCase() : '';
  const blackUsername = (game.black && game.black.username) ? game.black.username.toLowerCase() : '';
  if (whiteUsername === username) return 'white';
  if (blackUsername === username) return 'black';
  return '';
}

function simplifyResult(result) {
  if (!result) return '';
  const r = result.toLowerCase();
  if (r === 'win') return 'Win';
  if (r === 'agreed' || r === 'repetition' || r === 'stalemate' || r === 'insufficient' || r === '50move' || r === 'timevsinsufficient') return 'Draw';
  if (r === 'draw') return 'Draw';
  return 'Loss';
}

function normalizeMethod(causeCandidate, terminationTag) {
  const val = (causeCandidate || '').toString().toLowerCase();
  const term = (terminationTag || '').toString().toLowerCase();
  const map = {
    checkmated: 'Checkmated',
    checkmate: 'Checkmated',
    resigned: 'Resigned',
    timeout: 'Timeout',
    abandoned: 'Abandoned',
    stalemate: 'Stalemate',
    agreed: 'Draw agreed',
    repetition: 'Draw by repetition',
    insufficient: 'Insufficient material',
    '50move': 'Draw by 50-move rule',
    timevsinsufficient: 'Draw by timeout vs insufficient material',
    kingofthehill: 'Opponent king reached the hill',
    threecheck: 'Checked for the 3rd time',
    bughousepartnerlose: 'Bughouse partner lost',
    '': ''
  };
  if (map[val]) return map[val];
  const keys = Object.keys(map);
  for (let i = 0; i < keys.length; i++) {
    if (keys[i] && term.indexOf(keys[i]) !== -1) return map[keys[i]];
  }
  return term ? term.charAt(0).toUpperCase() + term.slice(1) : '';
}

function parseTerminationTag(terminationTag) {
  const raw = (terminationTag || '').toString();
  const t = raw.trim();
  if (!t) return { winner: '', cause: '' };
  const wonMatch = t.match(/^\s*([^\s].*?)\s+won\s+by\s+(.+?)\s*$/i);
  if (wonMatch) {
    return { winner: wonMatch[1], cause: wonMatch[2] };
  }
  const drawMatch = t.match(/\bdrawn\s+by\s+(.+?)\s*$/i);
  if (drawMatch) {
    return { winner: '', cause: drawMatch[1] };
  }
  return { winner: '', cause: t };
}

function normalizePgnDateToDate(pgnDate) {
  if (!pgnDate) return '';
  if (pgnDate instanceof Date) return pgnDate;
  const s = pgnDate.toString();
  const parts = s.split('.');
  if (parts.length !== 3) return '';
  let [y, m, d] = parts;
  if (!/^[0-9?]+$/.test(y) || !/^[0-9?]+$/.test(m) || !/^[0-9?]+$/.test(d)) return '';
  if (y.includes('?') || m.includes('?') || d.includes('?')) return '';
  if (y.length === 3) y = '2' + y;
  if (y.length === 2) y = '20' + y;
  if (y.length !== 4) return '';
  const iso = y + '-' + m.padStart(2, '0') + '-' + d.padStart(2, '0');
  const dt = new Date(iso);
  return isNaN(dt.getTime()) ? '' : dt;
}

function normalizePgnTimeToHms(pgnTime) {
  if (!pgnTime) return '';
  const s = pgnTime.toString().trim();
  if (!s || s.indexOf('?') !== -1) return '';
  const match = s.match(/^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2})(?:\.\d+)?)?$/);
  if (!match) return '';
  let hh = parseInt(match[1], 10);
  let mm = match[2] !== undefined ? parseInt(match[2], 10) : 0;
  let ss = match[3] !== undefined ? parseInt(match[3], 10) : 0;
  if (isNaN(hh) || isNaN(mm) || isNaN(ss)) return '';
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59 || ss < 0 || ss > 59) return '';
  const hStr = String(hh).padStart(2, '0');
  const mStr = String(mm).padStart(2, '0');
  const sStr = String(ss).padStart(2, '0');
  return hStr + ':' + mStr + ':' + sStr;
}

function combinePgnDateAndTimeToDate(pgnDate, pgnTime, assumeUtc) {
  try {
    if (!pgnDate || !pgnTime) return '';
    const ds = pgnDate.toString();
    const parts = ds.split('.');
    if (parts.length !== 3) return '';
    let [y, m, d] = parts;
    if (!/^[0-9]+$/.test(y) || !/^[0-9]+$/.test(m) || !/^[0-9]+$/.test(d)) return '';
    if (y.length === 3) y = '2' + y;
    if (y.length === 2) y = '20' + y;
    if (y.length !== 4) return '';
    const mm = m.padStart(2, '0');
    const dd = d.padStart(2, '0');
    const hms = normalizePgnTimeToHms(pgnTime);
    if (!hms) return '';
    const iso = y + '-' + mm + '-' + dd + 'T' + hms + (assumeUtc ? 'Z' : '');
    const dt = new Date(iso);
    return isNaN(dt.getTime()) ? '' : dt;
  } catch (e) {
    return '';
  }
}

function getLocalTimezoneName() {
  const date = new Date();
  const match = date.toString().match(/\(([^)]+)\)$/);
  return match && match[1] ? match[1] : '';
}

function parsePgnTimezone(tags) {
  if (tags && (tags.UTCDate || tags.UTCTime)) return 'UTC';
  if (tags && (tags.TimeZone || tags.Zone)) return tags.TimeZone || tags.Zone;
  return '';
}

function extractPgnTags(pgn) {
  const tags = {};
  if (!pgn) return tags;
  try {
    // Only parse header section (before first blank line)
    const headerEnd = pgn.indexOf('\n\n');
    const headerText = headerEnd === -1 ? pgn : pgn.slice(0, headerEnd);
    const re = /^\s*\[([A-Za-z0-9_]+)\s+"([^"]*)"\s*\]\s*$/gm;
    let m;
    while ((m = re.exec(headerText)) !== null) {
      const key = m[1];
      const val = m[2];
      tags[key] = val;
    }
    // Normalize a few common aliases if present in non-standard casing
    if (tags.ECOURL && !tags.ECOUrl) tags.ECOUrl = tags.ECOURL;
    if (tags.Timezone && !tags.TimeZone) tags.TimeZone = tags.Timezone;
  } catch (e) {
    // Return best-effort tags on parse errors
  }
  return tags;
}

function computeOffsetHoursFromTz(date, tz) {
  try {
    const z = Utilities.formatDate(date, tz, 'Z'); // +0530 or -0700
    if (!z || z.length < 5) return '';
    const sign = z[0] === '-' ? -1 : 1;
    const hh = parseInt(z.substring(1, 3), 10);
    const mm = parseInt(z.substring(3, 5), 10);
    if (isNaN(hh) || isNaN(mm)) return '';
    return Math.round((sign * (hh + mm / 60)) * 100) / 100;
  } catch (e) {
    return '';
  }
}

function extractEcoUrlTail(ecoUrl) {
  if (!ecoUrl) return '';
  const idx = ecoUrl.indexOf('/openings/');
  if (idx === -1) return '';
  const tail = ecoUrl.substring(idx + '/openings/'.length);
  return tail.replace(/^\/*/, '');
}

function extractMovesFromPgn(pgn) {
  if (!pgn) return { moves: '', times: '', movesPerSide: '' };
  const blankLineIndex = pgn.indexOf('\n\n');
  if (blankLineIndex === -1) return { moves: '', times: '', movesPerSide: '' };
  let movesText = pgn.slice(blankLineIndex + 2).trim();
  movesText = movesText.replace(/\s+(1-0|0-1|1\/2-1\/2|\*)\s*$/, '');
  movesText = movesText.replace(/\([^)]*\)/g, '');
  movesText = movesText.replace(/\$\d+/g, '');

  const movesMap = {};
  const pairRegex = /(\d+)\.\s*(?!\.\.\.)([^\s{}]+)(?:\s*\{([^}]*)\})?(?:\s+(?:(?:\d+)?\.{3}\s*)?([^\s{}]+)(?:\s*\{([^}]*)\})?)?/g;
  let m;
  while ((m = pairRegex.exec(movesText)) !== null) {
    const moveNo = m[1];
    const whiteSan = m[2];
    const whiteComment = m[3] || '';
    const blackSan = m[4];
    const blackComment = m[5] || '';
    if (whiteSan && whiteSan !== '...' && whiteSan !== '.') {
      const wClkMatch = whiteComment.match(/\[%clk\s+([0-9:\.]+)\]/);
      movesMap[`${moveNo}w`] = [whiteSan, wClkMatch ? wClkMatch[1] : ''];
    }
    if (blackSan && blackSan !== '...' && blackSan !== '.') {
      const bClkMatch = blackComment.match(/\[%clk\s+([0-9:\.]+)\]/);
      movesMap[`${moveNo}b`] = [blackSan, bClkMatch ? bClkMatch[1] : ''];
    }
  }
  const blackOnlyRegex = /(\d+)\.\s*\.\.\.\s*([^\s{}]+)(?:\s*\{([^}]*)\})?/g;
  while ((m = blackOnlyRegex.exec(movesText)) !== null) {
    const moveNo = m[1];
    const blackSan = m[2];
    const blackComment = m[3] || '';
    if (!movesMap[`${moveNo}b`]) {
      const bClkMatch = blackComment.match(/\[%clk\s+([0-9:\.]+)\]/);
      movesMap[`${moveNo}b`] = [blackSan, bClkMatch ? bClkMatch[1] : ''];
    }
  }

  const sanList = [];
  const clkList = [];
  const maxMove = Object.keys(movesMap).reduce((mx, k) => {
    const n = parseInt(k, 10);
    return Number.isFinite(n) ? Math.max(mx, n) : mx;
  }, 0);
  for (let n = 1; n <= maxMove; n++) {
    const w = movesMap[`${n}w`];
    if (w) { sanList.push(w[0] || ''); clkList.push(w[1] || ''); }
    const b = movesMap[`${n}b`];
    if (b) { sanList.push(b[0] || ''); clkList.push(b[1] || ''); }
  }
  const movesPerSide = maxMove || '';
  return { moves: sanList.join(','), times: clkList.join(','), movesPerSide };
}

// ===== ChessManager.gs (Part 2/5) =====

// =================================================================
// ROW BUILDER
// =================================================================

function gameToRow(game, username) {
  const headers = HEADERS.GAMES;
  const tc = parseTimeControl(game.time_control);
  const myColor = getPlayerPerspective(game, username);
  const my = myColor === 'white' ? game.white : game.black;
  const opp = myColor === 'white' ? game.black : game.white;
  const myRating = my && my.rating ? my.rating : '';
  const oppRating = opp && opp.rating ? opp.rating : '';
  const resultRaw = myColor === 'white' ? game.white_result : game.black_result;
  const result = simplifyResult(resultRaw);
  const pgn = game.pgn || '';
  const tags = extractPgnTags(pgn);
  const mv = extractMovesFromPgn(pgn);
  const parsedTerm = parseTerminationTag(tags.Termination || '');
  const termination = normalizeMethod(parsedTerm.cause || resultRaw, tags.Termination || '');

  let winner = '';
  const myUsername = my && my.username ? my.username : '';
  if (parsedTerm.winner) {
    winner = parsedTerm.winner;
  } else if (result === 'Win') {
    winner = myUsername || (myColor || '').charAt(0).toUpperCase() + (myColor || '').slice(1);
  } else if (result === 'Loss') {
    winner = opp && opp.username ? opp.username : (myColor === 'white' ? 'Black' : (myColor === 'black' ? 'White' : ''));
  } else {
    winner = 'Draw';
  }

  // Prefer PGN-based UTC Start; fallback to JSON start_time for daily
  const utcStart = (tags && tags.UTCDate && tags.UTCTime)
    ? combinePgnDateAndTimeToDate(tags.UTCDate, tags.UTCTime, true)
    : (game.time_class === 'daily' && game.start_time ? new Date(game.start_time * 1000) : '');

  // UTC End: if PGN EndDate/EndTime with Timezone=UTC, else JSON end_time
  const pgnTzLabel = parsePgnTimezone(tags);
  const utcEndFromPgn = (tags && tags.EndDate && tags.EndTime && pgnTzLabel === 'UTC')
    ? combinePgnDateAndTimeToDate(tags.EndDate, tags.EndTime, true)
    : '';
  const utcEnd = utcEndFromPgn || (game.end_time ? new Date(game.end_time * 1000) : '');

  // Duration (sec)
  const durationSec = (utcStart && utcEnd)
    ? Math.max(0, Math.round((utcEnd.getTime() - utcStart.getTime()) / 1000))
    : '';

  // Local tz context
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const localTz = ss.getSpreadsheetTimeZone();
  const impliedLocalTz = localTz || getLocalTimezoneName();
  const offsetHours = utcStart ? computeOffsetHoursFromTz(utcStart, impliedLocalTz) : '';

  // Split fields (Sheet displays Dates in spreadsheet tz)
  const utcStartDate = utcStart || '';
  const utcStartTime = utcStart || '';
  const utcEndDate = utcEnd || '';
  const utcEndTime = utcEnd || '';

  const localStartDate = utcStart || '';
  const localStartTime = utcStart || '';
  const localEndDate = utcEnd || '';
  const localEndTime = utcEnd || '';

  // Legacy End Time
  const endCombinedLegacy = utcEnd || (game.end_time ? new Date(game.end_time * 1000) : '');

  // ECO URL and Tail
  const ecoUrlRaw = tags.ECOUrl || game.eco || '';
  const ecoTail = extractEcoUrlTail(ecoUrlRaw);

  // Warnings
  const warnings = [];
  if ((tags.WhiteElo && game.white && game.white.rating && String(tags.WhiteElo) !== String(game.white.rating)) ||
      (tags.BlackElo && game.black && game.black.rating && String(tags.BlackElo) !== String(game.black.rating))) {
    warnings.push('rating-mismatch');
  }
  if (!utcStart) warnings.push('utc-missing');
  if (!utcEnd) warnings.push('utc-missing');
  if (!utcStart && !utcEnd && !game.end_time) warnings.push('time-missing');
  if (!game.accuracies || (game.accuracies && game.accuracies.white == null && game.accuracies.black == null)) {
    warnings.push('accuracy-missing');
  }

  const row = [];
  const push = (v) => row.push(v === undefined ? '' : v);

  // Core/meta
  push(game.url || '');
  push(game.time_control || '');
  push(tc.baseMinutes);
  push(tc.incrementSeconds);
  push(game.rated === true ? 'TRUE' : 'FALSE');
  push(game.time_class || '');
  push(game.rules || '');
  push(computeFormat(game.rules));

  // Legacy End Time + duration
  push(endCombinedLegacy || '');
  push(durationSec);

  // Players/opponent/result
  push(myRating);
  push(myColor);
  push(opp && opp.username ? opp.username : '');
  push(oppRating);
  const resultPhrase = (tags.Termination || '').toString();
  push(resultPhrase);
  push(termination);
  push(winner);

  // PGN headers
  push(tags.Event || '');
  push(tags.Site || '');
  push(normalizePgnDateToDate(tags.Date || ''));
  push(tags.Round || '');
  push(tags.Opening || '');
  push(tags.ECO || '');
  push(ecoUrlRaw || '');

  // Opening placeholders + helper
  push(ecoTail);
  push(''); // Opening Canonical
  push(''); // Opening Variation
  push(''); // Opening Subvariation

  // UTC splits
  push(utcStartDate || '');
  push(utcStartTime || '');
  push(utcEndDate || '');
  push(utcEndTime || '');

  // Local splits
  push(localStartDate || '');
  push(localStartTime || '');
  push(localEndDate || '');
  push(localEndTime || '');

  // Timezone/context
  push(pgnTzLabel || '');
  push(impliedLocalTz || '');
  push(offsetHours);

  // PGN timing raw (compat)
  push(normalizePgnDateToDate(tags.UTCDate || ''));
  push(tags.UTCTime || '');
  push(tags.StartTime || '');
  push(normalizePgnDateToDate(tags.EndDate || ''));
  push(tags.EndTime || '');

  // Positions & PGN
  push(tags.CurrentPosition || '');
  push(game.fen || '');
  push(pgn);

  // Moves CSV and counts
  push(mv.moves);
  push(mv.times);
  push(mv.movesPerSide);

  // Legacy identity/timezone
  push(impliedLocalTz || '');
  push(''); // Local Start Time (Live) legacy
  push(offsetHours);

  // Identity & rating delta
  const myId = my && (my.player_id || my.playerId || my.id) ? (my.player_id || my.playerId || my.id) : '';
  const myUuid = my && (my.uuid || my.uuid4 || my.guid) ? (my.uuid || my.uuid4 || my.guid) : '';
  const oppColor = myColor === 'white' ? 'black' : (myColor === 'black' ? 'white' : '');
  const oppUsername = opp && opp.username ? opp.username : '';
  const oppId = opp && (opp.player_id || opp.playerId || opp.id) ? (opp.player_id || opp.playerId || opp.id) : '';
  const oppUuid = opp && (opp.uuid || opp.uuid4 || opp.guid) ? (opp.uuid || opp.uuid4 || opp.guid) : '';

  push(myUsername);
  push('');       // Rating Change (computed elsewhere)
  push(myId);
  push(myUuid);
  push(oppColor);
  push(oppUsername);
  push(oppId);
  push(oppUuid);

  // Accuracies (blank if missing)
  const myAcc = (game.accuracies && myColor === 'white') ? game.accuracies.white : (game.accuracies ? game.accuracies.black : '');
  const oppAcc = (game.accuracies && myColor === 'white') ? game.accuracies.black : (game.accuracies ? game.accuracies.white : '');
  push(myAcc == null ? '' : myAcc);
  push(oppAcc == null ? '' : oppAcc);

  // Opening classification (kept blank; manual)
  push(''); // Opening from URL
  push(''); // Opening from ECO

  // Warnings
  push(warnings.join(','));

  // Align to headers
  while (row.length < headers.length) row.push('');
  if (row.length > headers.length) row.length = headers.length;
  return row;
}

// =================================================================
// ARCHIVES & GAMES (WITH DEDUPE/UPSERT)
// =================================================================

function fetchArchives(username) {
  const url = buildApiUrl('player/' + encodeURIComponent(username) + '/games/archives');
  const data = fetchJson(url);
  return (data && data.archives) ? data.archives : [];
}

function updateArchivesSheet(username, archives) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ARCHIVES);
  if (archives.length === 0) return;
  const now = new Date();
  const values = archives.map(u => {
    const ym = u.split('/').slice(-2).join('-');
    return [u, ym, now];
  });
  sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  if (sheet.getLastRow() > values.length + 1) {
    sheet.getRange(values.length + 2, 1, sheet.getLastRow() - values.length - 1, sheet.getLastColumn()).clearContent();
  }
}

function getAllArchivesFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ARCHIVES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => r[0]).filter(Boolean);
}

function getCurrentArchive(username) {
  const d = new Date();
  const y = d.getUTCFullYear();
  const m = (d.getUTCMonth() + 1).toString().padStart(2, '0');
  return buildApiUrl('player/' + encodeURIComponent(username) + '/games/' + y + '/' + m);
}

function fetchGamesFromArchive(archiveUrl) {
  const data = fetchJson(archiveUrl);
  return (data && data.games) ? data.games : [];
}

function getGameSheetAndData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GAMES);
  if (!sheet) return { sheet: null, data: [], headers: [] };
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { sheet, data: [], headers: sheet.getRange(1, 1, 1, Math.max(1, lastCol)).getValues()[0] };
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return { sheet, data, headers };
}

function getHeaderIndex(headers, name) {
  return headers.findIndex(h => String(h).trim().toLowerCase() === String(name).trim().toLowerCase());
}

function upsertRows(rows) {
  if (!rows || rows.length === 0) return { added: 0, dup_seen: 0, dup_updated: 0, dup_skipped: 0, utc_missing: 0, acc_missing: 0 };
  const { sheet, data, headers } = getGameSheetAndData();
  if (!sheet) return { added: 0, dup_seen: 0, dup_updated: 0, dup_skipped: 0, utc_missing: 0, acc_missing: 0 };

  const urlIdx = getHeaderIndex(headers, 'Game URL');
  const warnIdx = getHeaderIndex(headers, 'Data Warnings');

  // Map existing URLs
  const urlToRow = new Map();
  for (let i = 0; i < data.length; i++) {
    const u = data[i][urlIdx];
    if (u) urlToRow.set(String(u), i);
  }

  let added = 0, dup_seen = 0, dup_updated = 0, dup_skipped = 0, utc_missing = 0, acc_missing = 0;
  const toAppend = [];

  for (let r = 0; r < rows.length; r++) {
    const newRow = rows[r];
    const newUrl = newRow[urlIdx];
    const warningsNew = (newRow[warnIdx] || '').toString();
    if (warningsNew.indexOf('utc-missing') !== -1) utc_missing++;
    if (warningsNew.indexOf('accuracy-missing') !== -1) acc_missing++;

    if (newUrl && urlToRow.has(String(newUrl))) {
      dup_seen++;
      const i = urlToRow.get(String(newUrl));
      const existing = data[i];
      let updated = false;

      for (let c = 0; c < headers.length && c < newRow.length; c++) {
        const oldVal = existing[c];
        const newVal = newRow[c];
        if ((oldVal === '' || oldVal === null) && newVal !== '' && newVal !== null) {
          existing[c] = newVal;
          updated = true;
        }
      }
      if (updated) {
        const prevWarn = (existing[warnIdx] || '').toString();
        const tokens = prevWarn ? prevWarn.split(',').map(s => s.trim()).filter(Boolean) : [];
        if (tokens.indexOf('duplicate-updated') === -1) tokens.push('duplicate-updated');
        existing[warnIdx] = tokens.join(',');
        dup_updated++;
      } else {
        dup_skipped++;
      }
    } else {
      toAppend.push(newRow);
      if (newUrl) {
        urlToRow.set(String(newUrl), data.length + toAppend.length - 1);
      }
      added++;
    }
  }

  // Write updated existing data
  if (dup_updated > 0 && data.length > 0) {
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }

  // Append new rows
  if (toAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toAppend.length, headers.length).setValues(toAppend);
  }

  // Formats
  applyNumberFormatsForNewColumns();

  // User feedback and logs
  const toast = 'Imported ' + added + ' new. Duplicates seen: ' + dup_seen + ' (' + dup_updated + ' updated, ' + dup_skipped + ' skipped).';
  SpreadsheetApp.getActiveSpreadsheet().toast(toast, 'Ingest Summary', 8);
  logExecution('upsertRows', getUsername(), 'SUCCESS', 0,
    'added=' + added + ', dup_seen=' + dup_seen + ', dup_updated=' + dup_updated + ', dup_skipped=' + dup_skipped +
    ', utc_missing=' + utc_missing + ', acc_missing=' + acc_missing);

  return { added, dup_seen, dup_updated, dup_skipped, utc_missing, acc_missing };
}

function fetchCurrentMonthGames() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const archiveUrl = getCurrentArchive(username);
  const games = fetchGamesFromArchive(archiveUrl);

  const rows = [];
  for (let i = 0; i < games.length; i++) {
    rows.push(gameToRow(games[i], username));
  }
  const res = upsertRows(rows);

  logExecution('fetchCurrentMonthGames', username, 'SUCCESS', new Date().getTime() - startMs,
    'Added ' + res.added + '; dup_seen=' + res.dup_seen + '; dup_updated=' + res.dup_updated + '; dup_skipped=' + res.dup_skipped +
    '; utc_missing=' + res.utc_missing + '; acc_missing=' + res.acc_missing);
}

function fetchAllGames() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const archives = getAllArchivesFromSheet();

  let added = 0, dup_seen = 0, dup_updated = 0, dup_skipped = 0, utc_missing = 0, acc_missing = 0, processed = 0;

  for (let i = 0; i < archives.length; i++) {
    const archiveUrl = archives[i];
    try {
      const games = fetchGamesFromArchive(archiveUrl);
      const rows = [];
      for (let j = 0; j < games.length; j++) {
        rows.push(gameToRow(games[j], username));
      }
      const res = upsertRows(rows);
      added += res.added; dup_seen += res.dup_seen; dup_updated += res.dup_updated; dup_skipped += res.dup_skipped;
      utc_missing += res.utc_missing; acc_missing += res.acc_missing; processed++;
      Utilities.sleep(200);
    } catch (e) {
      console.log('Archive fetch error: ' + archiveUrl + ' => ' + e.message);
    }
  }

  logExecution('fetchAllGames', username, 'SUCCESS', new Date().getTime() - startMs,
    'archives_processed=' + processed + ', added=' + added + ', dup_seen=' + dup_seen + ', dup_updated=' + dup_updated +
    ', dup_skipped=' + dup_skipped + ', utc_missing=' + utc_missing + ', acc_missing=' + acc_missing);
}

function updateArchives() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const archives = fetchArchives(username);
  updateArchivesSheet(username, archives);
  logExecution('updateArchives', username, 'SUCCESS', new Date().getTime() - startMs, 'Archives: ' + archives.length);
}

// ===== ChessManager.gs (Part 3/5) =====

// =================================================================
// DAILY AGGREGATION
// =================================================================

function updateDailyData() {
  const startMs = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gameSheet = ss.getSheetByName(SHEETS.GAMES);
  const dailySheet = ss.getSheetByName(SHEETS.DAILY);
  if (!gameSheet || gameSheet.getLastRow() < 2) return;

  const data = gameSheet.getRange(1, 1, gameSheet.getLastRow(), gameSheet.getLastColumn()).getValues();
  const header = data[0];
  const idx = {
    endTime: header.indexOf('End Time'),
    timeClass: header.indexOf('Time Class'),
    winner: header.indexOf('Winner'),
    myUsername: header.indexOf('My Username'),
    myRating: header.indexOf('My Rating'),
    duration: header.indexOf('Game Duration (sec)')
  };

  const dailyMap = new Map();

  function keyFor(date) {
    const y = date.getFullYear();
    const m = (date.getMonth() + 1).toString().padStart(2, '0');
    const d = date.getDate().toString().padStart(2, '0');
    return y + '-' + m + '-' + d;
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dt = row[idx.endTime];
    const timeClass = (row[idx.timeClass] || '').toString().toLowerCase();
    const winnerVal = (row[idx.winner] || '').toString();
    const myName = (row[idx.myUsername] || '').toString();
    let result = 'Draw';
    if (winnerVal && winnerVal.toLowerCase() !== 'draw') {
      result = (myName && winnerVal.toLowerCase() === myName.toLowerCase()) ? 'Win' : 'Loss';
    }
    const rating = Number(row[idx.myRating] || 0);
    const dur = Number(row[idx.duration] || 0);
    if (!dt || !(dt instanceof Date)) continue;
    const dKey = keyFor(dt);
    if (!dailyMap.has(dKey)) {
      dailyMap.set(dKey, {
        bullet: { w: 0, l: 0, d: 0, rating: 0, change: 0, time: 0, firstTime: null, firstRating: null, lastTime: null, lastRating: null },
        blitz:  { w: 0, l: 0, d: 0, rating: 0, change: 0, time: 0, firstTime: null, firstRating: null, lastTime: null, lastRating: null },
        rapid:  { w: 0, l: 0, d: 0, rating: 0, change: 0, time: 0, firstTime: null, firstRating: null, lastTime: null, lastRating: null },
        totals: { g: 0, w: 0, l: 0, d: 0, ratingSum: 0, change: 0, time: 0 }
      });
    }
    const agg = dailyMap.get(dKey);
    const bucket = (timeClass === 'bullet' || timeClass === 'blitz' || timeClass === 'rapid') ? timeClass : null;
    if (bucket) {
      if (result === 'Win') agg[bucket].w += 1;
      else if (result === 'Loss') agg[bucket].l += 1;
      else if (result === 'Draw') agg[bucket].d += 1;
      if (isFinite(rating) && rating > 0) {
        if (agg[bucket].firstTime === null || dt < agg[bucket].firstTime) {
          agg[bucket].firstTime = dt;
          agg[bucket].firstRating = rating;
        }
        if (agg[bucket].lastTime === null || dt > agg[bucket].lastTime) {
          agg[bucket].lastTime = dt;
          agg[bucket].lastRating = rating;
        }
      }
      agg[bucket].time += dur;
    }
    agg.totals.g += 1;
    if (result === 'Win') agg.totals.w += 1;
    else if (result === 'Loss') agg.totals.l += 1;
    else if (result === 'Draw') agg.totals.d += 1;
    agg.totals.ratingSum += rating;
    agg.totals.time += dur;
  }

  const rows = [];
  const keys = Array.from(dailyMap.keys()).sort();
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const a = dailyMap.get(k);
    const classes = ['bullet', 'blitz', 'rapid'];
    let totalChange = 0;
    for (let ci = 0; ci < classes.length; ci++) {
      const cls = classes[ci];
      const bucket = a[cls];
      bucket.rating = (bucket.lastRating !== null && bucket.lastRating !== undefined) ? bucket.lastRating : 0;
      bucket.change = (bucket.firstRating !== null && bucket.firstRating !== undefined && bucket.lastRating !== null && bucket.lastRating !== undefined)
        ? (bucket.lastRating - bucket.firstRating)
        : 0;
      totalChange += bucket.change;
    }
    const parts = k.split('-');
    const localDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    const avgDur = a.totals.g > 0 ? a.totals.time / a.totals.g : 0;
    rows.push([
      localDate,
      a.bullet.w, a.bullet.l, a.bullet.d, a.bullet.rating, a.bullet.change, a.bullet.time,
      a.blitz.w,  a.blitz.l,  a.blitz.d,  a.blitz.rating,  a.blitz.change,  a.blitz.time,
      a.rapid.w,  a.rapid.l,  a.rapid.d,  a.rapid.rating,  a.rapid.change,  a.rapid.time,
      a.totals.g, a.totals.w, a.totals.l, a.totals.d, a.totals.ratingSum, totalChange, a.totals.time, avgDur
    ]);
  }

  if (rows.length > 0) {
    dailySheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    if (dailySheet.getLastRow() > rows.length + 1) {
      dailySheet.getRange(rows.length + 2, 1, dailySheet.getLastRow() - rows.length - 1, dailySheet.getLastColumn()).clearContent();
    }
  }

  logExecution('updateDailyData', getUsername(), 'SUCCESS', new Date().getTime() - startMs, 'Days: ' + rows.length);
}

// =================================================================
// PROFILE & STATS
// =================================================================

function fetchPlayerStats(username) {
  const url = buildApiUrl('player/' + encodeURIComponent(username) + '/stats');
  return fetchJson(url);
}

function updateStats() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const stats = fetchPlayerStats(username);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.STATS);
  const flat = flattenObject(stats);
  const rows = [['Field', 'Value']];
  rows.push(['Pulled At', new Date()]);
  Object.keys(flat).sort().forEach(k => rows.push([k, flat[k]]));
  sheet.clear();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows).setFontWeight('bold');
  sheet.getRange(2, 1, rows.length - 1, 2).setFontWeight('normal');
  logExecution('updateStats', username, 'SUCCESS', new Date().getTime() - startMs, 'Stats fields: ' + (rows.length - 2));
}

function fetchPlayerProfile(username) {
  const url = buildApiUrl('player/' + encodeURIComponent(username));
  return fetchJson(url);
}

function updateProfile() {
  const startMs = new Date().getTime();
  const username = getUsername();
  const profile = fetchPlayerProfile(username);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PROFILE);
  const entries = Object.entries(profile);
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([[HEADERS.PROFILE[0], HEADERS.PROFILE[1]]]).setFontWeight('bold');
  if (entries.length > 0) sheet.getRange(2, 1, entries.length, 2).setValues(entries);
  logExecution('updateProfile', username, 'SUCCESS', new Date().getTime() - startMs, 'Profile fields: ' + entries.length);
}

function flattenObject(obj, prefix = '', res = {}) {
  if (obj === null || obj === undefined) return res;
  Object.keys(obj).forEach(k => {
    const v = obj[k];
    const p = prefix ? prefix + '.' + k : k;
    if (typeof v === 'object' && v !== null && !Array.isArray(v)) flattenObject(v, p, res);
    else res[p] = Array.isArray(v) ? JSON.stringify(v) : v;
  });
  return res;
}// ===== ChessManager.gs (Part 3/5 continued) =====

// =================================================================
// CATEGORIZATION FUNCTIONS (WITH CACHING) — MANUAL ONLY
// =================================================================

function categorizeOpeningsFromUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gameSheet = ss.getSheetByName(SHEETS.GAMES);
  const lookupSheet = ss.getSheetByName(SHEETS.OPENINGS_URLS);
  let addOpeningsSheet = ss.getSheetByName(SHEETS.ADD_OPENINGS);

  if (!lookupSheet || lookupSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('The "Opening URLS" sheet is missing or empty.');
    return;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'openingUrlMap';
  let openingMap;

  const cachedMap = cache.get(cacheKey);
  if (cachedMap) {
    openingMap = new Map(JSON.parse(cachedMap));
  } else {
    const lookupData = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 2).getValues();
    openingMap = new Map(lookupData.map(row => [row[1], row[0]]));
    cache.put(cacheKey, JSON.stringify(Array.from(openingMap.entries())), 21600);
  }

  if (!addOpeningsSheet) {
    addOpeningsSheet = ss.insertSheet(SHEETS.ADD_OPENINGS);
    addOpeningsSheet.getRange(1, 1).setValue('URLs to Categorize').setFontWeight('bold');
  }

  const existingAddUrls = new Set(addOpeningsSheet.getLastRow() > 1 ? addOpeningsSheet.getRange(2, 1, addOpeningsSheet.getLastRow() - 1, 1).getValues().flat() : []);

  const gameData = gameSheet.getRange(1, 1, gameSheet.getLastRow(), gameSheet.getLastColumn()).getValues();
  const headers = gameData[0];
  const ecoUrlIndex = headers.indexOf('ECO URL');
  const targetColIndex = headers.indexOf('Opening from URL');

  if (ecoUrlIndex === -1 || targetColIndex === -1) { return; }

  let categorizedCount = 0;
  const newUrlsToAdd = [];
  const targetValues = [];

  for (let i = 1; i < gameData.length; i++) {
    const gameEcoUrl = gameData[i][ecoUrlIndex];
    let foundFamily = '';
    if (gameEcoUrl) {
      for (const [baseUrl, familyName] of openingMap.entries()) {
        if (gameEcoUrl.toString().startsWith(baseUrl)) {
          foundFamily = familyName;
          categorizedCount++;
          break;
        }
      }
      if (!foundFamily && !existingAddUrls.has(gameEcoUrl)) {
        newUrlsToAdd.push([gameEcoUrl]);
        existingAddUrls.add(gameEcoUrl);
      }
    }
    targetValues.push([foundFamily]);
  }

  if (targetValues.length > 0) {
    gameSheet.getRange(2, targetColIndex + 1, targetValues.length, 1).setValues(targetValues);
  }
  if (newUrlsToAdd.length > 0) {
    addOpeningsSheet.getRange(addOpeningsSheet.getLastRow() + 1, 1, newUrlsToAdd.length, 1).setValues(newUrlsToAdd);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('URL Categorization complete! Updated ' + categorizedCount + ' games.', 'Success', 8);
}

function categorizeOpeningsFromEco() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gameSheet = ss.getSheetByName(SHEETS.GAMES);
  const lookupSheet = ss.getSheetByName(SHEETS.CHESS_ECO);

  if (!lookupSheet || lookupSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('The "Chess ECO" sheet is missing or empty.');
    return;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'ecoMap';
  let ecoMap;

  const cachedMap = cache.get(cacheKey);
  if (cachedMap) {
    ecoMap = new Map(JSON.parse(cachedMap));
  } else {
    const lookupData = lookupSheet.getRange(2, 1, lookupSheet.getLastRow() - 1, 2).getValues();
    ecoMap = new Map(lookupData.map(row => [row[0].toString().trim(), row[1]]));
    cache.put(cacheKey, JSON.stringify(Array.from(ecoMap.entries())), 21600);
  }

  const gameData = gameSheet.getRange(1, 1, gameSheet.getLastRow(), gameSheet.getLastColumn()).getValues();
  const headers = gameData[0];
  const ecoIndex = headers.indexOf('ECO');
  const targetColIndex = headers.indexOf('Opening from ECO');

  if (ecoIndex === -1 || targetColIndex === -1) { return; }

  let categorizedCount = 0;
  const targetValues = [];
  for (let i = 1; i < gameData.length; i++) {
    const gameEcoCode = gameData[i][ecoIndex];
    const key = gameEcoCode ? gameEcoCode.toString().trim() : '';
    let val = '';
    if (key && ecoMap.has(key)) {
      val = ecoMap.get(key);
      categorizedCount++;
    }
    targetValues.push([val]);
  }

  if (targetValues.length > 0) {
    gameSheet.getRange(2, targetColIndex + 1, targetValues.length, 1).setValues(targetValues);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('ECO categorization complete! Updated ' + categorizedCount + ' games.', 'Success', 8);
}

// ===== ChessManager.gs (Part 4/5) =====

// =================================================================
// TRIGGER MANAGEMENT
// =================================================================

function setupTriggersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let triggerSheet = ss.getSheetByName(SHEETS.TRIGGERS);
  if (!triggerSheet) {
    triggerSheet = ss.insertSheet(SHEETS.TRIGGERS);
  }
  triggerSheet.clear().setFrozenRows(1);

  const headers = ['Function to Run', 'Frequency', 'Day of the Week (for Weekly)', 'Time of Day (for Daily/Weekly)', 'Status', 'Trigger ID'];
  triggerSheet.getRange('A1:F1').setValues([headers]).setFontWeight('bold');

  const functionRule = SpreadsheetApp.newDataValidation().requireValueInList(TRIGGERABLE_FUNCTIONS, true).build();

  const frequencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
        'Every hour', 'Every 2 hours', 'Every 4 hours', 'Every 6 hours', 'Every 12 hours',
        'Daily', 'Weekly'
    ], true).build();

  const dayOfWeekRule = SpreadsheetApp.newDataValidation().requireValueInList(['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], true).build();

  triggerSheet.getRange('A2:A').setDataValidation(functionRule);
  triggerSheet.getRange('B2:B').setDataValidation(frequencyRule);
  triggerSheet.getRange('C2:C').setDataValidation(dayOfWeekRule);

  triggerSheet.autoResizeColumns(1, 4);
  SpreadsheetApp.getActiveSpreadsheet().toast('Triggers sheet has been set up!', 'Success', 5);
}

function applyTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (TRIGGERABLE_FUNCTIONS.includes(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggerSheet = ss.getSheetByName(SHEETS.TRIGGERS);
  if (!triggerSheet) { throw new Error('Triggers sheet not found.'); }

  const settings = triggerSheet.getRange(2, 1, Math.max(triggerSheet.getLastRow() - 1, 0), 6).getValues();

  settings.forEach((row, index) => {
    const [functionName, frequency, dayOfWeek, timeOfDay] = row;
    if (!functionName || !frequency) {
      triggerSheet.getRange(index + 2, 5, 1, 2).setValues([['Inactive', '']]);
      return;
    }
    try {
      let newTriggerBuilder;

      if (frequency.startsWith('Every')) {
        const hoursMatch = frequency.match(/\d+/);
        const hours = hoursMatch ? parseInt(hoursMatch[0]) : 1;
        newTriggerBuilder = ScriptApp.newTrigger(functionName).timeBased().everyHours(hours);
      }
      else if (frequency === 'Daily' && timeOfDay) {
        const parts = timeOfDay.toString().split(':');
        const hour = parseInt(parts[0], 10);
        const minute = parts.length > 1 ? parseInt(parts[1], 10) : 0;
        newTriggerBuilder = ScriptApp.newTrigger(functionName).timeBased().everyDays(1).atHour(hour).nearMinute(minute);
      }
      else if (frequency === 'Weekly' && dayOfWeek && timeOfDay) {
        const parts = timeOfDay.toString().split(':');
        const hour = parseInt(parts[0], 10);
        const minute = parts.length > 1 ? parseInt(parts[1], 10) : 0;
        const weekDay = ScriptApp.WeekDay[dayOfWeek.toUpperCase()];
        newTriggerBuilder = ScriptApp.newTrigger(functionName).timeBased().onWeekDay(weekDay).atHour(hour).nearMinute(minute);
      } else {
        throw new Error('Invalid settings for frequency: ' + frequency);
      }

      const newTrigger = newTriggerBuilder.create();
      triggerSheet.getRange(index + 2, 5, 1, 2).setValues([['Active', newTrigger.getUniqueId()]]);

    } catch (e) {
      console.error('Error on row ' + (index + 2) + ': ' + e.message);
      triggerSheet.getRange(index + 2, 5).setValue('Error: ' + e.message);
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast('Triggers have been updated successfully!', 'Triggers Synced', 8);
}

function deleteAllTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    ScriptApp.deleteTrigger(trigger);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggerSheet = ss.getSheetByName(SHEETS.TRIGGERS);
  if (triggerSheet && triggerSheet.getLastRow() > 1) {
    triggerSheet.getRange(2, 5, triggerSheet.getLastRow() - 1, 2).clearContent();
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Deleted ' + allTriggers.length + ' triggers.', 'All Triggers Removed', 5);
}

// =================================================================
// FOCUSED RECENT REFRESH
// =================================================================

function refreshRecentData() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { throw new Error('Another run is in progress'); }
  const startTime = new Date().getTime();
  let username = '';
  const results = [];
  try {
    username = getUsername();
    console.log('Starting recent data refresh for user: ' + username);

    try {
      fetchCurrentMonthGames();
      results.push('✓ Current month games fetched');
    } catch (error) {
      results.push('✗ Current month games failed: ' + error.message);
    }

    try {
      updateDailyData();
      results.push('✓ Daily data processed');
    } catch (error) {
      results.push('✗ Daily data failed: ' + error.message);
    }

    const executionTime = new Date().getTime() - startTime;
    logExecution('refreshRecentData', username, 'SUCCESS', executionTime, results.join(', '));
    SpreadsheetApp.getActiveSpreadsheet().toast('Recent data refresh complete!\n\n' + results.join('\n'), 'Refresh Complete', 8);

  } catch (error) {
    const executionTime = new Date().getTime() - startTime;
    logExecution('refreshRecentData', username, 'ERROR', executionTime, error.message);
    throw error;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// =================================================================
// HIGH-LEVEL WORKFLOWS
// =================================================================

function quickUpdate() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { throw new Error('Another run is in progress'); }
  const startMs = new Date().getTime();
  const username = getUsername();
  const results = [];
  try {
    try {
      updateArchives();
      results.push('Archives updated');
    } catch (e) {
      results.push('Archives failed: ' + e.message);
    }
    try {
      fetchCurrentMonthGames();
      results.push('Current month games fetched');
    } catch (e) {
      results.push('Fetch current month failed: ' + e.message);
    }
    try {
      updateDailyData();
      results.push('Daily data updated');
    } catch (e) {
      results.push('Daily update failed: ' + e.message);
    }
    try {
      updateStats();
      results.push('Stats updated');
    } catch (e) {
      results.push('Stats failed: ' + e.message);
    }
    try {
      updateProfile();
      results.push('Profile updated');
    } catch (e) {
      results.push('Profile failed: ' + e.message);
    }
    SpreadsheetApp.getActiveSpreadsheet().toast(results.join('\n'), 'Quick Update', 8);
    logExecution('quickUpdate', username, 'SUCCESS', new Date().getTime() - startMs, results.join(', '));
  } catch (error) {
    logExecution('quickUpdate', username, 'ERROR', new Date().getTime() - startMs, error.message);
    throw error;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function completeUpdate() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { throw new Error('Another run is in progress'); }
  const startMs = new Date().getTime();
  const username = getUsername();
  const results = [];
  try {
    try {
      updateArchives();
      results.push('Archives updated');
    } catch (e) {
      results.push('Archives failed: ' + e.message);
    }
    try {
      fetchAllGames();
      results.push('All games fetched');
    } catch (e) {
      results.push('Fetch all games failed: ' + e.message);
    }
    try {
      updateDailyData();
      results.push('Daily data updated');
    } catch (e) {
      results.push('Daily update failed: ' + e.message);
    }
    try {
      updateStats();
      results.push('Stats updated');
    } catch (e) {
      results.push('Stats failed: ' + e.message);
    }
    try {
      updateProfile();
      results.push('Profile updated');
    } catch (e) {
      results.push('Profile failed: ' + e.message);
    }
    SpreadsheetApp.getActiveSpreadsheet().toast(results.join('\n'), 'Complete Update', 8);
    logExecution('completeUpdate', username, 'SUCCESS', new Date().getTime() - startMs, results.join(', '));
  } catch (error) {
    logExecution('completeUpdate', username, 'ERROR', new Date().getTime() - startMs, error.message);
    throw error;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// ===== ChessManager.gs (Part 5/5) =====

// =================================================================
// DAILY AUTOMATION (COMBINED)
// =================================================================

const DAILY_SHEET_NAME = SHEETS.DAILY;
const GAME_DATA_SHEET_NAME = SHEETS.GAMES;

function installDailyTrigger() {
  removeDailyTriggers();
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  ScriptApp.newTrigger('dailyRoll').timeBased().inTimezone(tz).everyDays(1).atHour(0).create();
}

function removeDailyTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'dailyRoll') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function dailyRoll() {
  ensureGameDataComputedColumns();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const daily = ss.getSheetByName(DAILY_SHEET_NAME);
  if (!daily) {
    throw new Error('Sheet "' + DAILY_SHEET_NAME + '" not found.');
  }

  const tz = ss.getSpreadsheetTimeZone();
  const todayString = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const headerRow = 1, dateCol = 1;

  const lastRow = daily.getLastRow();
  const lastCol = daily.getLastColumn();
  if (lastRow < headerRow + 1) {
    daily.insertRowAfter(headerRow);
    const dateCell = daily.getRange(headerRow + 1, dateCol);
    dateCell.setValue(todayString).setNumberFormat('yyyy-mm-dd');
    SpreadsheetApp.flush();
    return;
  }

  const topDate = daily.getRange(headerRow + 1, dateCol).getDisplayValue();
  if (topDate === todayString) {
    return;
  }

  daily.insertRowAfter(headerRow);
  const dateCell = daily.getRange(headerRow + 1, dateCol);
  dateCell.setValue(todayString).setNumberFormat('yyyy-mm-dd');

  const sourceRowIndex = headerRow + 2;
  if (lastCol > 1) {
    const sourceFormulaRange = daily.getRange(sourceRowIndex, 2, 1, lastCol - 1);
    const formulasR1C1 = sourceFormulaRange.getFormulasR1C1();
    if (formulasR1C1 && formulasR1C1.length > 0) {
      daily.getRange(headerRow + 1, 2, 1, lastCol - 1).setFormulasR1C1(formulasR1C1);
    } else {
      const values = daily.getRange(sourceRowIndex, 2, 1, lastCol - 1).getValues();
      daily.getRange(headerRow + 1, 2, 1, lastCol - 1).setValues(values);
    }
  }

  SpreadsheetApp.flush();

  const newLastRow = daily.getLastRow();
  if (newLastRow >= headerRow + 2) {
    const freezeStartRow = headerRow + 2;
    const numRowsToFreeze = newLastRow - freezeStartRow + 1;
    if (numRowsToFreeze > 0) {
      const range = daily.getRange(freezeStartRow, 1, numRowsToFreeze, lastCol);
      const values = range.getValues();
      range.setValues(values);
    }
  }
}

function ensureGameDataComputedColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GAME_DATA_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + GAME_DATA_SHEET_NAME + '" not found.');
  }

  const headerRow = 1;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, Math.max(1, lastCol)).getValues()[0];

  const getColIndexByHeader = (name) => headers.findIndex(h => String(h).trim().toLowerCase() === name.toLowerCase()) + 1;
  const endTimeCol   = getColIndexByHeader('End Time');
  const winnerCol    = getColIndexByHeader('Winner');
  const myUserCol    = getColIndexByHeader('My Username');
  const timeClassCol = getColIndexByHeader('Time Class');
  const myRatingCol  = getColIndexByHeader('My Rating');

  const dateColIdx      = findOrCreateColumn(sheet, headers, 'Date');
  const resultBinColIdx = findOrCreateColumn(sheet, headers, 'ResultBinary');
  const ratingChgColIdx = findOrCreateColumn(sheet, headers, 'RatingChange');

  if (endTimeCol) {
    const endColLetter = columnToLetter(endTimeCol);
    const dateFormula =
      '=ARRAYFORMULA({"Date"; IF(LEN(' + endColLetter + '2:' + endColLetter + ')=0, , INT(' + endColLetter + '2:' + endColLetter + '))})';
    sheet.getRange(1, dateColIdx).setFormula(dateFormula);
  }

  if (winnerCol && myUserCol) {
    const winL = columnToLetter(winnerCol);
    const meL  = columnToLetter(myUserCol);
    const resultBinFormula =
      '=ARRAYFORMULA({"ResultBinary"; IF(LEN(' + winL + '2:' + winL + ')=0, , IF(LOWER(' + winL + '2:' + winL + ')="draw", 0.5, IF(LOWER(' + winL + '2:' + winL + ')=LOWER(' + meL + '2:' + meL + '), 1, 0)))})';
    sheet.getRange(1, resultBinColIdx).setFormula(resultBinFormula);
  }

  if (timeClassCol && myRatingCol && endTimeCol) {
    const fmtL = columnToLetter(timeClassCol);
    const ratL = columnToLetter(myRatingCol);
    const endL = columnToLetter(endTimeCol);
    const ratingChangeFormula =
      '=RATING_CHANGE(' + fmtL + '2:' + fmtL + ', ' + ratL + '2:' + ratL + ', ' + endL + '2:' + endL + ')';
    sheet.getRange(2, ratingChgColIdx).setFormula(ratingChangeFormula);
  }

  applyNumberFormatsForNewColumns();
}

function RATING_CHANGE(formatCol, ratingCol, endTimeCol) {
  const toFlat = (arr) => {
    if (!Array.isArray(arr)) return [];
    if (Array.isArray(arr[0])) return arr.map(r => r[0]);
    return arr;
  };
  const fmt = toFlat(formatCol);
  const rat = toFlat(ratingCol).map(v => (v === '' || v == null ? null : Number(v)));
  const end = toFlat(endTimeCol).map(v => v instanceof Date ? v : (v ? new Date(v) : null));

  const n = Math.max(fmt.length, rat.length, end.length);
  const rows = [];
  for (let i = 0; i < n; i++) {
    rows.push({
      idx: i,
      format: fmt[i] == null ? '' : String(fmt[i]),
      rating: rat[i] == null || isNaN(rat[i]) ? null : rat[i],
      endTime: end[i] instanceof Date && !isNaN(end[i].getTime()) ? end[i] : null,
    });
  }

  const valid = rows.map((r, i) => ({ ...r, origIndex: i }))
    .filter(r => r.format !== '' && r.rating != null && r.endTime != null);

  valid.sort((a, b) => a.endTime - b.endTime || a.origIndex - b.origIndex);

  const lastByFormat = new Map();
  const deltaByOrigIndex = new Map();
  for (const r of valid) {
    const key = r.format;
    const prev = lastByFormat.get(key);
    const delta = prev == null ? 0 : r.rating - prev;
    deltaByOrigIndex.set(r.origIndex, delta);
    lastByFormat.set(key, r.rating);
  }

  const out = new Array(n).fill('');
  for (let i = 0; i < n; i++) {
    if (deltaByOrigIndex.has(i)) {
      out[i] = deltaByOrigIndex.get(i);
    } else {
      out[i] = '';
    }
  }
  return out.map(v => [v]);
}

// =================================================================
// FORMATTING HELPERS
// =================================================================

function applyNumberFormatsForNewColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.GAMES);
  if (!sheet || sheet.getLastRow() < 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const namesDateOnly = ['UTC Start Date','UTC End Date','Local Start Date','Local End Date','Date','PGN End Date','UTC Date'];
  const namesTimeOnly = ['UTC Start Time','UTC End Time','Local Start Time','Local End Time','PGN Start Time','PGN End Time','UTC Time'];
  const namesOffset   = ['Local − UTC Offset (hrs)','Hour Differential (hrs)'];

  const dateFormat = 'yyyy-mm-dd';
  const timeFormat = 'hh:mm:ss';
  const numFormat  = '0.00';

  function setFormatByNames(names, fmt) {
    names.forEach(n => {
      const idx = headers.indexOf(n);
      if (idx !== -1) {
        const col = idx + 1;
        const numRows = Math.max(sheet.getLastRow() - 1, 0);
        if (numRows > 0) {
          sheet.getRange(2, col, numRows, 1).setNumberFormat(fmt);
        }
      }
    });
  }

  setFormatByNames(namesDateOnly, dateFormat);
  setFormatByNames(namesTimeOnly, timeFormat);
  setFormatByNames(namesOffset, numFormat);
}

// =================================================================
// HELPERS
// =================================================================

function findOrCreateColumn(sheet, headersRowValues, headerName) {
  const idx = headersRowValues.findIndex(h => String(h).trim().toLowerCase() === headerName.toLowerCase());
  if (idx >= 0) {
    return idx + 1;
  }
  const lastCol = sheet.getLastColumn();
  const insertAt = lastCol + 1;
  sheet.getRange(1, insertAt).setValue(headerName);
  return insertAt;
}

function columnToLetter(column) {
  let temp = '';
  let col = column;
  while (col > 0) {
    let rem = (col - 1) % 26;
    temp = String.fromCharCode(rem + 65) + temp;
    col = Math.floor((col - rem) / 26);
  }
  return temp;
}
