// Minimal Chess.com monthly archive importer (PGN headers only)
// - Creates `Config` and `Games Raw` sheets
// - Fetches a specified month archive for a username
// - Parses PGN headers and writes one row per game using union of all headers

const SHEET_NAMES = {
  CONFIG: 'Config',
  RAW: 'Games Raw'
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Archive Import')
    .addItem('Setup Sheets', 'setupSheets')
    .addItem('Fetch Month from Config', 'fetchMonthFromConfig')
    .addToUi();
}

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config sheet with minimal settings
  let config = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!config) config = ss.insertSheet(SHEET_NAMES.CONFIG);
  const configHeaders = ['Setting', 'Value'];
  config.clear();
  config.getRange(1, 1, 1, configHeaders.length)
    .setValues([configHeaders])
    .setFontWeight('bold');
  config.setFrozenRows(1);

  const today = new Date();
  const defaultRows = [
    ['username', ''],
    ['year', String(today.getUTCFullYear())],
    ['month', String(today.getUTCMonth() + 1).padStart(2, '0')]
  ];
  config.getRange(2, 1, defaultRows.length, defaultRows[0].length).setValues(defaultRows);

  // Raw output sheet
  let raw = ss.getSheetByName(SHEET_NAMES.RAW);
  if (!raw) raw = ss.insertSheet(SHEET_NAMES.RAW);
  raw.clear();
  raw.setFrozenRows(1);

  SpreadsheetApp.getActiveSpreadsheet().toast('Sheets ready. Fill Config → username/year/month.', 'Setup Complete', 6);
}

function fetchMonthFromConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(SHEET_NAMES.CONFIG);
  if (!config) {
    throw new Error('Config sheet missing. Run "Setup Sheets" first.');
  }

  const cfg = readConfig(config);
  const username = String(cfg.username || '').trim().toLowerCase();
  const year = String(cfg.year || '').trim();
  const month = String(cfg.month || '').trim().padStart(2, '0');
  if (!username) throw new Error('Config username is required.');
  if (!/^[0-9]{4}$/.test(year)) throw new Error('Config year must be YYYY.');
  if (!/^[0-9]{2}$/.test(month)) throw new Error('Config month must be MM.');

  const url = 'https://api.chess.com/pub/player/' + encodeURIComponent(username) + '/games/' + year + '/' + month;
  const json = fetchJson(url);
  const games = (json && json.games) ? json.games : [];

  const headerUnion = new Set();
  const gameHeaderMaps = [];
  for (let i = 0; i < games.length; i++) {
    const pgn = games[i] && games[i].pgn ? String(games[i].pgn) : '';
    const map = parsePgnHeadersToMap(pgn);
    gameHeaderMaps.push(map);
    Object.keys(map).forEach(k => headerUnion.add(k));
  }

  const headers = Array.from(headerUnion);
  headers.sort();

  const raw = ss.getSheetByName(SHEET_NAMES.RAW) || ss.insertSheet(SHEET_NAMES.RAW);
  raw.clear();
  if (headers.length > 0) {
    raw.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    const values = new Array(gameHeaderMaps.length);
    for (let r = 0; r < gameHeaderMaps.length; r++) {
      const m = gameHeaderMaps[r];
      const row = new Array(headers.length);
      for (let c = 0; c < headers.length; c++) {
        const key = headers[c];
        row[c] = m.hasOwnProperty(key) ? m[key] : '';
      }
      values[r] = row;
    }
    if (values.length > 0) {
      raw.getRange(2, 1, values.length, headers.length).setValues(values);
    }
  } else {
    // No PGN headers found (no games or empty PGNs); leave sheet empty with just frozen header row
    raw.setFrozenRows(1);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Imported games: ' + games.length, 'Done', 6);
}

function readConfig(configSheet) {
  const numRows = configSheet.getLastRow();
  const values = numRows > 1 ? configSheet.getRange(2, 1, numRows - 1, 2).getValues() : [];
  const out = {};
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    const val = values[i][1];
    if (key) out[key] = val;
  }
  return out;
}

function fetchJson(url) {
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  if (code >= 200 && code < 300) {
    const text = resp.getContentText();
    return text ? JSON.parse(text) : {};
  }
  throw new Error('HTTP ' + code + ' for ' + url);
}

function parsePgnHeadersToMap(pgn) {
  const map = {};
  if (!pgn) return map;
  const headerEnd = pgn.indexOf('\n\n');
  const headerText = headerEnd === -1 ? pgn : pgn.slice(0, headerEnd);
  const re = /^\s*\(([^")]+\)|\[([A-Za-z0-9_]+)\s+"([^\"]*)"\s*\])\s*$/gm; // not used branch kept minimal
  // Simpler pass just for classic [Key "Value"] lines
  const simpleRe = /^\s*\[([A-Za-z0-9_]+)\s+"([^\"]*)"\s*\]\s*$/gm;
  let m;
  while ((m = simpleRe.exec(headerText)) !== null) {
    const key = m[1];
    const val = m[2];
    map[key] = val;
  }
  return map;
}

