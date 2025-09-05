/**
 * Shared helpers for Sheets header management, ordering, categories, and defaults.
 * Google Apps Script compatible (V8 runtime).
 */

const HEADER_CONFIG = {
  rawConfigSheetName: 'Config - Raw Headers',
  calcConfigSheetName: 'Config - Calc Headers',
  rawDataSheetName: 'Raw',
  calcDataSheetName: 'Calculated',
  propertiesKeyRaw: 'DEFAULT_HEADERS_RAW',
  propertiesKeyCalc: 'DEFAULT_HEADERS_CALC'
};

const CATEGORY_COLORS = {
  Meta: '#e8f0fe',
  Players: '#e6f4ea',
  Ratings: '#fff7e0',
  Time: '#fde7e9',
  PGN: '#f1f3f4',
  Accuracy: '#e1f5fe',
  Result: '#fce8b2',
  Variant: '#f3e8fd',
  Derived: '#fff0f0',
  Opponent: '#e0f7fa'
};

function getSpreadsheet_() {
  return SpreadsheetApp.getActive();
}

function getSheetByNameOrThrow_(name) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Missing sheet: ${name}`);
  return sh;
}

function trim_(s) {
  return typeof s === 'string' ? s.trim() : s;
}

function readHeaderRow_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  const values = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return values.map(v => trim_(v)).filter(v => v && String(v).length);
}

function gatherExamplesForHeaders_(sheet, headers, maxSamples) {
  const out = {};
  const headerIndex = new Map(headers.map((h, i) => [h, i + 1]));
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return out;
  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = range.getValues();
  for (const h of headers) {
    const col = headerIndex.get(h);
    if (!col) continue;
    const examples = [];
    for (let r = 0; r < values.length && examples.length < maxSamples; r++) {
      const cell = values[r][col - 1];
      if (cell !== '' && cell !== null && cell !== undefined) {
        examples.push(String(cell));
      }
    }
    if (examples.length) out[h] = examples.join(' | ');
  }
  return out;
}

function categoryGuess_(name) {
  const n = String(name).toLowerCase();
  if (/(username|white|black|opponent)/.test(n)) return 'Players';
  if (/(rating|elo|expected|accuracy)/.test(n)) return 'Ratings';
  if (/(time_control|time-class|time_class|end_time|utc|clk|time|delay|increment|days)/.test(n)) return 'Time';
  if (/(pgn|san|fen|move|castl|promotion|check|mate)/.test(n)) return 'PGN';
  if (/(result|termination|winner|score|draw)/.test(n)) return 'Result';
  if (/(variant|mode|daily960|chess960|rules)/.test(n)) return 'Variant';
  if (/(opponent)/.test(n)) return 'Opponent';
  return 'Meta';
}

function colorForCategory_(cat) {
  return CATEGORY_COLORS[cat] || '#ffffff';
}

function ensureConfigSheet_(sheetName) {
  const ss = getSpreadsheet_();
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function clearAndPrepareConfigSheet_(sheet) {
  sheet.clear({ contentsOnly: true });
  const headers = ['Field', 'Description', 'Example', 'Category', 'Order', 'Hidden', 'ColorHex', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

function writeConfigRows_(sheet, rows) {
  if (!rows.length) return;
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  // Autofit
  sheet.autoResizeColumns(1, rows[0].length);
}

function paintCategoryColors_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const catCol = 4; // Category
  const colorCol = 7; // ColorHex
  const cats = sheet.getRange(2, catCol, lastRow - 1, 1).getValues().map(r => r[0]);
  const colors = sheet.getRange(2, colorCol, lastRow - 1, 1).getValues().map(r => r[0]);
  const bg = [];
  for (let i = 0; i < cats.length; i++) {
    const hex = String(colors[i] || colorForCategory_(cats[i] || 'Meta'));
    bg.push([hex]);
  }
  sheet.getRange(2, colorCol, lastRow - 1, 1).setValues(bg);
}

function loadDefaults_(key) {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(key);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function saveDefaults_(key, obj) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(key, JSON.stringify(obj));
}

function orderFromDefaultsOrAlpha_(headers, defaults) {
  if (!defaults || !Array.isArray(defaults.order)) return headers.slice().sort((a,b)=>a.localeCompare(b));
  const desired = [];
  const defaultSet = new Set(defaults.order);
  for (const name of defaults.order) if (headers.includes(name)) desired.push(name);
  const remaining = headers.filter(h => !defaultSet.has(h)).sort((a,b)=>a.localeCompare(b));
  return desired.concat(remaining);
}

function hiddenFromDefaults_(headers, defaults) {
  const map = new Map();
  if (defaults && defaults.hidden && typeof defaults.hidden === 'object') {
    for (const [k, v] of Object.entries(defaults.hidden)) map.set(k, Boolean(v));
  }
  const out = {};
  headers.forEach(h => out[h] = map.has(h) ? map.get(h) : false);
  return out;
}

function buildConfigRowsForHeaders_(headers, examples, defaults) {
  const desiredOrder = orderFromDefaultsOrAlpha_(headers, defaults);
  const hiddenByDefault = hiddenFromDefaults_(headers, defaults);
  const rows = [];
  let index = 1;
  for (const h of desiredOrder) {
    const category = (defaults && defaults.category && defaults.category[h]) || categoryGuess_(h);
    const color = (defaults && defaults.colors && defaults.colors[h]) || colorForCategory_(category);
    rows.push([
      h,
      '',
      examples[h] || '',
      category,
      index,
      hiddenByDefault[h] ? true : false,
      color,
      ''
    ]);
    index += 1;
  }
  return rows;
}

function readConfig_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { order: [], hidden: {}, category: {}, colors: {} };
  const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const order = [];
  const hidden = {};
  const category = {};
  const colors = {};
  values.sort((a,b) => Number(a[4-1]||99999) - Number(b[4-1]||99999));
  for (const row of values) {
    const name = trim_(row[0]);
    if (!name) continue;
    order.push(name);
    hidden[name] = Boolean(row[5-1]);
    category[name] = trim_(row[4-1]) || categoryGuess_(name);
    colors[name] = trim_(row[7-1]) || colorForCategory_(category[name]);
  }
  return { order, hidden, category, colors };
}

function applyColumnOrderAndVisibility_(targetSheet, desiredOrder) {
  // Reorder columns in-place using moveColumns. If not available, this will throw.
  const currentHeaders = readHeaderRow_(targetSheet);
  const indexMap = new Map(currentHeaders.map((h,i)=>[h, i+1]));
  let dest = 1;
  for (const name of desiredOrder) {
    const currentIndex = indexMap.get(name);
    if (!currentIndex) continue;
    if (currentIndex !== dest) {
      targetSheet.moveColumns(targetSheet.getRange(1, currentIndex, targetSheet.getMaxRows(), 1), dest);
      // After moving, rebuild the index map
      const refreshed = readHeaderRow_(targetSheet);
      for (let i=0;i<refreshed.length;i++) indexMap.set(refreshed[i], i+1);
    }
    dest += 1;
  }
}

function applyHidden_(targetSheet, hiddenMap) {
  const headers = readHeaderRow_(targetSheet);
  for (let i=0;i<headers.length;i++) {
    const name = headers[i];
    const col = i+1;
    const shouldHide = !!hiddenMap[name];
    if (shouldHide) targetSheet.hideColumns(col);
    else targetSheet.showColumns(col);
  }
}

function paintHeaderCategoryColors_(targetSheet, categoryMap, colorsMap) {
  const headers = readHeaderRow_(targetSheet);
  const bg = headers.map(h => [colorsMap[h] || colorForCategory_(categoryMap[h] || categoryGuess_(h))]);
  if (headers.length) targetSheet.getRange(1, 1, 1, headers.length).setBackgrounds([bg.map(b=>b[0])]);
}

function groupColumnsByCategory_(targetSheet, categoryMap) {
  const headers = readHeaderRow_(targetSheet);
  if (!headers.length) return;
  const lastCol = headers.length;
  // Clear existing groups by reducing depth generously
  try {
    targetSheet.getRange(1, 1, targetSheet.getMaxRows(), lastCol).shiftColumnGroupDepth(-8);
  } catch (e) {
    // Not all domains support grouping; ignore
  }
  let start = 1;
  let currentCat = categoryMap[headers[0]] || categoryGuess_(headers[0]);
  for (let i = 2; i <= lastCol + 1; i++) {
    const name = i <= lastCol ? headers[i - 1] : null;
    const cat = name ? (categoryMap[name] || categoryGuess_(name)) : null;
    if (cat !== currentCat) {
      const width = i - start;
      if (width > 1) {
        try {
          targetSheet.getRange(1, start, targetSheet.getMaxRows(), width).shiftColumnGroupDepth(1);
        } catch (e) {}
      }
      start = i;
      currentCat = cat;
    }
  }
}

