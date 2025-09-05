/**
 * Apps Script functions to generate/apply header configurations for Raw and Calculated sheets.
 */

function GenerateRawHeadersConfig() {
  const dataSheet = getSheetByNameOrThrow_(HEADER_CONFIG.rawDataSheetName);
  const configSheet = ensureConfigSheet_(HEADER_CONFIG.rawConfigSheetName);
  clearAndPrepareConfigSheet_(configSheet);

  const headers = readHeaderRow_(dataSheet);
  const examples = gatherExamplesForHeaders_(dataSheet, headers, 3);
  const defaults = loadDefaults_(HEADER_CONFIG.propertiesKeyRaw);
  const rows = buildConfigRowsForHeaders_(headers, examples, defaults);
  writeConfigRows_(configSheet, rows);
  paintCategoryColors_(configSheet);
}

function ApplyRawHeadersConfig() {
  const dataSheet = getSheetByNameOrThrow_(HEADER_CONFIG.rawDataSheetName);
  const configSheet = getSheetByNameOrThrow_(HEADER_CONFIG.rawConfigSheetName);
  const cfg = readConfig_(configSheet);
  if (!cfg.order.length) throw new Error('Config has no rows. Run GenerateRawHeadersConfig first.');
  applyColumnOrderAndVisibility_(dataSheet, cfg.order);
  paintHeaderCategoryColors_(dataSheet, cfg.category, cfg.colors);
  groupColumnsByCategory_(dataSheet, cfg.category);
}

function SaveRawHeadersDefaults() {
  const configSheet = getSheetByNameOrThrow_(HEADER_CONFIG.rawConfigSheetName);
  const cfg = readConfig_(configSheet);
  saveDefaults_(HEADER_CONFIG.propertiesKeyRaw, cfg);
}

function GenerateCalcHeadersConfig() {
  const dataSheet = getSheetByNameOrThrow_(HEADER_CONFIG.calcDataSheetName);
  const configSheet = ensureConfigSheet_(HEADER_CONFIG.calcConfigSheetName);
  clearAndPrepareConfigSheet_(configSheet);

  const headers = readHeaderRow_(dataSheet);
  const examples = gatherExamplesForHeaders_(dataSheet, headers, 3);
  const defaults = loadDefaults_(HEADER_CONFIG.propertiesKeyCalc);
  const rows = buildConfigRowsForHeaders_(headers, examples, defaults);
  writeConfigRows_(configSheet, rows);
  paintCategoryColors_(configSheet);
}

function ApplyCalcHeadersConfig() {
  const dataSheet = getSheetByNameOrThrow_(HEADER_CONFIG.calcDataSheetName);
  const configSheet = getSheetByNameOrThrow_(HEADER_CONFIG.calcConfigSheetName);
  const cfg = readConfig_(configSheet);
  if (!cfg.order.length) throw new Error('Config has no rows. Run GenerateCalcHeadersConfig first.');
  applyColumnOrderAndVisibility_(dataSheet, cfg.order);
  paintHeaderCategoryColors_(dataSheet, cfg.category, cfg.colors);
  groupColumnsByCategory_(dataSheet, cfg.category);
}

function SaveCalcHeadersDefaults() {
  const configSheet = getSheetByNameOrThrow_(HEADER_CONFIG.calcConfigSheetName);
  const cfg = readConfig_(configSheet);
  saveDefaults_(HEADER_CONFIG.propertiesKeyCalc, cfg);
}

function AddHeaderConfigMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Chess Headers')
    .addItem('Generate Raw Header Config', 'GenerateRawHeadersConfig')
    .addItem('Apply Raw Header Config', 'ApplyRawHeadersConfig')
    .addItem('Save Raw Defaults', 'SaveRawHeadersDefaults')
    .addSeparator()
    .addItem('Generate Calc Header Config', 'GenerateCalcHeadersConfig')
    .addItem('Apply Calc Header Config', 'ApplyCalcHeadersConfig')
    .addItem('Save Calc Defaults', 'SaveCalcHeadersDefaults')
    .addSeparator()
    .addItem('Generate Calculated Formulas Sheet', 'GenerateCalculatedDataFormulas')
    .addItem('Interactive: Add Calculated Header (AI)', 'InteractiveAddCalculatedHeaderWithAI')
    .addToUi();
}

// onOpen is merged into the main menu in code.gs

