// ============================================================
//  Google Sheets Power Query  –  Code.gs
//  Entry point: menu, sidebar, server-side bridge functions
// ============================================================

// ── Menu setup ──────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Power Query')
    .addItem('Open Power Query (full window)', 'openFullWindow')
    .addItem('Open Power Query (panel)', 'openSidebar')
    .addSeparator()
    .addItem('Run saved query…', 'runSavedQueryPrompt')
    .addItem('Clear all saved queries', 'clearAllQueries')
    .addToUi();
}

/**
 * Opens Power Query as a full browser tab using the deployed web app URL.
 * Requires the script to be deployed as a Web App (Execute as: Me, Who: Anyone with Google account).
 */
function openFullWindow() {
  // PASTE YOUR COPIED URL BETWEEN THE QUOTES BELOW
  var manualUrl = "https://script.google.com/macros/s/AKfycbx4dA18X3ZlxAjfNse7ynC-I7nyWacCBV77jzyBPppkO4EalyFaPENE8S4X7FtQF_-tRg/exec";

  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var url = manualUrl + '?ssId=' + ssId;

  var html = HtmlService.createHtmlOutput(
    '<style>' +
    '  body { font-family: sans-serif; font-size: 13px; padding: 18px 20px; ' +
    '         background:#0f0f17; color:#e2e2f0; margin:0; }' +
    '  a { display:block; margin-top:12px; padding:10px 16px; background:#7c6af7; ' +
    '      color:#fff; text-decoration:none; border-radius:7px; text-align:center; ' +
    '      font-weight:600; font-size:13px; }' +
    '  a:hover { background:#9d8fff; }' +
    '  p { color:#9090b0; font-size:12px; margin-top:10px; text-align:center; }' +
    '</style>' +
    '<div>Click below to open Power Query in a new tab:</div>' +
    '<a href="' + url + '" target="_blank" onclick="google.script.host.close()">⚡ Open Power Query</a>' +
    '<p>You can close this dialog after clicking.</p>'
  ).setWidth(320).setHeight(140);

  SpreadsheetApp.getUi().showModalDialog(html, '⚡ Open Power Query');
}

/**
 * Fallback: open as a wide modeless dialog (no deployment needed).
 */
function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(620)
    .setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, '⚡ Power Query');
}

/**
 * Web App entry point — serves the full-page Power Query UI.
 * Binds to the spreadsheet passed via ?ssId= query param.
 */
function doGet(e) {
  var ssId = e && e.parameter && e.parameter.ssId;
  if (ssId) {
    try { SpreadsheetApp.openById(ssId); } catch(err) { /* ignore, best-effort bind */ }
  }
  return HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('⚡ Power Query')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ── Data helpers called from sidebar ────────────────────────

/** Returns all sheet names in the active spreadsheet */
function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(function(s) { return s.getName(); });
}

/** Returns the raw 2-D array for a given sheet (first 1000 rows for speed) */
function getSheetData(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  var lastRow  = Math.min(sheet.getLastRow(),  1000);
  var lastCol  = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];

  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

/** Returns just the header row of a sheet */
function getSheetHeaders(sheetName) {
  var data = getSheetData(sheetName);
  return data.length > 0 ? data[0].map(String) : [];
}


// ── Core run function ────────────────────────────────────────

/**
 * Accepts a query definition object from the sidebar, runs the
 * Engine pipeline, and writes results to a new sheet.
 *
 * queryDef shape:
 * {
 *   sourceName: 'Sheet1',
 *   headerRows: 1,          // 1 or 2
 *   steps: [ { type, params } … ],
 *   outputName: 'PQ_Result',
 *   replaceExisting: true
 * }
 */
function runQuery(queryDef) {
  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var raw    = getSheetData(queryDef.sourceName);
    var result = Engine.run(raw, queryDef.headerRows || 1, queryDef.steps || []);

    var outName = queryDef.outputName ||
      ('PQ_' + queryDef.sourceName + '_' + _timestamp());

    // Remove existing sheet with same name if requested
    var existing = ss.getSheetByName(outName);
    if (existing) {
      if (queryDef.replaceExisting) {
        ss.deleteSheet(existing);
      } else {
        outName = outName + '_' + _timestamp();
      }
    }

    var outSheet = ss.insertSheet(outName);
    if (result.length > 0) {
      outSheet.getRange(1, 1, result.length, result[0].length)
              .setValues(result);

      // Style the header row
      outSheet.getRange(1, 1, 1, result[0].length)
              .setBackground('#1a1a2e')
              .setFontColor('#e0e0ff')
              .setFontWeight('bold');

      // Auto-resize columns
      for (var c = 1; c <= result[0].length; c++) {
        outSheet.autoResizeColumn(c);
      }
    }

    ss.setActiveSheet(outSheet);
    return { success: true, outputSheet: outName, rows: result.length - 1 };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


/**
 * Preview-only: runs the pipeline and returns the first 50 rows.
 * Does NOT write to any sheet.
 */
function previewQuery(queryDef) {
  try {
    var raw    = getSheetData(queryDef.sourceName);
    var result = Engine.run(raw, queryDef.headerRows || 1, queryDef.steps || []);
    var preview = result.slice(0, 51); // header + 50 data rows
    return { success: true, data: preview };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// ── Query persistence ────────────────────────────────────────

var PROP_KEY = 'pq_saved_queries';

function saveQuery(name, queryDef) {
  try {
    var store = _loadStore();
    store[name] = queryDef;
    PropertiesService.getDocumentProperties()
      .setProperty(PROP_KEY, JSON.stringify(store));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function loadQuery(name) {
  var store = _loadStore();
  return store[name] || null;
}

function listSavedQueries() {
  return Object.keys(_loadStore());
}

function deleteQuery(name) {
  var store = _loadStore();
  delete store[name];
  PropertiesService.getDocumentProperties()
    .setProperty(PROP_KEY, JSON.stringify(store));
  return { success: true };
}

function clearAllQueries() {
  PropertiesService.getDocumentProperties().deleteProperty(PROP_KEY);
  SpreadsheetApp.getUi().alert('All saved queries cleared.');
}

function runSavedQueryPrompt() {
  var names = listSavedQueries();
  if (names.length === 0) {
    SpreadsheetApp.getUi().alert('No saved queries found. Open Power Query to create one.');
    return;
  }
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Run saved query',
    'Available: ' + names.join(', ') + '\n\nEnter query name:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() === ui.Button.OK) {
    var qd = loadQuery(resp.getResponseText().trim());
    if (!qd) { ui.alert('Query not found.'); return; }
    var res = runQuery(qd);
    if (res.success) {
      ui.alert('Done! Written to sheet "' + res.outputSheet + '" (' + res.rows + ' rows).');
    } else {
      ui.alert('Error: ' + res.error);
    }
  }
}


// ── Private helpers ──────────────────────────────────────────

function _loadStore() {
  var raw = PropertiesService.getDocumentProperties().getProperty(PROP_KEY);
  return raw ? JSON.parse(raw) : {};
}

function _timestamp() {
  var d = new Date();
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMdd_HHmm');
}