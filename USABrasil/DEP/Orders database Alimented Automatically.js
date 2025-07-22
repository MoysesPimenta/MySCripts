/***************************************************************************************
 * DEP Automation Script - Orders Database Sync
 * Key: (Order ID, Invoice #)  |  REQUIRE_ORDER_ID = true
 * Last updated: 2025-07-17
 ***************************************************************************************/

/** ---------------------------------------------------------------------------
 * CONFIG (namespaced to avoid global collisions)
 * ------------------------------------------------------------------------- */
const DEP_CFG = {
  SPREADSHEET_ID: '', // fill if you want explicit; else active spreadsheet
  SOURCE_SHEET_NAME: 'DEP Data',
  TARGET_SHEET_NAME: 'Orders Database',
  FULL_SYNC_ON_EDIT: true,   // full sync when tracked cols edited
  REQUIRE_ORDER_ID: true,    // skip row if Order ID blank
  TIMEZONE: 'America/Sao_Paulo',
  TIMESTAMP_FMT: "yyyy-MM-dd'T'HH:mm:ss"
};

/** Canonical column names */
const CANON = {
  ORDER_ID: 'Order ID',
  INVOICE: 'Invoice #',
  ORDER_DATE: 'Hardware Order Date',
  SHIP_DATE: 'Hardware Ship Date',
  RESELLER_PO: 'Reseller PO',
  LAST_SYNCED: 'Last Synced'
};

/** Target header order */
const TARGET_HEADERS = [
  CANON.ORDER_ID,
  CANON.INVOICE,
  CANON.ORDER_DATE,
  CANON.SHIP_DATE,
  CANON.RESELLER_PO,
  CANON.LAST_SYNCED
];

/** Columns watched onEdit */
const TRACKED_SOURCE_COLUMNS = [
  CANON.ORDER_ID,
  CANON.INVOICE,
  CANON.ORDER_DATE,
  CANON.SHIP_DATE,
  CANON.RESELLER_PO
];

/** ===========================================================================
 * MENUS & TRIGGERS
 * ======================================================================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DEP Automation')
    .addItem('Sync Orders Database', 'syncOrdersDatabase')
    .addToUi();
}

/**
 * Simple trigger wrapper (attach installable "On edit" trigger in UI).
 */
function onEdit(e) {
  depOnEdit_(e);
}

/**
 * Run full sync if tracked columns in DEP Data edited.
 */
function depOnEdit_(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== DEP_CFG.SOURCE_SHEET_NAME) return;

    // Header change? full sync.
    if (e.range.getRow() === 1) {
      log_('onEdit header change -> full sync');
      syncOrdersDatabase();
      return;
    }

    const colMap = buildColumnMapFromSheet_(sheet, false);
    const trackedCols = trackedSourceColIndexes_(colMap);
    const cStart = e.range.getColumn();
    const cEnd = cStart + e.range.getNumColumns() - 1;
    const intersects = trackedCols.some(c => c >= cStart && c <= cEnd);

    if (!intersects) return;

    if (DEP_CFG.FULL_SYNC_ON_EDIT) {
      log_('onEdit tracked edit -> full sync');
      syncOrdersDatabase();
    } else {
      incrementalSyncRow_(sheet, e.range.getRow(), colMap); // optional
    }
  } catch (err) {
    log_('depOnEdit_ error: %s', err.stack || err);
  }
}

/** ===========================================================================
 * MAIN SYNC
 * ======================================================================== */
function syncOrdersDatabase() {
  const start = new Date();
  const ss = getSpreadsheet_();
  const src = ss.getSheetByName(DEP_CFG.SOURCE_SHEET_NAME);
  if (!src) throw new Error('Source sheet "' + DEP_CFG.SOURCE_SHEET_NAME + '" not found.');
  const tgt = getSheetByNameOrCreate_(ss, DEP_CFG.TARGET_SHEET_NAME, TARGET_HEADERS);

  const srcColMap = buildColumnMapFromSheet_(src, true); // throw if missing required
  const srcRows = readDataRows_(src, srcColMap);
  const tgtData = readTargetRows_(tgt);

  const result = buildUpsertedRecords_(srcRows, tgtData.mapByKey, tgtData.invoiceToOrderMap);
  writeOrdersDatabaseFull_(tgt, result.recordsArray);

  const elapsedMs = new Date() - start;
  const summary = [
    'Orders Sync:',
    'Scanned=' + result.metrics.scannedRows,
    'Eligible=' + result.metrics.eligibleRows,
    'Inserted=' + result.metrics.inserted,
    'Updated=' + result.metrics.updated,
    'SkippedMissing=' + result.metrics.skippedMissingFields,
    'SkippedInvalidDate=' + result.metrics.skippedInvalidDate,
    'InvoiceConflicts=' + result.metrics.invoiceConflicts,
    'TotalWritten=' + result.recordsArray.length,
    'ms=' + elapsedMs
  ].join(' ');
  log_(summary);
  ss.toast(summary, 'Orders Sync', 5);
}

/** ===========================================================================
 * OPTIONAL INCREMENTAL SYNC (if DEP_CFG.FULL_SYNC_ON_EDIT = false)
 * ======================================================================== */
function incrementalSyncRow_(srcSheet, row, srcColMap) {
  const ss = srcSheet.getParent();
  const tgt = getSheetByNameOrCreate_(ss, DEP_CFG.TARGET_SHEET_NAME, TARGET_HEADERS);
  const tgtData = readTargetRows_(tgt);

  const lastCol = srcSheet.getLastColumn();
  const arr = srcSheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const rowObj = rowArrayToObj_(arr, srcColMap);
  rowObj.__rowNum = row;

  const metrics = initMetrics_();
  metrics.scannedRows = 1;

  const rec = buildRecordFromSourceRow_(rowObj, metrics);
  if (!rec) {
    log_('Incremental row %s not eligible', row);
    return;
  }
  metrics.eligibleRows = 1;

  const invKey = normalizeKey_(rec.invoice);
  const ordKey = normalizeKey_(rec.orderId);
  const existingOrderForInvoice = tgtData.invoiceToOrderMap.get(invKey);
  if (existingOrderForInvoice && existingOrderForInvoice !== ordKey) {
    metrics.invoiceConflicts++;
    log_('Incremental conflict: invoice %s already under order %s', rec.invoice, existingOrderForInvoice);
    return;
  }

  const key = compositeKey_(rec.orderId, rec.invoice);
  const existing = tgtData.mapByKey.get(key);
  if (!existing) {
    tgtData.mapByKey.set(key, rec);
    metrics.inserted++;
  } else {
    if (mergeRecordIntoExisting_(existing, rec)) metrics.updated++;
  }
  if (invKey) tgtData.invoiceToOrderMap.set(invKey, ordKey);

  const out = mapToSortedArray_(tgtData.mapByKey);
  writeOrdersDatabaseFull_(tgt, out);
  log_('Incremental sync row %s done: ins=%s upd=%s', row, metrics.inserted, metrics.updated);
}

/** ===========================================================================
 * SOURCE READ
 * ======================================================================== */
function buildColumnMapFromSheet_(sheet, throwIfMissing) {
  const headers = getHeaderValues_(sheet);
  const map = buildColumnMap_(headers);
  if (throwIfMissing) {
    const required = [CANON.ORDER_ID, CANON.INVOICE, CANON.ORDER_DATE, CANON.SHIP_DATE];
    const missing = required.filter(k => !map[k]);
    if (missing.length) {
      throw new Error('Missing required columns in "' + sheet.getName() + '": ' + missing.join(', '));
    }
  }
  return map;
}

function readDataRows_(sheet, colMap) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const arr = values[i];
    const obj = rowArrayToObj_(arr, colMap);
    obj.__rowNum = i + 2;
    rows.push(obj);
  }
  return rows;
}

function rowArrayToObj_(arr, colMap) {
  const o = {};
  Object.keys(colMap).forEach(k => {
    const idx = colMap[k];
    if (!idx) return;
    o[k] = arr[idx - 1];
  });
  return o;
}

/** ===========================================================================
 * TARGET READ
 * ======================================================================== */
function readTargetRows_(tgt) {
  const out = {
    mapByKey: new Map(),
    invoiceToOrderMap: new Map()
  };
  const lastRow = tgt.getLastRow();
  if (lastRow < 2) return out;

  const values = tgt.getRange(2, 1, lastRow - 1, TARGET_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rec = {
      orderId: normalizeString_(row[0]),
      invoice: normalizeString_(row[1]),
      orderDate: row[2] instanceof Date ? row[2] : normalizeDate_(row[2]),
      shipDate: row[3] instanceof Date ? row[3] : normalizeDate_(row[3]),
      resellerPO: normalizeString_(row[4] === undefined ? '' : row[4])
    };
    const key = compositeKey_(rec.orderId, rec.invoice);
    out.mapByKey.set(key, rec);

    const invKey = normalizeKey_(rec.invoice);
    const ordKey = normalizeKey_(rec.orderId);
    if (invKey && !out.invoiceToOrderMap.has(invKey)) {
      out.invoiceToOrderMap.set(invKey, ordKey);
    }
  }
  return out;
}

/** ===========================================================================
 * BUILD / UPSERT
 * ======================================================================== */
function buildUpsertedRecords_(srcRows, tgtMap, invoiceToOrderMap) {
  const metrics = initMetrics_();
  const workMap = new Map();
  tgtMap.forEach((rec, key) => workMap.set(key, Object.assign({}, rec)));
  const invMap = new Map(invoiceToOrderMap);

  for (let i = 0; i < srcRows.length; i++) {
    metrics.scannedRows++;
    const rowObj = srcRows[i];
    const rec = buildRecordFromSourceRow_(rowObj, metrics);
    if (!rec) continue;
    metrics.eligibleRows++;

    const invKey = normalizeKey_(rec.invoice);
    const ordKey = normalizeKey_(rec.orderId);
    const existingOrderForInvoice = invMap.get(invKey);
    if (existingOrderForInvoice && existingOrderForInvoice !== ordKey) {
      metrics.invoiceConflicts++;
      log_(
        'Conflict: Invoice "%s" row %s tied to Order "%s" but already mapped to "%s"; skipping.',
        rec.invoice,
        rowObj.__rowNum,
        rec.orderId,
        existingOrderForInvoice
      );
      continue;
    }

    const key = compositeKey_(rec.orderId, rec.invoice);
    const existing = workMap.get(key);
    if (!existing) {
      workMap.set(key, rec);
      metrics.inserted++;
    } else {
      if (mergeRecordIntoExisting_(existing, rec)) metrics.updated++;
    }
    if (invKey) invMap.set(invKey, ordKey);
  }

  const recordsArray = mapToSortedArray_(workMap);
  return {recordsArray, metrics};
}

function mergeRecordIntoExisting_(existingRec, newRec) {
  let changed = false;
  if (newRec.orderDate && !datesEqual_(existingRec.orderDate, newRec.orderDate)) {
    existingRec.orderDate = newRec.orderDate;
    changed = true;
  }
  if (newRec.shipDate && !datesEqual_(existingRec.shipDate, newRec.shipDate)) {
    existingRec.shipDate = newRec.shipDate;
    changed = true;
  }
  if (newRec.resellerPO !== '' && existingRec.resellerPO !== newRec.resellerPO) {
    existingRec.resellerPO = newRec.resellerPO;
    changed = true;
  }
  return changed;
}

function mapToSortedArray_(recMap) {
  const arr = [];
  recMap.forEach(rec => arr.push(rec));
  arr.sort((a, b) => {
    const ao = normalizeKey_(a.orderId);
    const bo = normalizeKey_(b.orderId);
    if (ao < bo) return -1;
    if (ao > bo) return 1;
    const ai = normalizeKey_(a.invoice);
    const bi = normalizeKey_(b.invoice);
    if (ai < bi) return -1;
    if (ai > bi) return 1;
    return 0;
  });
  return arr;
}

/** ===========================================================================
 * RECORD BUILD
 * ======================================================================== */
function buildRecordFromSourceRow_(rowObj, metrics) {
  const orderId = normalizeString_(rowObj[CANON.ORDER_ID]);
  const invoice = normalizeString_(rowObj[CANON.INVOICE]);
  let orderDate = rowObj[CANON.ORDER_DATE];
  let shipDate = rowObj[CANON.SHIP_DATE];
  const resellerPO = normalizeString_(rowObj[CANON.RESELLER_PO]);

  if (!invoice) {
    metrics.skippedMissingFields++;
    return null;
  }
  if (DEP_CFG.REQUIRE_ORDER_ID && !orderId) {
    metrics.skippedMissingFields++;
    return null;
  }

  orderDate = normalizeDate_(orderDate);
  shipDate = normalizeDate_(shipDate);
  if (!orderDate || !shipDate) {
    metrics.skippedInvalidDate++;
    return null;
  }

  return {
    orderId: orderId,
    invoice: invoice,
    orderDate: orderDate,
    shipDate: shipDate,
    resellerPO: resellerPO
  };
}

/** ===========================================================================
 * TARGET WRITE
 * ======================================================================== */
function writeOrdersDatabaseFull_(tgt, recordsArray) {
  ensureHeaders_(tgt, TARGET_HEADERS);

  // Clear existing body (only data columns)
  const lastRow = tgt.getLastRow();
  if (lastRow > 1) {
    tgt.getRange(2, 1, lastRow - 1, TARGET_HEADERS.length).clearContent();
  }

  if (!recordsArray.length) return;

  const ts = timestamp_();
  const out = recordsArray.map(rec => [
    rec.orderId,
    rec.invoice,
    rec.orderDate,
    rec.shipDate,
    rec.resellerPO,
    ts
  ]);
  tgt.getRange(2, 1, out.length, TARGET_HEADERS.length).setValues(out);
}

/** ===========================================================================
 * SHEET HELPERS
 * ======================================================================== */
function getSpreadsheet_() {
  if (DEP_CFG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(DEP_CFG.SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetByNameOrCreate_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    ensureHeaders_(sheet, headers);
  }
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  const existing = getHeaderValues_(sheet);
  if (existing.length !== headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  // Overwrite to normalize
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function getHeaderValues_(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1);
  const vals = range.getValues()[0];
  return vals.map(v => normalizeString_(v));
}

function buildColumnMap_(headers) {
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const canon = canonicalizeHeader_(headers[i]);
    if (canon) map[canon] = i + 1;
  }
  return map;
}

function canonicalizeHeader_(h) {
  if (!h) return '';
  const s = String(h).trim().toLowerCase();
  const norm = s.replace(/[#]/g, '').replace(/[^a-z0-9]+/g, '');

  // Order ID
  if (norm === 'orderid' || norm === 'order' || norm === 'ordernumber' || norm === 'orderno') {
    return CANON.ORDER_ID;
  }
  // Invoice
  if (norm === 'invoice' || norm === 'invoiceno' || norm === 'invoicenumber' || norm === 'invoiceid') {
    return CANON.INVOICE;
  }
  // Hardware Order Date
  if (norm === 'hardwareorderdate' || norm === 'orderdate' || norm === 'hworderdate') {
    return CANON.ORDER_DATE;
  }
  // Hardware Ship Date
  if (norm === 'hardwareshipdate' || norm === 'shipdate' || norm === 'hwshipdate' || norm === 'shippingdate') {
    return CANON.SHIP_DATE;
  }
  // Reseller PO
  if (norm === 'resellerpo' || norm === 'po' || norm === 'purchaseorder' || norm === 'resellerpurchaseorder') {
    return CANON.RESELLER_PO;
  }
  // Last Synced
  if (norm === 'lastsynced' || norm === 'synced' || norm === 'lastupdate' || norm === 'lastupdated') {
    return CANON.LAST_SYNCED;
  }
  return '';
}

function trackedSourceColIndexes_(colMap) {
  const out = [];
  TRACKED_SOURCE_COLUMNS.forEach(k => {
    if (colMap[k]) out.push(colMap[k]);
  });
  return out;
}

/** ===========================================================================
 * NORMALIZATION
 * ======================================================================== */
function normalizeString_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function normalizeKey_(v) {
  return normalizeString_(v).toUpperCase();
}

function normalizeDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (v === null || v === undefined || v === '') return null;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}

function datesEqual_(d1, d2) {
  if (!d1 && !d2) return true;
  if (!d1 || !d2) return false;
  return d1.getTime() === d2.getTime();
}

function compositeKey_(orderId, invoice) {
  return normalizeKey_(orderId) + '::' + normalizeKey_(invoice);
}

/** ===========================================================================
 * METRICS & LOGGING
 * ======================================================================== */
function initMetrics_() {
  return {
    scannedRows: 0,
    eligibleRows: 0,
    inserted: 0,
    updated: 0,
    skippedMissingFields: 0,
    skippedInvalidDate: 0,
    invoiceConflicts: 0
  };
}

function timestamp_() {
  return Utilities.formatDate(new Date(), DEP_CFG.TIMEZONE, DEP_CFG.TIMESTAMP_FMT);
}

function log_() {
  Logger.log.apply(Logger, arguments);
  try { console.log.apply(console, arguments); } catch (_e) {}
}

/** ===========================================================================
 * TEST HARNESS
 * ======================================================================== */
function test_syncOrdersDatabase() {
  const ss = getSpreadsheet_();
  const src = ss.getSheetByName(DEP_CFG.SOURCE_SHEET_NAME);
  const tgt = getSheetByNameOrCreate_(ss, DEP_CFG.TARGET_SHEET_NAME, TARGET_HEADERS);

  const now = new Date();
  const d1 = now;
  const d2 = new Date(now.getTime() + 24 * 3600 * 1000);
  const d3 = new Date(now.getTime() + 2 * 24 * 3600 * 1000);

  const colMap = buildColumnMapFromSheet_(src, false);
  const lastCol = src.getLastColumn();
  const newRows = [];

  const blankRow = () => Array(lastCol).fill('');

  // Test row A: valid new
  let r = blankRow();
  if (colMap[CANON.ORDER_ID]) r[colMap[CANON.ORDER_ID] - 1] = 'TEST-1';
  if (colMap[CANON.INVOICE]) r[colMap[CANON.INVOICE] - 1] = 'INV-1';
  if (colMap[CANON.ORDER_DATE]) r[colMap[CANON.ORDER_DATE] - 1] = d1;
  if (colMap[CANON.SHIP_DATE]) r[colMap[CANON.SHIP_DATE] - 1] = d2;
  if (colMap[CANON.RESELLER_PO]) r[colMap[CANON.RESELLER_PO] - 1] = 'PO-1';
  newRows.push(r);

  // Test row B: same order, new invoice
  r = blankRow();
  if (colMap[CANON.ORDER_ID]) r[colMap[CANON.ORDER_ID] - 1] = 'TEST-1';
  if (colMap[CANON.INVOICE]) r[colMap[CANON.INVOICE] - 1] = 'INV-2';
  if (colMap[CANON.ORDER_DATE]) r[colMap[CANON.ORDER_DATE] - 1] = d1;
  if (colMap[CANON.SHIP_DATE]) r[colMap[CANON.SHIP_DATE] - 1] = d3;
  if (colMap[CANON.RESELLER_PO]) r[colMap[CANON.RESELLER_PO] - 1] = 'PO-2';
  newRows.push(r);

  // Test row C: missing Order ID -> skip
  r = blankRow();
  if (colMap[CANON.INVOICE]) r[colMap[CANON.INVOICE] - 1] = 'INV-NO-ORDER';
  if (colMap[CANON.ORDER_DATE]) r[colMap[CANON.ORDER_DATE] - 1] = d1;
  if (colMap[CANON.SHIP_DATE]) r[colMap[CANON.SHIP_DATE] - 1] = d2;
  newRows.push(r);

  // Test row D: conflicting invoice
  r = blankRow();
  if (colMap[CANON.ORDER_ID]) r[colMap[CANON.ORDER_ID] - 1] = 'TEST-2';
  if (colMap[CANON.INVOICE]) r[colMap[CANON.INVOICE] - 1] = 'INV-1'; // conflict
  if (colMap[CANON.ORDER_DATE]) r[colMap[CANON.ORDER_DATE] - 1] = d1;
  if (colMap[CANON.SHIP_DATE]) r[colMap[CANON.SHIP_DATE] - 1] = d3;
  newRows.push(r);

  const startRow = src.getLastRow() + 1;
  src.getRange(startRow, 1, newRows.length, lastCol).setValues(newRows);
  log_('Test rows written at row ' + startRow);

  syncOrdersDatabase();

  const data = tgt.getDataRange().getValues();
  log_('Orders Database after test: ' + JSON.stringify(data));
  // Uncomment to clean test rows
  // src.deleteRows(startRow, newRows.length);
}
