/* =============================================================================
 * PICK LIST
 * Builds a warehouse pick list from Masterordersheet > stockcheckRTS.
 * ============================================================================= */

var PICK_LIST_SHEET_NAME = 'Picklist';

// [SECURITY PATCH] Removed hardcoded Spreadsheet IDs.
var PICK_LIST_ORDER_WORKFLOW_SPREADSHEET_ID = 'YOUR_MASTER_SPREADSHEET_ID_HERE';
var PICK_LIST_SOURCE_SHEET_NAME = 'stockcheckRTS';

// [SECURITY PATCH] Removed hardcoded Spreadsheet IDs.
var PICK_LIST_INVENTORY_SPREADSHEET_ID =
  typeof CATALOG_SPREADSHEET_ID !== 'undefined'
    ? CATALOG_SPREADSHEET_ID
    : 'YOUR_INVENTORY_SPREADSHEET_ID_HERE';

var PICK_LIST_MATCHING_TABLE_SHEET_NAME = 'MatchingTable';
var PICK_LIST_INVENTORY_SHEET_NAME = 'Inventory';
var PICK_LIST_CONFIGURATION_SHEET_NAMES = ['Configuration', 'Config'];
var PICK_LIST_DRIVE_ROOT_FOLDER_NAME = 'picklist';

var PICK_LIST_HEADERS = [
  'PortalOrderID',
  'Internal OrderNo',
  'Staff Notes',
  'Qty',
  'Image',
  'Image URL',
  'SKU',
  'Location'
];

var PICK_LIST_COLS = {
  PORTAL_ORDER_ID: 1,
  INTERNAL_ORDERNO: 2,
  STAFF_NOTES: 3,
  QTY: 4,
  IMAGE: 5,
  IMAGE_URL: 6,
  SKU: 7,
  LOCATION: 8
};

var PICK_LIST_MATCHING_COLS = {
  VendorSKU: 1,
  PortalSKU: 2
};

var PICK_LIST_INVENTORY_COLS = {
  VendorSKU: 1,
  Image: 2,
  Item_Name: 3,
  Stock: 4,
  Product_Cost: 5,
  Location: 6,
  IMAGE_URL: 7
};

/**
 * Menu action: creates/refreshes the PickList sheet and opens a visual preview.
 */
function generatePickList() {
  var readySs = SpreadsheetApp.openById(PICK_LIST_ORDER_WORKFLOW_SPREADSHEET_ID);
  var readySheet = readySs.getSheetByName(PICK_LIST_SOURCE_SHEET_NAME);

  if (!readySheet) {
    pickListShowAlert_('Pick List Error', 'stockcheckRTS sheet was not found in the Masterordersheet spreadsheet.', 'error');
    return;
  }

  var inventorySs = SpreadsheetApp.openById(PICK_LIST_INVENTORY_SPREADSHEET_ID);
  var matchingSheet = inventorySs.getSheetByName(PICK_LIST_MATCHING_TABLE_SHEET_NAME);
  var inventorySheet = inventorySs.getSheetByName(PICK_LIST_INVENTORY_SHEET_NAME);
  if (!matchingSheet || !inventorySheet) {
    pickListShowAlert_('Pick List Notice', 'Inventory or MatchingTable was not found. Continuing with stockcheckRTS location values only.', 'warning');
  }

  var rows = buildPickListRows_(readySheet, matchingSheet, inventorySheet);
  var phoneNumbers = getPickListPhoneNumbers_(SpreadsheetApp.getActiveSpreadsheet());
  var pickSheet = getOrCreatePickListSheet_(readySs);
  writePickListSheet_(pickSheet, rows);
  showPickListPopup_(rows, pickSheet, phoneNumbers);
}

/**
 * Trigger-safe refresh used by stockcheckRTS/inventory sync.
 * It updates the Picklist sheet but does not open the popup.
 */
function refreshPickListFromStockcheckRTS_(suppressUi) {
  var readySs = SpreadsheetApp.openById(PICK_LIST_ORDER_WORKFLOW_SPREADSHEET_ID);
  var readySheet = readySs.getSheetByName(PICK_LIST_SOURCE_SHEET_NAME);
  if (!readySheet) {
    if (!suppressUi) pickListShowAlert_('Pick List Error', 'stockcheckRTS sheet was not found.', 'error');
    return { rows: 0 };
  }

  var matchingSheet = null;
  var inventorySheet = null;
  try {
    var inventorySs = SpreadsheetApp.openById(PICK_LIST_INVENTORY_SPREADSHEET_ID);
    matchingSheet = inventorySs.getSheetByName(PICK_LIST_MATCHING_TABLE_SHEET_NAME);
    inventorySheet = inventorySs.getSheetByName(PICK_LIST_INVENTORY_SHEET_NAME);
  } catch (err) {
    Logger.log('refreshPickListFromStockcheckRTS_ inventory read error: ' + err);
  }

  var rows = buildPickListRows_(readySheet, matchingSheet, inventorySheet);
  var pickSheet = getOrCreatePickListSheet_(readySs);
  writePickListSheet_(pickSheet, rows);

  if (!suppressUi) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Picklist refreshed from stockcheckRTS.', 'Picklist', 5);
  }

  return { rows: rows.length };
}

/**
 * Reads stockcheckRTS rows and enriches each row with Inventory location data.
 * Image URL is always taken from the order flow (All Orders -> stockcheckRTS),
 * never from Inventory/Matching lookups.
 * @param {Sheet} readySheet
 * @param {Sheet} matchingSheet
 * @param {Sheet} inventorySheet
 * @return {Object[]}
 */
function buildPickListRows_(readySheet, matchingSheet, inventorySheet) {
  var headerRow = pickListDetectHeaderRowByAliases_(readySheet, ['PortalOrderID', 'SKU'], 8);
  var lastRow = readySheet.getLastRow();
  var lastCol = readySheet.getLastColumn();
  if (lastRow <= headerRow || lastCol < 1) return [];

  var headers = readySheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var headerMap = buildPickListHeaderMap_(headers);
  var data = readySheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues();

  var portalMap = buildPickListPortalToVendorMap_(matchingSheet);
  var inventoryMap = buildPickListInventoryMap_(inventorySheet);
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var source = data[i];
    var sku = String(valueFromPickListAliases_(source, headerMap, ['sku', 'portalsku'])).trim();
    var portalOrderId = String(valueFromPickListAliases_(source, headerMap, ['portalorderid', 'portalorder', 'orderid'])).trim();
    var qty = valueFromPickListAliases_(source, headerMap, ['qty', 'quantity']);

    if (!sku && !portalOrderId && !qty) continue;

    var vendorSKU = portalMap[sku] || '';
    var inv = vendorSKU ? inventoryMap[vendorSKU] : null;
    var readyImageUrl = String(valueFromPickListAliases_(source, headerMap, ['imageurl', 'imageurl1'])).trim();
    var sourceLocation = String(valueFromPickListAliases_(source, headerMap, ['location', 'bin', 'rack'])).trim();
    var imageUrl = readyImageUrl;
    var location = '';
    var issue = '';

    if (!sku) {
      issue = 'SKU missing';
    } else if (sourceLocation) {
      location = sourceLocation;
    } else if (!vendorSKU) {
      issue = 'No MatchingTable entry';
    } else if (!inv) {
      issue = 'Vendor SKU missing in Inventory';
    } else if (!inv.location) {
      issue = 'Location missing';
    } else {
      location = inv.location;
    }

    rows.push({
      portalOrderId: portalOrderId,
      internalOrderNo: valueFromPickListAliases_(source, headerMap, [
        'internalorderno',
        'internalorderid',
        'internalordernumber'
      ]),
      staffNotes: valueFromPickListAliases_(source, headerMap, ['staffnotes', 'staffnote']),
      qty: qty,
      imageUrl: imageUrl,
      sku: sku,
      location: location || issue,
      issue: issue,
      vendorSKU: vendorSKU
    });
  }

  rows.sort(function(a, b) {
    var aIssue = a.issue ? 1 : 0;
    var bIssue = b.issue ? 1 : 0;
    if (aIssue !== bIssue) return aIssue - bIssue;

    var locA = String(a.location || '').toLowerCase();
    var locB = String(b.location || '').toLowerCase();
    if (locA < locB) return -1;
    if (locA > locB) return 1;

    var ordA = String(a.portalOrderId || '').toLowerCase();
    var ordB = String(b.portalOrderId || '').toLowerCase();
    if (ordA < ordB) return -1;
    if (ordA > ordB) return 1;
    return 0;
  });

  return rows;
}

/**
 * Builds VendorSKU -> inventory details for pick-list display.
 * @param {Sheet} sheet
 * @return {Object}
 */
function buildPickListInventoryMap_(sheet) {
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var data = sheet.getRange(2, 1, lastRow - 1, Math.max(PICK_LIST_INVENTORY_COLS.IMAGE_URL, PICK_LIST_INVENTORY_COLS.Location)).getValues();
  var map = {};

  for (var i = 0; i < data.length; i++) {
    var vendorSKU = String(data[i][PICK_LIST_INVENTORY_COLS.VendorSKU - 1]).trim();
    if (!vendorSKU) continue;

    map[vendorSKU] = {
      location: String(data[i][PICK_LIST_INVENTORY_COLS.Location - 1]).trim(),
      imageUrl: String(data[i][PICK_LIST_INVENTORY_COLS.IMAGE_URL - 1]).trim()
    };
  }

  return map;
}

function buildPickListPortalToVendorMap_(sheet) {
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};

  for (var i = 0; i < data.length; i++) {
    var vendorSKU = String(data[i][PICK_LIST_MATCHING_COLS.VendorSKU - 1]).trim();
    var portalSKU = String(data[i][PICK_LIST_MATCHING_COLS.PortalSKU - 1]).trim();
    if (portalSKU && vendorSKU) {
      map[portalSKU] = vendorSKU;
    }
  }

  return map;
}

function buildPickListHeaderMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = pickListNormalizeHeaderKey_(headers[i]);
    if (key && !Object.prototype.hasOwnProperty.call(map, key)) {
      map[key] = i;
    }
  }
  return map;
}

function valueFromPickListAliases_(row, headerMap, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var key = pickListNormalizeHeaderKey_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(headerMap, key)) {
      return row[headerMap[key]];
    }
  }
  return '';
}

function pickListNormalizeHeaderKey_(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function pickListDetectHeaderRowByAliases_(sheet, aliases, maxRowsToScan) {
  if (!sheet) return 1;

  var maxRows = Math.min(maxRowsToScan || 5, sheet.getLastRow());
  var maxCols = sheet.getLastColumn();
  if (maxRows < 1 || maxCols < 1) return 1;

  var normalizedAliases = {};
  for (var i = 0; i < aliases.length; i++) {
    normalizedAliases[pickListNormalizeHeaderKey_(aliases[i])] = true;
  }

  for (var r = 1; r <= maxRows; r++) {
    var row = sheet.getRange(r, 1, 1, maxCols).getValues()[0];
    for (var c = 0; c < row.length; c++) {
      if (normalizedAliases[pickListNormalizeHeaderKey_(row[c])]) return r;
    }
  }

  return 1;
}

function pickListShowAlert_(title, message, type) {
  if (typeof showStyledAlert_ === 'function') {
    showStyledAlert_(title, message, type);
    return;
  }

  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function getOrCreatePickListSheet_(ss) {
  var sheet = ss.getSheetByName(PICK_LIST_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PICK_LIST_SHEET_NAME);
  }
  return sheet;
}

function ensurePickListSheetSize_(sheet, rowCount, colCount) {
  if (sheet.getMaxRows() < rowCount) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rowCount - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < colCount) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), colCount - sheet.getMaxColumns());
  }
}

function getPickListPhoneNumbers_(sourceSpreadsheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  if (!ss) return [];
  var sheet = ss.getSheetByName('Config') || ss.getSheetByName('Configuration');
  if (!sheet) return [];

  var values = sheet.getDataRange().getValues();
  var phones = [];
  var seen = {};

  var phoneCol = -1;
  var nameCol = -1;
  var headerRow = -1;

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var val = String(values[r][c] || '').replace(/[^A-Za-z0-9]/g, '').toLowerCase();
      if (val === 'picklistphone' || val === 'picklistphones') {
        phoneCol = c;
        if (headerRow === -1) headerRow = r;
      } else if (val === 'contactname' || val === 'contact') {
        nameCol = c;
        if (headerRow === -1) headerRow = r;
      }
    }
    if (phoneCol !== -1) break;
  }

  if (phoneCol !== -1) {
    for (var rr = headerRow + 1; rr < values.length; rr++) {
      if (!values[rr]) continue;
      var phoneVal = values[rr][phoneCol];
      var nameVal = nameCol !== -1 ? values[rr][nameCol] : '';
      addPickListPhonesFromValue_(phoneVal, nameVal, phones, seen);
    }
  }

  return phones;
}

function findPickListContactNameCol_(headerRow, phoneCol) {
  for (var c = 0; c < headerRow.length; c++) {
    var key = pickListNormalizeHeaderKey_(headerRow[c]);
    if (key === 'contactname' || key === 'contact') {
      return c;
    }
  }

  if (phoneCol + 1 < headerRow.length) {
    var nextKey = pickListNormalizeHeaderKey_(headerRow[phoneCol + 1]);
    if (nextKey === 'contactname' || nextKey === 'contact') {
      return phoneCol + 1;
    }
  }

  return -1;
}

function addPickListPhonesFromValue_(value, displayName, phones, seen) {
  var raw = String(value == null ? '' : value).trim();
  if (!raw) return;

  var parts = raw.split(/[\n,;|]+/);
  var name = String(displayName == null ? '' : displayName).trim();
  for (var i = 0; i < parts.length; i++) {
    var phoneText = String(parts[i]).trim();
    if (!phoneText) continue;

    var whatsappNumber = normalizePickListWhatsAppNumber_(phoneText);
    if (!whatsappNumber || seen[whatsappNumber]) continue;

    phones.push({
      display: name || phoneText,
      whatsappNumber: whatsappNumber
    });
    seen[whatsappNumber] = true;
  }
}

function normalizePickListWhatsAppNumber_(value) {
  var digits = String(value || '').replace(/\D/g, '');
  if (!digits) return '';

  if (digits.length === 10) {
    return '91' + digits;
  }

  return digits;
}

/**
 * Rebuilds the PickList sheet with colorful columns and IMAGE formulas.
 * @param {Sheet} sheet
 * @param {Object[]} rows
 */
function writePickListSheet_(sheet, rows) {
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  var bandings = sheet.getBandings();
  for (var b = 0; b < bandings.length; b++) {
    bandings[b].remove();
  }

  sheet.clear();
  ensurePickListSheetSize_(sheet, Math.max(rows.length + 1, 2), PICK_LIST_HEADERS.length);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, PICK_LIST_HEADERS.length).setValues([PICK_LIST_HEADERS]);

  var headerColors = [
    '#1d4ed8',
    '#7c3aed',
    '#0f766e',
    '#ea580c',
    '#be123c',
    '#0891b2',
    '#4b5563',
    '#15803d'
  ];

  sheet.getRange(1, 1, 1, PICK_LIST_HEADERS.length)
    .setBackgrounds([headerColors])
    .setFontColors([['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff']])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  if (rows.length === 0) {
    sheet.getRange(2, 1).setValue('No stockcheckRTS rows found.');
    sheet.setColumnWidths(1, PICK_LIST_HEADERS.length, 145);
    return;
  }

  var values = [];
  var formulas = [];
  var backgrounds = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    values.push([
      row.portalOrderId,
      row.internalOrderNo,
      row.staffNotes,
      row.qty,
      '',
      row.imageUrl,
      row.sku,
      row.location
    ]);

    if (row.imageUrl) {
      formulas.push(['=IMAGE("' + String(row.imageUrl).replace(/"/g, '""') + '")']);
    } else {
      formulas.push(['']);
    }

    backgrounds.push([
      '#eff6ff',
      '#f5f3ff',
      '#ecfdf5',
      '#fff7ed',
      '#fdf2f8',
      '#ecfeff',
      '#f9fafb',
      row.issue ? '#fee2e2' : '#f0fdf4'
    ]);
  }

  sheet.getRange(2, 1, values.length, PICK_LIST_HEADERS.length)
    .setValues(values)
    .setBackgrounds(backgrounds)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sheet.getRange(2, PICK_LIST_COLS.IMAGE, formulas.length, 1).setFormulas(formulas);

  sheet.setColumnWidth(PICK_LIST_COLS.PORTAL_ORDER_ID, 170);
  sheet.setColumnWidth(PICK_LIST_COLS.INTERNAL_ORDERNO, 135);
  sheet.setColumnWidth(PICK_LIST_COLS.STAFF_NOTES, 220);
  sheet.setColumnWidth(PICK_LIST_COLS.QTY, 70);
  sheet.setColumnWidth(PICK_LIST_COLS.IMAGE, 96);
  sheet.setColumnWidth(PICK_LIST_COLS.IMAGE_URL, 260);
  sheet.setColumnWidth(PICK_LIST_COLS.SKU, 150);
  sheet.setColumnWidth(PICK_LIST_COLS.LOCATION, 150);
  sheet.setRowHeight(1, 34);
  sheet.setRowHeights(2, rows.length, 76);
  sheet.getRange(2, PICK_LIST_COLS.QTY, rows.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
  sheet.getRange(2, PICK_LIST_COLS.LOCATION, rows.length, 1).setFontWeight('bold');
  sheet.getRange(1, 1, rows.length + 1, PICK_LIST_HEADERS.length).createFilter();
}

function showPickListPopup_(rows, pickSheet, phoneNumbers) {
  phoneNumbers = phoneNumbers || [];
  var totalQty = 0;
  var missingCount = 0;
  var locationMap = {};
  var tableRows = '';
  var baseMessage = '';
  var phoneCheckboxes = '';

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var qtyNum = Number(row.qty) || 0;
    totalQty += qtyNum;
    if (row.issue) missingCount++;
    if (row.location && !row.issue) locationMap[row.location] = true;

    tableRows += '<tr class="' + (row.issue ? 'warn' : '') + '">'
      + '<td>' + pickListEscapeHtml_(row.portalOrderId) + '</td>'
      + '<td>' + pickListEscapeHtml_(row.internalOrderNo) + '</td>'
      + '<td class="notes">' + pickListEscapeHtml_(row.staffNotes) + '</td>'
      + '<td class="qty">' + pickListEscapeHtml_(row.qty) + '</td>'
      + '<td class="thumb">' + buildPickListPopupImage_(row.imageUrl) + '</td>'
      + '<td class="url">' + pickListEscapeHtml_(row.imageUrl) + '</td>'
      + '<td class="sku">' + pickListEscapeHtml_(row.sku) + '</td>'
      + '<td class="loc">' + pickListEscapeHtml_(row.location) + '</td>'
      + '</tr>';
  }

  baseMessage = buildPickListWhatsAppIntroMessage_();

  if (!tableRows) {
    tableRows = '<tr><td colspan="8" class="empty">No stockcheckRTS rows found.</td></tr>';
  }

  for (var p = 0; p < phoneNumbers.length; p++) {
    phoneCheckboxes += '<label class="contact">'
      + '<input type="checkbox" class="pickContact" value="' + pickListEscapeHtml_(phoneNumbers[p].whatsappNumber) + '" checked>'
      + '<span>'
      + pickListEscapeHtml_(phoneNumbers[p].display)
      + '</span></label>';
  }

  if (!phoneCheckboxes) {
    phoneCheckboxes = '<div class="noContacts">No Pick_List_phone found</div>';
  }

  var locationCount = 0;
  for (var loc in locationMap) {
    if (Object.prototype.hasOwnProperty.call(locationMap, loc)) locationCount++;
  }

  var html = '<!DOCTYPE html><html><head><base target="_top">'
    + '<style>'
    + '*{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;background:#f8fafc;color:#111827}'
    + '.top{padding:18px 22px;background:linear-gradient(135deg,#0f766e,#2563eb);color:white}'
    + 'h2{margin:0;font-size:20px;line-height:1.2}.sub{margin-top:4px;font-size:12px;opacity:.88}'
    + '.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;padding:14px 16px;background:#ffffff;border-bottom:1px solid #e5e7eb}'
    + '.stat{border-radius:10px;padding:10px 12px;color:#111827;border:1px solid #e5e7eb;background:#f9fafb}'
    + '.stat b{display:block;font-size:20px}.stat span{font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:#64748b}'
    + '.stat.warn{background:#fff7ed;border-color:#fed7aa}.stat.bad{background:#fef2f2;border-color:#fecaca}'
    + '.wrap{height:470px;overflow:auto;background:#ffffff}'
    + 'table{width:100%;border-collapse:separate;border-spacing:0;min-width:980px}'
    + 'th{position:sticky;top:0;z-index:2;padding:9px 10px;text-align:left;font-size:11px;text-transform:uppercase;color:#fff;background:#334155}'
    + 'td{padding:8px 10px;border-bottom:1px solid #e5e7eb;font-size:12px;vertical-align:middle;background:#fff}'
    + 'tr:nth-child(even) td{background:#f8fafc}tr.warn td{background:#fff1f2}'
    + '.qty{text-align:center;font-weight:700;font-size:15px}.sku,.loc{font-weight:700}.loc{color:#166534}.warn .loc{color:#b91c1c}'
    + '.notes{max-width:190px;white-space:normal}.url{max-width:220px;word-break:break-all;color:#0369a1;font-size:11px}'
    + '.thumb{width:72px;text-align:center}.thumb img{width:54px;height:54px;object-fit:contain;border-radius:8px;border:1px solid #e5e7eb;background:white}'
    + '.empty{text-align:center;padding:32px;color:#64748b}.share{display:grid;grid-template-columns:auto 1fr auto;gap:10px;align-items:center;padding:12px 16px;background:#eef2ff;border-top:1px solid #dbeafe}'
    + '.shareTitle{font-size:12px;font-weight:700;color:#334155}.contacts{display:flex;gap:8px;align-items:center;overflow:auto;padding-bottom:2px}'
    + '.contact{display:flex;align-items:center;gap:6px;white-space:nowrap;border:1px solid #c7d2fe;background:white;border-radius:999px;padding:7px 10px;font-size:12px;font-weight:700;color:#1e293b}'
    + '.contact input{margin:0}.contact span{max-width:160px;overflow:hidden;text-overflow:ellipsis}.noContacts{font-size:12px;color:#64748b}'
    + '.shareButtons{display:flex;gap:8px;align-items:center}.share .hint{grid-column:2 / 4;font-size:11px;color:#64748b}.actions{display:flex;justify-content:flex-end;gap:10px;padding:12px 16px;background:#f8fafc;border-top:1px solid #e5e7eb}'
    + 'button{border:0;border-radius:8px;padding:9px 16px;font-weight:700;cursor:pointer;color:#fff;background:#2563eb}'
    + 'button.secondary{background:#64748b}button.whatsapp{background:#16a34a}button.copy{background:#7c3aed}button.mini{padding:7px 10px;background:#475569}button:disabled{opacity:.55;cursor:not-allowed}'
    + '</style></head><body>'
    + '<div class="top"><h2>Warehouse Pick List</h2><div class="sub">Sheet refreshed: '
    + pickListEscapeHtml_(pickSheet.getName()) + '</div></div>'
    + '<div class="stats">'
    + '<div class="stat"><b>' + rows.length + '</b><span>Order lines</span></div>'
    + '<div class="stat"><b>' + totalQty + '</b><span>Total qty</span></div>'
    + '<div class="stat"><b>' + locationCount + '</b><span>Locations</span></div>'
    + '<div class="stat ' + (missingCount ? 'bad' : '') + '"><b>' + missingCount + '</b><span>Needs check</span></div>'
    + '</div>'
    + '<div class="wrap"><table><thead><tr>'
    + '<th>PortalOrderID</th><th>Internal OrderNo</th><th>Staff Notes</th><th>Qty</th>'
    + '<th>Image</th><th>Image URL</th><th>SKU</th><th>Location</th>'
    + '</tr></thead><tbody>' + tableRows + '</tbody></table></div>'
    + '<div class="share">'
    + '<div class="shareTitle">WhatsApp</div><div class="contacts">' + phoneCheckboxes + '</div>'
    + '<div class="shareButtons">'
    + '<button class="mini" onclick="setPickListContacts(true)" ' + (phoneNumbers.length ? '' : 'disabled') + '>Select All</button>'
    + '<button class="mini" onclick="setPickListContacts(false)" ' + (phoneNumbers.length ? '' : 'disabled') + '>Clear</button>'
    + '<button class="whatsapp" onclick="openPickListWhatsApp()" ' + (phoneNumbers.length ? '' : 'disabled') + '>Send Whatsapp </button>'
    + '<button class="copy" onclick="copyPickListMessage()">Copy PDF Message</button></div>'
    + '<span class="hint" id="copyStatus">Contact names come from Configuration > Contact Name, numbers from Pick_List_phone. PDFs are saved in Drive > picklist > Month Pick List.</span>'
    + '</div>'
    + '<div class="actions"><button class="secondary" onclick="google.script.host.close()">Close</button>'
    + '<button onclick="google.script.run.withSuccessHandler(function(){google.script.host.close()}).withFailureHandler(function(e){console.error(\'Picklist error:\',e);google.script.host.close()}).generatePickList()">Refresh</button></div>'
    + '<script>'
    + 'var PICK_LIST_BASE_MESSAGE=' + JSON.stringify(baseMessage) + ';'
    + 'var PICK_LIST_PDF_CACHE=null;'
    + 'function getPickListCheckedContacts(){return Array.prototype.slice.call(document.querySelectorAll(".pickContact:checked"));}'
    + 'function setPickListContacts(checked){Array.prototype.slice.call(document.querySelectorAll(".pickContact")).forEach(function(el){el.checked=checked;});}'
    + 'function withPickListPdf(callback){'
    + 'if(PICK_LIST_PDF_CACHE){callback(null,PICK_LIST_PDF_CACHE);return;}'
    + 'var status=document.getElementById("copyStatus");if(status){status.textContent="Preparing PDF...";}'
    + 'google.script.run'
    + '.withSuccessHandler(function(res){PICK_LIST_PDF_CACHE=res||null;callback(null,PICK_LIST_PDF_CACHE);})'
    + '.withFailureHandler(function(err){callback(err||new Error("Could not create PDF."));})'
    + '.createPickListPdfForWhatsApp();'
    + '}'
    + 'function buildPickListWhatsAppText(pdf){'
    + 'var lines=[PICK_LIST_BASE_MESSAGE];'
    + 'if(pdf&&pdf.fileUrl){lines.push("");lines.push("PDF ("+(pdf.fileName||"picklist.pdf")+"): "+pdf.fileUrl);}'
    + 'return lines.join("\\n");'
    + '}'
    + 'function openPickListWhatsApp(){'
    + 'var selected=getPickListCheckedContacts();'
    + 'if(!selected.length){alert("Select at least one contact.");return;}'
    + 'withPickListPdf(function(err,pdf){'
    + 'var status=document.getElementById("copyStatus");'
    + 'if(err){if(status){status.textContent="Could not create PDF. Check script permissions and try again.";}alert("Could not create PDF.");return;}'
    + 'var text=buildPickListWhatsAppText(pdf);'
    + 'selected.forEach(function(el){var url="https://wa.me/"+el.value+"?text="+encodeURIComponent(text);window.open(url,"_blank");});'
    + 'if(status){status.textContent="Opened "+selected.length+" WhatsApp chat(s) with PDF link: "+(pdf&&pdf.fileName?pdf.fileName:"picklist.pdf")+".";}' 
    + '});'
    + '}'
    + 'function copyPickListMessage(){'
    + 'var status=document.getElementById("copyStatus");'
    + 'withPickListPdf(function(err,pdf){'
    + 'if(err){if(status){status.textContent="Could not create PDF. Check script permissions and try again.";}alert("Could not create PDF.");return;}'
    + 'var text=buildPickListWhatsAppText(pdf);'
    + 'function done(){if(status){status.textContent="PDF message copied. Paste it in WhatsApp if needed.";}}'
    + 'if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(text).then(done);return;}'
    + 'var t=document.createElement("textarea");t.value=text;document.body.appendChild(t);t.select();document.execCommand("copy");document.body.removeChild(t);done();'
    + '});'
    + '}'
    + '</script>'
    + '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(1120).setHeight(680),
    'Pick List'
  );
}

function buildPickListWhatsAppMessage_(rows, totalQty) {
  var lines = [];
  lines.push('Warehouse Pick List');
  lines.push('Order lines: ' + rows.length);
  lines.push('Total qty: ' + totalQty);
  lines.push('');

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    lines.push((i + 1) + '. Order: ' + (row.portalOrderId || '-'));
    lines.push('Internal: ' + (row.internalOrderNo || '-'));
    lines.push('SKU: ' + (row.sku || '-'));
    lines.push('Qty: ' + (row.qty || '-'));
    lines.push('Location: ' + (row.location || '-'));
    if (row.staffNotes) lines.push('Staff note: ' + row.staffNotes);
    if (row.imageUrl) lines.push('Image: ' + row.imageUrl);
    lines.push('');
  }

  return lines.join('\n');
}

function buildPickListWhatsAppIntroMessage_() {
  return [
    'Here is the picklist PDF.',
    'Please pick and bring all listed orders.'
  ].join('\n');
}

function createPickListPdfForWhatsApp() {
  var ss = SpreadsheetApp.openById(PICK_LIST_ORDER_WORKFLOW_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PICK_LIST_SHEET_NAME);
  if (!sheet) {
    throw new Error('PickList sheet was not found.');
  }

  var timezone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Kolkata';
  var now = new Date();
  var stamp = Utilities.formatDate(now, timezone, 'dd-MM-yy_HH-mm-ss');
  var fileName = 'picklist_' + stamp + '.pdf';
  var imageUrlCol = PICK_LIST_COLS.IMAGE_URL;
  var imageUrlWasHidden = sheet.isColumnHiddenByUser(imageUrlCol);
  var blob;

  try {
    if (!imageUrlWasHidden) {
      sheet.hideColumns(imageUrlCol, 1);
    }
    SpreadsheetApp.flush();
    Utilities.sleep(500);
    blob = exportPickListSheetPdfBlob_(ss.getId(), sheet.getSheetId(), fileName);
  } finally {
    if (!imageUrlWasHidden) {
      sheet.showColumns(imageUrlCol, 1);
    }
  }

  var targetFolder = getPickListDriveMonthFolder_(now, timezone);
  var file = targetFolder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // Some domains block link sharing. The file URL still works for authorized users.
  }

  return {
    fileId: file.getId(),
    fileName: file.getName(),
    fileUrl: file.getUrl(),
    folderName: targetFolder.getName()
  };
}

function exportPickListSheetPdfBlob_(spreadsheetId, sheetId, fileName) {
  var base = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export';
  var params = [
    'format=pdf',
    'size=A4',
    'portrait=false',
    'fitw=true',
    'sheetnames=false',
    'printtitle=false',
    'pagenumbers=false',
    'gridlines=false',
    'fzr=true',
    'top_margin=0.30',
    'bottom_margin=0.30',
    'left_margin=0.25',
    'right_margin=0.25',
    'attachment=true',
    'gid=' + encodeURIComponent(String(sheetId))
  ];

  var response = UrlFetchApp.fetch(base + '?' + params.join('&'), {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  var statusCode = response.getResponseCode();
  if (statusCode !== 200) {
    throw new Error('PDF export failed with code ' + statusCode + ': ' + response.getContentText());
  }

  return response.getBlob().setName(fileName);
}

function getPickListDriveMonthFolder_(date, timezone) {
  var root = getOrCreatePickListRootFolder_();
  var monthLabel = Utilities.formatDate(date, timezone, 'MMMM');
  var monthFolderName = monthLabel + ' Pick List';
  var folders = root.getFoldersByName(monthFolderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return root.createFolder(monthFolderName);
}

function getOrCreatePickListRootFolder_() {
  var rootFolder = DriveApp.getRootFolder();
  var folders = rootFolder.getFoldersByName(PICK_LIST_DRIVE_ROOT_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }

  var titleCaseName = PICK_LIST_DRIVE_ROOT_FOLDER_NAME.charAt(0).toUpperCase()
    + PICK_LIST_DRIVE_ROOT_FOLDER_NAME.slice(1).toLowerCase();
  folders = rootFolder.getFoldersByName(titleCaseName);
  if (folders.hasNext()) {
    return folders.next();
  }

  folders = DriveApp.getFoldersByName(PICK_LIST_DRIVE_ROOT_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }

  folders = DriveApp.getFoldersByName(titleCaseName);
  if (folders.hasNext()) {
    return folders.next();
  }

  return rootFolder.createFolder(PICK_LIST_DRIVE_ROOT_FOLDER_NAME);
}

function buildPickListPopupImage_(url) {
  var cleanUrl = String(url || '').trim();
  if (!cleanUrl) return '<span style="color:#94a3b8;font-size:11px;">No image</span>';
  return '<img src="' + pickListEscapeHtml_(cleanUrl) + '" alt="">';
}

function pickListEscapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
