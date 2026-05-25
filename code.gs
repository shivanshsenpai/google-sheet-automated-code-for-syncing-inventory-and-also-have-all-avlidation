/* =============================================================================
 * ORDER MANAGEMENT SYSTEM — new_order.gs
 * Google Apps Script for New_Orders sheet automation
 * ============================================================================= */

// ─────────────────────────────────────────────────────────────────────────────
//  1. CONSTANTS & COLUMN MAPPINGS
// ─────────────────────────────────────────────────────────────────────────────

var SHEET_NAMES = {
  NEW_ORDERS: "New_Orders",
  ALL_ORDERS: "All Orders",
  CONFIG: "Config",
  STOCKCHECK_RTS: "stockcheckRTS",
  CREATE_LABEL: "CreateLabels",
  PICKLIST: "Picklist",
  DISPATCHED: "Dispached",
  MATCHING_TABLE: "MatchingTable",
  INVENTORY: "Inventory",
};

// [SECURITY UPDATE] Removed hardcoded IDs
var SHIPPING_TARGET_SPREADSHEET_ID = "YOUR_SHIPPING_SPREADSHEET_ID_HERE";
var CATALOG_SPREADSHEET_ID = "YOUR_CATALOG_SPREADSHEET_ID_HERE";

var INVENTORY_HISTORY_LOG_SHEET_NAME = "log";
var NEW_ORDERS_EDIT_TRIGGER_HANDLER = "handleNewOrdersAuthorizedEdit";
var NEW_ORDERS_EDIT_TRIGGER_INSTALLED_KEY =
  "NEW_ORDERS_AUTHORIZED_EDIT_TRIGGER_INSTALLED";

var HEADERS = {
  PORTAL_ORDER_ID: 1, DELIVERY_PICKUP_DATE: 2, ORDER_STATUS: 3, PAYMENT_STATUS: 4,
  STAFF_NOTES: 5, BRAND_NAME: 6, INTERNAL_ORDERNO: 7, PURCHASE_DATE: 8,
  SALES_CHANNEL: 9, SKU: 10, ITEM_NAME: 11, CATEGORY: 12, IMAGE: 13,
  ORDERNOTE: 14, QTY: 15, CURRENCY: 16, CURRENCY_PRICE: 17, FULLNAME: 18,
  ADDRESSLINE1: 19, ADDRESSLINE2: 20, CITY: 21, STATE: 22, PINCODE: 23,
  COUNTRY: 24, PHONE: 25, COURIER_NAME: 26, TRACKING_CODE: 27, STATUS: 28,
  SHIPPING_CHARGE: 29, IMAGE_URL: 30, LISTING_URL: 31, NEW_TRACKING: 32,
  Conversion_Rate: 33, Price_in_INR: 34, Shipping_Charge_Product: 35,
  Product_Cost: 36, Maximum_Expense: 37, Actual_Expense: 38, Maximum_Profit: 39,
  Actual_Profit: 40, STOCK: 41, what_to_fix: 42,
};

var TOTAL_COLUMNS = 42;
var HEADER_ROW = 2;
var DATA_START_ROW = 3;
var WORKFLOW_HEADER_ROW = 1;
var WORKFLOW_DATA_START_ROW = 2;

var STOCKCHECK_COLS = { PORTAL_ORDER_ID: 1, INTERNAL_ORDERNO: 2, STAFF_NOTES: 3, QTY: 4, IMAGE: 5, IMAGE_URL: 6, SKU: 7, LOCATION: 8, ActualStockStatus: 9, DoCreateLabels: 10 };
var CREATE_LABEL_COLS = { INTERNAL_ORDERNO: 1, SKU: 2, FULLNAME: 3, ADDRESSLINE1: 4, ADDRESSLINE2: 5, CITY: 6, STATE: 7, PINCODE: 8, COUNTRY: 9, PHONE: 10, LabelCreated: 11 };
var DISPATCHED_COLS = { INTERNAL_ORDERNO: 1, SKU: 2, FULLNAME: 3, ADDRESSLINE1: 4, ADDRESSLINE2: 5, CITY: 6, STATE: 7, PINCODE: 8, COUNTRY: 9, PHONE: 10, DISPACHED_STATUS: 11 };
var CONFIG_COLS = { Country: 1, Charge: 2, Magic_Number: 3, Currency_Code: 4, Brand_Name: 5, Initial_code: 6, Last_Maximum: 7, Unique_Sales_Channel: 8, Category: 9, Sales_Channel: 10, Payment_Status: 11 };
var MATCHING_COLS = { VendorSKU: 1, PortalSKU: 2 };
var INVENTORY_COLS = { VendorSKU: 1, Image: 2, Item_Name: 3, Stock: 4, Product_Cost: 5, Location: 6, IMAGE_URL: 7 };

var VALIDATION_MAP = { CATEGORY: "Category", COUNTRY: "Country", CURRENCY: "Currency_Code", BRAND_NAME: "Brand_Name", SALES_CHANNEL: "Sales_Channel", PAYMENT_STATUS: "Payment_Status" };
var MANDATORY_FIELDS = [
  { key: "PORTAL_ORDER_ID", label: "Portal Order ID" }, { key: "BRAND_NAME", label: "Brand Name" },
  { key: "SKU", label: "SKU" }, { key: "QTY", label: "Quantity" },
  { key: "CURRENCY", label: "Currency" }, { key: "CURRENCY_PRICE", label: "Currency Price" },
  { key: "COUNTRY", label: "Country" }, { key: "FULLNAME", label: "Full Name" },
  { key: "SALES_CHANNEL", label: "Sales Channel" }, { key: "CATEGORY", label: "Category" },
  { key: "DELIVERY_PICKUP_DATE", label: "Delivery/Pickup Date" }, { key: "PAYMENT_STATUS", label: "Payment Status" },
  { key: "ITEM_NAME", label: "Item Name" }, { key: "ORDERNOTE", label: "Order Note" },
  { key: "ADDRESSLINE1", label: "Address Line1" }, { key: "CITY", label: "City" },
  { key: "STATE", label: "State" }, { key: "PINCODE", label: "Pincode" },
  { key: "IMAGE_URL", label: "Image url" }, { key: "PURCHASE_DATE", label: "Purchase date" },
];

var ERROR_RED = "#ffcdd2"; 
var CLEAR_COLOR = "#ffffff"; 

// [SECURITY UPDATE] Utility for preventing XSS in UI rendering
function escapeHtml_(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// [SECURITY UPDATE] Utility for mitigating Formula Injection
function sanitizeInput_(str) {
  if (str === null || str === undefined) return "";
  var s = String(str);
  if (/^[=+\-@]/.test(s)) {
    return "'" + s; 
  }
  return s;
}

// ─────────────────────────────────────────────────────────────────────────────
//  2. CONFIG HELPERS
// ─────────────────────────────────────────────────────────────────────────────

function getCatalogSpreadsheet_() {
  return SpreadsheetApp.openById(CATALOG_SPREADSHEET_ID);
}

function getCatalogSheet_(sheetName) {
  var ss = getCatalogSpreadsheet_();
  return ss ? ss.getSheetByName(sheetName) : null;
}

function getMatchingTableSheet_() {
  return getCatalogSheet_(SHEET_NAMES.MATCHING_TABLE);
}

function getInventorySheet_() {
  return getCatalogSheet_(SHEET_NAMES.INVENTORY);
}

function getConfigSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
}

function getConfigValues_(configColAliases) {
  var sheet = getConfigSheet_();
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colIdx = -1;
  var aliases = Array.isArray(configColAliases) ? configColAliases : [configColAliases];

  var normalizedAliases = {};
  for (var a = 0; a < aliases.length; a++) {
    normalizedAliases[normalizeHeaderKey_(aliases[a])] = true;
  }

  for (var c = 0; c < headers.length; c++) {
    if (normalizedAliases[normalizeHeaderKey_(headers[c])]) {
      colIdx = c + 1;
      break;
    }
  }

  if (colIdx === -1) {
    if (typeof configColAliases === "string" && CONFIG_COLS[configColAliases]) {
      colIdx = CONFIG_COLS[configColAliases];
    } else {
      return [];
    }
  }

  var data = sheet.getRange(2, colIdx, lastRow - 1, 1).getValues();
  var values = [];
  var seen = {};
  for (var i = 0; i < data.length; i++) {
    var v = String(data[i][0]).trim();
    if (v !== "" && !seen[v]) {
      values.push(v);
      seen[v] = true;
    }
  }
  return values;
}

function getConfigColRange_(configColAliases) {
  var sheet = getConfigSheet_();
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return null;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colIdx = -1;
  var aliases = Array.isArray(configColAliases) ? configColAliases : [configColAliases];

  var normalizedAliases = {};
  for (var a = 0; a < aliases.length; a++) {
    normalizedAliases[normalizeHeaderKey_(aliases[a])] = true;
  }

  for (var c = 0; c < headers.length; c++) {
    if (normalizedAliases[normalizeHeaderKey_(headers[c])]) {
      colIdx = c + 1;
      break;
    }
  }

  if (colIdx === -1) {
    if (typeof configColAliases === "string" && CONFIG_COLS[configColAliases]) {
      colIdx = CONFIG_COLS[configColAliases];
    } else {
      return null;
    }
  }

  var maxRows = Math.max(2, sheet.getMaxRows());
  return sheet.getRange(2, colIdx, maxRows - 1, 1);
}

var _configCache = null;
function getConfigData_() {
  if (_configCache) return _configCache;
  var sheet = getConfigSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  _configCache = data;
  return data;
}

function getBrandConfigMap_() {
  var data = getConfigData_();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var brand = String(data[i][CONFIG_COLS.Brand_Name - 1]).trim();
    if (brand) {
      map[brand] = {
        Initial_code: String(data[i][CONFIG_COLS.Initial_code - 1]).trim(),
        Last_Maximum: Number(data[i][CONFIG_COLS.Last_Maximum - 1]) || 0,
        Charge: Number(data[i][CONFIG_COLS.Charge - 1]) || 0,
        Magic_Number: Number(data[i][CONFIG_COLS.Magic_Number - 1]) || 0,
        configRow: i + 2,
      };
    }
  }
  return map;
}

function getCountryConfigMap_() {
  var data = getConfigData_();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var country = String(data[i][CONFIG_COLS.Country - 1]).trim();
    if (country) {
      var info = {
        Charge: Number(data[i][CONFIG_COLS.Charge - 1]) || 0,
        Magic_Number: Number(data[i][CONFIG_COLS.Magic_Number - 1]) || 0,
      };
      map[country] = info;
      map[normalizeCountryKey_(country)] = info; 
    }
  }
  return map;
}

function normalizeCountryKey_(country) {
  return String(country || "").trim().toLowerCase().replace(/\s+/g, " ");
}

// ─────────────────────────────────────────────────────────────────────────────
//  3. VALIDATION LOGIC
// ─────────────────────────────────────────────────────────────────────────────

function applyValidationsToRange_(sheet, startRow, endRow) {
  if (startRow > endRow) return;
  var numRows = endRow - startRow + 1;
  var configSheet = getConfigSheet_();
  if (!configSheet) return;

  var maxConfigRows = Math.max(2, configSheet.getMaxRows());

  for (var headerKey in VALIDATION_MAP) {
    var colIdx = HEADERS[headerKey];
    if (!colIdx) continue;

    var configColKey = VALIDATION_MAP[headerKey];
    var configColIdx = CONFIG_COLS[configColKey];
    if (!configColIdx) continue;

    var listRange = configSheet.getRange(2, configColIdx, maxConfigRows - 1, 1);
    var targetRange = sheet.getRange(startRow, colIdx, numRows, 1);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(listRange, true) 
      .setAllowInvalid(false)
      .build();
    targetRange.setDataValidation(rule);
  }
}

function applyAllValidations() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;
  applyValidationsToRange_(sheet, DATA_START_ROW, lastRow);
  showStyledAlert_("Validations Applied", "Dropdown validations have been applied to all data rows.", "success");
}

// ─────────────────────────────────────────────────────────────────────────────
//  4. PRODUCT COST LOOKUP
// ─────────────────────────────────────────────────────────────────────────────

function updateImageFormulas_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var urls = sheet.getRange(startRow, HEADERS.IMAGE_URL, numRows, 1).getValues();
  var formulas = [];

  for (var i = 0; i < numRows; i++) {
    var url = String(urls[i][0]).trim();
    if (!url) {
      formulas.push([""]);
      continue;
    }
    formulas.push(['=IMAGE("' + url.replace(/"/g, '""') + '")']);
  }
  sheet.getRange(startRow, HEADERS.IMAGE, numRows, 1).setFormulas(formulas);
}

function updateInventoryImageFormulas_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var urls = sheet.getRange(startRow, INVENTORY_COLS.IMAGE_URL, numRows, 1).getValues();
  var formulas = [];
  for (var i = 0; i < numRows; i++) {
    var url = String(urls[i][0]).trim();
    if (!url) {
      formulas.push([""]);
    } else {
      formulas.push(['=IMAGE("' + url.replace(/"/g, '""') + '")']);
    }
  }
  sheet.getRange(startRow, INVENTORY_COLS.Image, numRows, 1).setFormulas(formulas);
}

function getPortalToVendorMap_() {
  var sheet = getMatchingTableSheet_();
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var vendorSKU = String(data[i][MATCHING_COLS.VendorSKU - 1]).trim();
    var portalSKU = String(data[i][MATCHING_COLS.PortalSKU - 1]).trim();
    if (portalSKU && vendorSKU) {
      map[portalSKU] = vendorSKU;
    }
  }
  return map;
}

function getInventoryMap_() {
  var sheet = getInventorySheet_();
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var map = {};
  for (var i = 0; i < data.length; i++) {
    var vsku = String(data[i][INVENTORY_COLS.VendorSKU - 1]).trim();
    if (vsku) {
      map[vsku] = {
        Image: data[i][INVENTORY_COLS.Image - 1],
        Item_Name: String(data[i][INVENTORY_COLS.Item_Name - 1]).trim(),
        Stock: Number(data[i][INVENTORY_COLS.Stock - 1]) || 0,
        Product_Cost: Number(data[i][INVENTORY_COLS.Product_Cost - 1]) || 0,
        Location: String(data[i][INVENTORY_COLS.Location - 1]).trim(),
        invRow: i + 2,
      };
    }
  }
  return map;
}

function lookupProductInfo_(sku, portalMap, invMap) {
  var vendorSKU = portalMap[sku];
  if (!vendorSKU) return { ok: false, reason: "vendor_sku_missing", sku: sku };
  
  var inv = invMap[vendorSKU];
  if (!inv) return { ok: false, reason: "vendor_details_missing", sku: sku, vendorSKU: vendorSKU };
  
  return { ok: true, reason: "ok", Product_Cost: inv.Product_Cost, Item_Name: inv.Item_Name, Image: inv.Image, VendorSKU: vendorSKU, Stock: inv.Stock, invRow: inv.invRow };
}

function isSkuLookupErrorMessage_(msg) {
  var m = String(msg || "").toLowerCase();
  return (m.indexOf("sku not found") !== -1 || m.indexOf("vendor sku not found") !== -1 || m.indexOf("vendor details not found in inventory") !== -1);
}

function getStockAvailabilityStatus_(stock, qty) {
  var s = Number(stock);
  var q = Number(qty);
  if (isNaN(s) || isNaN(q)) return "";
  if (s >= q && s > 0) return "Available : " + s;
  var needed = Math.max(0, q - s);
  return "Out of stock : " + s + "\nNeed : " + needed;
}

function isPendingOrderStatus_(status) {
  return String(status || "").trim().toLowerCase() === "pending";
}

function updateProductInfoForRows_(sheet, rows, suppressAlerts) {
  if (!rows || rows.length === 0) return;
  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var errors = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var sku = String(sheet.getRange(row, HEADERS.SKU).getValue()).trim();
    if (!sku) {
      sheet.getRange(row, HEADERS.Product_Cost).setValue("");
      sheet.getRange(row, HEADERS.STOCK).setValue("");
      continue;
    }

    var info = lookupProductInfo_(sku, portalMap, invMap);
    var errCell = sheet.getRange(row, HEADERS.what_to_fix);
    var curVal = String(errCell.getValue()).trim();
    var errorsArr = curVal.split("|").map(function (e) { return e.trim(); }).filter(function (e) {
      return e.length > 0 && !isSkuLookupErrorMessage_(e) && e !== "Stock is 0" && e !== "✅ All fields complete";
    });

    if (info.ok) {
      sheet.getRange(row, HEADERS.Product_Cost).setValue(info.Product_Cost);
      sheet.getRange(row, HEADERS.STOCK).setValue(info.Stock);
      errCell.setValue(errorsArr.join(" | "));
    } else {
      sheet.getRange(row, HEADERS.STOCK).setValue("");
      if (info.reason === "vendor_sku_missing") errorsArr.push("Vendor SKU not found: " + sku);
      else if (info.reason === "vendor_details_missing") errorsArr.push("Vendor details not found in Inventory: " + info.vendorSKU);
      else errorsArr.push("SKU not found: " + sku);
      
      errCell.setValue(errorsArr.join(" | "));
      errors.push("Row " + row + ": " + errorsArr[errorsArr.length - 1]);
    }
  }

  if (errors.length > 0 && !suppressAlerts) {
    showStyledAlert_("SKU Lookup Warning", errors.join("\n"), "warning");
  }
}

function batchUpdateProductInfo_(sheet, startRow, endRow, suppressAlerts) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var skuValues = sheet.getRange(startRow, HEADERS.SKU, numRows, 1).getValues();
  var existingErrVals = sheet.getRange(startRow, HEADERS.what_to_fix, numRows, 1).getValues();

  var costOut = [], stockOut = [], errOut = [], errors = [];

  for (var i = 0; i < numRows; i++) {
    var sku = String(skuValues[i][0]).trim();
    var curVal = String(existingErrVals[i][0]).trim();
    var errorsArr = curVal.split("|").map(function (e) { return e.trim(); }).filter(function (e) {
      return e.length > 0 && !isSkuLookupErrorMessage_(e) && e !== "Stock is 0" && e !== "✅ All fields complete";
    });

    if (!sku) {
      costOut.push([""]); stockOut.push([""]); errOut.push([errorsArr.join(" | ")]);
      continue;
    }
    var info = lookupProductInfo_(sku, portalMap, invMap);
    if (info.ok) {
      costOut.push([info.Product_Cost]); stockOut.push([info.Stock]); errOut.push([errorsArr.join(" | ")]);
    } else {
      costOut.push([""]); stockOut.push([""]);
      if (info.reason === "vendor_sku_missing") errorsArr.push("Vendor SKU not found: " + sku);
      else if (info.reason === "vendor_details_missing") errorsArr.push("Vendor details not found in Inventory: " + info.vendorSKU);
      else errorsArr.push("SKU not found: " + sku);
      errOut.push([errorsArr.join(" | ")]);
      errors.push("Row " + (startRow + i) + ": " + errorsArr[errorsArr.length - 1]);
    }
  }

  sheet.getRange(startRow, HEADERS.Product_Cost, numRows, 1).setValues(costOut);
  sheet.getRange(startRow, HEADERS.STOCK, numRows, 1).setValues(stockOut);
  sheet.getRange(startRow, HEADERS.what_to_fix, numRows, 1).setValues(errOut);

  if (errors.length > 0 && !suppressAlerts) {
    showStyledAlert_("SKU Lookup Issues", errors.length + " SKU(s) not found:\n" + errors.join("\n"), "warning");
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  5. STOCK LOGIC
// ─────────────────────────────────────────────────────────────────────────────

function decrementStockForRows_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var invSheet = getInventorySheet_();
  if (!invSheet) return;

  var changes = {}; 
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var sku = String(sheet.getRange(row, HEADERS.SKU).getValue()).trim();
    var qty = Number(sheet.getRange(row, HEADERS.QTY).getValue()) || 0;
    if (!sku || qty <= 0) continue;

    var vendorSKU = portalMap[sku];
    if (!vendorSKU || !invMap[vendorSKU]) continue;
    changes[vendorSKU] = (changes[vendorSKU] || 0) + qty;
  }

  for (var vsku in changes) {
    var inv = invMap[vsku];
    var newStock = Math.max(0, inv.Stock - changes[vsku]);
    invSheet.getRange(inv.invRow, INVENTORY_COLS.Stock).setValue(newStock);
  }
}

function restoreStock_(sku, qty) {
  if (!sku || qty <= 0) return;
  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var invSheet = getInventorySheet_();
  if (!invSheet) return;

  var vendorSKU = portalMap[sku];
  if (!vendorSKU || !invMap[vendorSKU]) return;

  var inv = invMap[vendorSKU];
  invSheet.getRange(inv.invRow, INVENTORY_COLS.Stock).setValue(inv.Stock + qty);
}

function getMaxFromAllOrders_(prefix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.ALL_ORDERS);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return 0;

  var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  var internalColIdx = -1;
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim().toUpperCase();
    if (h === "INTERNAL ORDERNO" || h === "INTERNAL_ORDERNO") {
      internalColIdx = i; break;
    }
  }
  if (internalColIdx < 0) return 0;

  var data = sheet.getRange(DATA_START_ROW, internalColIdx + 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  var max = 0;
  var pfxUpper = prefix.toUpperCase();
  for (var j = 0; j < data.length; j++) {
    var val = String(data[j][0]).trim();
    var match = val.match(/^([A-Za-z]+)(\d+)(?:-[A-Za-z0-9]+)?$/);
    if (match && match[1].toUpperCase() === pfxUpper) {
      max = Math.max(max, parseInt(match[2], 10));
    }
  }
  return max;
}

// ─────────────────────────────────────────────────────────────────────────────
//  6. INTERNAL ORDER ID
// ─────────────────────────────────────────────────────────────────────────────

function normalizeKeyPart_(v) { return String(v || "").trim().toLowerCase().replace(/\s+/g, " "); }
function isIndiaCountry_(country) {
  var c = normalizeKeyPart_(country);
  return c === "india" || c === "in";
}

function buildForeignAddressKey_(brand, fullName, address1, city, state, pincode, country) {
  if (isIndiaCountry_(country)) return "";
  var parts = [normalizeKeyPart_(brand), normalizeKeyPart_(fullName), normalizeKeyPart_(address1), normalizeKeyPart_(city), normalizeKeyPart_(state), normalizeKeyPart_(pincode), normalizeKeyPart_(country)];
  for (var i = 0; i < parts.length; i++) { if (!parts[i]) return ""; }
  return parts.join("||");
}

function assignInternalOrderNos_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  var brandMap = getBrandConfigMap_();
  var configSheet = getConfigSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  var rowCount = lastRow - DATA_START_ROW + 1;
  var allBrands = sheet.getRange(DATA_START_ROW, HEADERS.BRAND_NAME, rowCount, 1).getValues();
  var allOrderNos = sheet.getRange(DATA_START_ROW, HEADERS.INTERNAL_ORDERNO, rowCount, 1).getValues();
  var allPortalIds = sheet.getRange(DATA_START_ROW, HEADERS.PORTAL_ORDER_ID, rowCount, 1).getValues();
  var allCountries = sheet.getRange(DATA_START_ROW, HEADERS.COUNTRY, rowCount, 1).getValues();
  var allFullNames = sheet.getRange(DATA_START_ROW, HEADERS.FULLNAME, rowCount, 1).getValues();
  var allAddress1 = sheet.getRange(DATA_START_ROW, HEADERS.ADDRESSLINE1, rowCount, 1).getValues();
  var allCities = sheet.getRange(DATA_START_ROW, HEADERS.CITY, rowCount, 1).getValues();
  var allStates = sheet.getRange(DATA_START_ROW, HEADERS.STATE, rowCount, 1).getValues();
  var allPincodes = sheet.getRange(DATA_START_ROW, HEADERS.PINCODE, rowCount, 1).getValues();

  var brandMaxMap = {}, portalIdToOrderNo = {}, foreignKeyToOrderNo = {};

  for (var i = 0; i < allBrands.length; i++) {
    var b = String(allBrands[i][0]).trim();
    var pId = String(allPortalIds[i][0]).trim();
    var orderNo = String(allOrderNos[i][0]).trim();

    if (!b || !brandMap[b]) continue;
    var prefix = brandMap[b].Initial_code;

    if (pId && orderNo) portalIdToOrderNo[pId] = orderNo;
    var foreignKey = buildForeignAddressKey_(b, allFullNames[i][0], allAddress1[i][0], allCities[i][0], allStates[i][0], allPincodes[i][0], allCountries[i][0]);
    if (foreignKey && orderNo) foreignKeyToOrderNo[foreignKey] = orderNo;

    if (orderNo.indexOf(prefix) === 0) {
      var num = parseInt(orderNo.substring(prefix.length), 10);
      if (!isNaN(num)) brandMaxMap[b] = Math.max(brandMaxMap[b] || 0, num);
    }
  }

  var configUpdates = {}, writeRows = [], writeVals = [];

  for (var j = 0; j < rows.length; j++) {
    var row = rows[j];
    var idx = row - DATA_START_ROW;
    if (idx < 0 || idx >= rowCount) continue;

    var pId = String(allPortalIds[idx][0]).trim();
    var brand = String(allBrands[idx][0]).trim();
    var country = allCountries[idx][0];
    var fullName = allFullNames[idx][0];
    var address1 = allAddress1[idx][0];
    var city = allCities[idx][0];
    var state = allStates[idx][0];
    var pincode = allPincodes[idx][0];

    if (!brand || !brandMap[brand]) continue;

    var foreignKeyRow = buildForeignAddressKey_(brand, fullName, address1, city, state, pincode, country);
    var finalOrderNo = "";
    if (foreignKeyRow && foreignKeyToOrderNo[foreignKeyRow]) {
      finalOrderNo = foreignKeyToOrderNo[foreignKeyRow];
    } else if (pId && portalIdToOrderNo[pId]) {
      finalOrderNo = portalIdToOrderNo[pId];
    } else {
      var prefix = brandMap[brand].Initial_code;
      if (brandMaxMap[brand] === undefined) {
        brandMaxMap[brand] = Math.max(getMaxFromAllOrders_(prefix), brandMap[brand].Last_Maximum);
      }
      var nextNum = brandMaxMap[brand] + 1;
      brandMaxMap[brand] = nextNum;

      finalOrderNo = prefix + nextNum;
      if (foreignKeyRow) foreignKeyToOrderNo[foreignKeyRow] = finalOrderNo;
      if (pId) portalIdToOrderNo[pId] = finalOrderNo;
      configUpdates[brand] = nextNum;
    }
    writeRows.push(row);
    writeVals.push([finalOrderNo]);
  }

  for (var w = 0; w < writeRows.length; w++) {
    sheet.getRange(writeRows[w], HEADERS.INTERNAL_ORDERNO).setValue(writeVals[w][0]);
  }

  for (var changedBrand in configUpdates) {
    configSheet.getRange(brandMap[changedBrand].configRow, CONFIG_COLS.Last_Maximum).setValue(configUpdates[changedBrand]);
  }
}

function regenerateInternalOrderNos_(sheet, brandName) {
  var brandMap = getBrandConfigMap_();
  var configSheet = getConfigSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  var numRows = lastRow - DATA_START_ROW + 1;
  var allBrands = sheet.getRange(DATA_START_ROW, HEADERS.BRAND_NAME, numRows, 1).getValues();
  var orderNos = sheet.getRange(DATA_START_ROW, HEADERS.INTERNAL_ORDERNO, numRows, 1).getValues();
  var allPortalIds = sheet.getRange(DATA_START_ROW, HEADERS.PORTAL_ORDER_ID, numRows, 1).getValues();
  var allCountries = sheet.getRange(DATA_START_ROW, HEADERS.COUNTRY, numRows, 1).getValues();
  var allFullNames = sheet.getRange(DATA_START_ROW, HEADERS.FULLNAME, numRows, 1).getValues();
  var allAddress1 = sheet.getRange(DATA_START_ROW, HEADERS.ADDRESSLINE1, numRows, 1).getValues();
  var allCities = sheet.getRange(DATA_START_ROW, HEADERS.CITY, numRows, 1).getValues();
  var allStates = sheet.getRange(DATA_START_ROW, HEADERS.STATE, numRows, 1).getValues();
  var allPincodes = sheet.getRange(DATA_START_ROW, HEADERS.PINCODE, numRows, 1).getValues();

  var brandRows = {};
  for (var i = 0; i < numRows; i++) {
    var b = String(allBrands[i][0]).trim();
    if (!b || !brandMap[b]) continue;
    if (brandName && b !== brandName) continue;
    if (!brandRows[b]) brandRows[b] = [];
    brandRows[b].push(i); 
  }

  for (var brand in brandRows) {
    var prefix = brandMap[brand].Initial_code;
    var indices = brandRows[brand];
    var minNum = Infinity;
    for (var k = 0; k < indices.length; k++) {
      var orderNo = String(orderNos[indices[k]][0]).trim();
      if (orderNo.indexOf(prefix) === 0) {
        var n = parseInt(orderNo.substring(prefix.length), 10);
        if (!isNaN(n) && n > 0) minNum = Math.min(minNum, n);
      }
    }

    var baseMax = getMaxFromAllOrders_(prefix);
    var maxNum = Math.max(baseMax, minNum === Infinity ? 0 : minNum - 1);
    var pIdMap = {};
    var foreignMap = {};

    for (var j = 0; j < indices.length; j++) {
      var rIdx = indices[j];
      var pId = String(allPortalIds[rIdx][0]).trim();
      var foreignKey = buildForeignAddressKey_(brand, allFullNames[rIdx][0], allAddress1[rIdx][0], allCities[rIdx][0], allStates[rIdx][0], allPincodes[rIdx][0], allCountries[rIdx][0]);
      var seqNum;

      if (foreignKey && foreignMap[foreignKey]) seqNum = foreignMap[foreignKey];
      else if (pId && pIdMap[pId]) seqNum = pIdMap[pId];
      else {
        maxNum++; seqNum = maxNum;
        if (foreignKey) foreignMap[foreignKey] = seqNum;
        if (pId) pIdMap[pId] = seqNum;
      }
      orderNos[rIdx][0] = prefix + seqNum;
    }
    configSheet.getRange(brandMap[brand].configRow, CONFIG_COLS.Last_Maximum).setValue(maxNum);
  }
  sheet.getRange(DATA_START_ROW, HEADERS.INTERNAL_ORDERNO, numRows, 1).setValues(orderNos);
}

function regenerateAllInternalOrderNos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (!sheet) return;
  regenerateInternalOrderNos_(sheet, null);
  showStyledAlert_("Order IDs Regenerated", "All internal order numbers have been re-sequenced without gaps.", "success");
}

function resequenceFromManualEdit_(sheet, editedRow) {
  var brandMap = getBrandConfigMap_();
  var thisBrand = String(sheet.getRange(editedRow, HEADERS.BRAND_NAME).getValue()).trim();
  if (!thisBrand || !brandMap[thisBrand]) return;

  var newOrderNo = String(sheet.getRange(editedRow, HEADERS.INTERNAL_ORDERNO).getValue()).trim();
  var prefix = brandMap[thisBrand].Initial_code;
  if (newOrderNo.indexOf(prefix) !== 0) return; 

  var startNum = parseInt(newOrderNo.substring(prefix.length), 10);
  if (isNaN(startNum)) return;

  var lastRow = sheet.getLastRow();
  if (lastRow <= editedRow) {
    var configSheet = getConfigSheet_();
    var currentConfigMax = Number(configSheet.getRange(brandMap[thisBrand].configRow, CONFIG_COLS.Last_Maximum).getValue()) || 0;
    if (startNum > currentConfigMax) {
      configSheet.getRange(brandMap[thisBrand].configRow, CONFIG_COLS.Last_Maximum).setValue(startNum);
    }
    return;
  }

  var numRows = lastRow - editedRow;
  var allBrands = sheet.getRange(editedRow + 1, HEADERS.BRAND_NAME, numRows, 1).getValues();
  var allPortalIds = sheet.getRange(editedRow + 1, HEADERS.PORTAL_ORDER_ID, numRows, 1).getValues();
  var allCountries = sheet.getRange(editedRow + 1, HEADERS.COUNTRY, numRows, 1).getValues();
  var allFullNames = sheet.getRange(editedRow + 1, HEADERS.FULLNAME, numRows, 1).getValues();
  var allAddress1 = sheet.getRange(editedRow + 1, HEADERS.ADDRESSLINE1, numRows, 1).getValues();
  var allCities = sheet.getRange(editedRow + 1, HEADERS.CITY, numRows, 1).getValues();
  var allStates = sheet.getRange(editedRow + 1, HEADERS.STATE, numRows, 1).getValues();
  var allPincodes = sheet.getRange(editedRow + 1, HEADERS.PINCODE, numRows, 1).getValues();
  var orderNos = sheet.getRange(editedRow + 1, HEADERS.INTERNAL_ORDERNO, numRows, 1).getValues();

  var currentMax = startNum;
  var pIdMap = {}, foreignMap = {};
  var editedPortalId = String(sheet.getRange(editedRow, HEADERS.PORTAL_ORDER_ID).getValue()).trim();
  if (editedPortalId) pIdMap[editedPortalId] = startNum;
  
  var editedForeignKey = buildForeignAddressKey_(thisBrand, sheet.getRange(editedRow, HEADERS.FULLNAME).getValue(), sheet.getRange(editedRow, HEADERS.ADDRESSLINE1).getValue(), sheet.getRange(editedRow, HEADERS.CITY).getValue(), sheet.getRange(editedRow, HEADERS.STATE).getValue(), sheet.getRange(editedRow, HEADERS.PINCODE).getValue(), sheet.getRange(editedRow, HEADERS.COUNTRY).getValue());
  if (editedForeignKey) foreignMap[editedForeignKey] = startNum;

  for (var i = 0; i < numRows; i++) {
    var b = String(allBrands[i][0]).trim();
    if (b !== thisBrand) continue;

    var pId = String(allPortalIds[i][0]).trim();
    var foreignKey = buildForeignAddressKey_(thisBrand, allFullNames[i][0], allAddress1[i][0], allCities[i][0], allStates[i][0], allPincodes[i][0], allCountries[i][0]);
    var seqNum;

    if (foreignKey && foreignMap[foreignKey]) seqNum = foreignMap[foreignKey];
    else if (pId && pIdMap[pId]) seqNum = pIdMap[pId];
    else {
      currentMax++; seqNum = currentMax;
      if (foreignKey) foreignMap[foreignKey] = seqNum;
      if (pId) pIdMap[pId] = seqNum;
    }
    orderNos[i][0] = prefix + seqNum;
  }

  sheet.getRange(editedRow + 1, HEADERS.INTERNAL_ORDERNO, numRows, 1).setValues(orderNos);
  var configSheet = getConfigSheet_();
  var currentConfigMax = Number(configSheet.getRange(brandMap[thisBrand].configRow, CONFIG_COLS.Last_Maximum).getValue()) || 0;
  if (currentMax > currentConfigMax) {
    configSheet.getRange(brandMap[thisBrand].configRow, CONFIG_COLS.Last_Maximum).setValue(currentMax);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  7. FINANCIAL CALCULATIONS
// ─────────────────────────────────────────────────────────────────────────────

var _exchangeRatesCache = {};

function getConversionRateFromAPI_(currency) {
  var cur = String(currency || "").trim().toUpperCase();
  if (!cur) return 0;
  if (cur === "INR") return 1;
  if (_exchangeRatesCache[cur]) return _exchangeRatesCache[cur];

  try {
    var response = UrlFetchApp.fetch("https://api.exchangerate-api.com/v4/latest/" + cur, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return 0;
    var data = JSON.parse(response.getContentText());
    var rate = data && data.rates ? Number(data.rates.INR) : 0;
    if (rate && !isNaN(rate)) {
      _exchangeRatesCache[cur] = rate;
      return rate;
    }
  } catch (err) {
    Logger.log("getConversionRateFromAPI_ error: " + err);
  }
  return 0;
}

function calculateFinancialsForRows_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  var countryMap = getCountryConfigMap_();

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var currency = String(sheet.getRange(row, HEADERS.CURRENCY).getValue()).trim();
    var country = String(sheet.getRange(row, HEADERS.COUNTRY).getValue()).trim();
    var productCost = Number(sheet.getRange(row, HEADERS.Product_Cost).getValue()) || 0;

    if (!currency && !country && productCost === 0) {
      sheet.getRange(row, HEADERS.Price_in_INR).setValue("");
      sheet.getRange(row, HEADERS.Shipping_Charge_Product).setValue("");
      sheet.getRange(row, HEADERS.Maximum_Expense).setValue("");
      sheet.getRange(row, HEADERS.Actual_Expense).setValue("");
      sheet.getRange(row, HEADERS.Maximum_Profit).setValue("");
      sheet.getRange(row, HEADERS.Actual_Profit).setValue("");
      continue;
    }

    var currencyPrice = Number(sheet.getRange(row, HEADERS.CURRENCY_PRICE).getValue()) || 0;
    var convRateRaw = sheet.getRange(row, HEADERS.Conversion_Rate).getValue();
    var convRate = Number(convRateRaw);

    if (!convRate || isNaN(convRate)) {
      convRate = getConversionRateFromAPI_(currency);
      sheet.getRange(row, HEADERS.Conversion_Rate).setValue(convRate);
    }

    var shippingCharge = Number(sheet.getRange(row, HEADERS.SHIPPING_CHARGE).getValue()) || 0;
    var qty = Number(sheet.getRange(row, HEADERS.QTY).getValue()) || 1;
    var countryInfo = countryMap[country] || countryMap[normalizeCountryKey_(country)] || { Charge: 0, Magic_Number: 0 };
    var shippingChargeProduct = Number(countryInfo.Charge) || 0;
    var magicNumber = Number(countryInfo.Magic_Number) || 0;

    var priceInINR = currencyPrice * convRate;
    var maxExpense = priceInINR * magicNumber;
    var actualExpense = productCost + shippingChargeProduct;
    var maxProfit = priceInINR * 0.2;
    var actualProfit = maxExpense - actualExpense + maxProfit;

    sheet.getRange(row, HEADERS.Price_in_INR).setValue(priceInINR);
    sheet.getRange(row, HEADERS.Shipping_Charge_Product).setValue(shippingChargeProduct);
    sheet.getRange(row, HEADERS.Maximum_Expense).setValue(maxExpense);
    sheet.getRange(row, HEADERS.Actual_Expense).setValue(actualExpense);
    sheet.getRange(row, HEADERS.Maximum_Profit).setValue(maxProfit);
    sheet.getRange(row, HEADERS.Actual_Profit).setValue(actualProfit);
  }
}

function batchCalculateFinancials_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var countryMap = getCountryConfigMap_();
  var currencyPrices = sheet.getRange(startRow, HEADERS.CURRENCY_PRICE, numRows, 1).getValues();
  var convRates = sheet.getRange(startRow, HEADERS.Conversion_Rate, numRows, 1).getValues();
  var productCosts = sheet.getRange(startRow, HEADERS.Product_Cost, numRows, 1).getValues();
  var shippingCharges = sheet.getRange(startRow, HEADERS.SHIPPING_CHARGE, numRows, 1).getValues();
  var countries = sheet.getRange(startRow, HEADERS.COUNTRY, numRows, 1).getValues();
  var qtys = sheet.getRange(startRow, HEADERS.QTY, numRows, 1).getValues();
  var currencies = sheet.getRange(startRow, HEADERS.CURRENCY, numRows, 1).getValues();

  var convRateUpdates = [], priceInINROut = [], shipProdOut = [], maxExpOut = [], actExpOut = [], maxProfOut = [], actProfOut = [];

  for (var i = 0; i < numRows; i++) {
    var cur = String(currencies[i][0]).trim();
    var ctry = String(countries[i][0]).trim();
    var pc = Number(productCosts[i][0]) || 0;

    if (!cur && !ctry && pc === 0) {
      convRateUpdates.push([""]); priceInINROut.push([""]); shipProdOut.push([""]); maxExpOut.push([""]); actExpOut.push([""]); maxProfOut.push([""]); actProfOut.push([""]);
      continue;
    }

    var cp = Number(currencyPrices[i][0]) || 0;
    var cr = Number(convRates[i][0]);
    if (!cr || isNaN(cr)) cr = getConversionRateFromAPI_(cur);
    convRateUpdates.push([cr]);

    var sc = Number(shippingCharges[i][0]) || 0;
    var qty = Number(qtys[i][0]) || 1;
    var countryInfo = countryMap[ctry] || countryMap[normalizeCountryKey_(ctry)] || { Charge: 0, Magic_Number: 0 };
    var shippingChargeProduct = Number(countryInfo.Charge) || 0;
    var magicNumber = Number(countryInfo.Magic_Number) || 0;

    var prInINR = cp * cr;
    var maxExp = prInINR * magicNumber;
    var actExp = pc + shippingChargeProduct;
    var maxProf = prInINR * 0.2;
    var actProf = maxExp - actExp + maxProf;

    priceInINROut.push([prInINR]); shipProdOut.push([shippingChargeProduct]); maxExpOut.push([maxExp]); actExpOut.push([actExp]); maxProfOut.push([maxProf]); actProfOut.push([actProf]);
  }

  sheet.getRange(startRow, HEADERS.Conversion_Rate, numRows, 1).setValues(convRateUpdates);
  sheet.getRange(startRow, HEADERS.Price_in_INR, numRows, 1).setValues(priceInINROut);
  sheet.getRange(startRow, HEADERS.Shipping_Charge_Product, numRows, 1).setValues(shipProdOut);
  sheet.getRange(startRow, HEADERS.Maximum_Expense, numRows, 1).setValues(maxExpOut);
  sheet.getRange(startRow, HEADERS.Actual_Expense, numRows, 1).setValues(actExpOut);
  sheet.getRange(startRow, HEADERS.Maximum_Profit, numRows, 1).setValues(maxProfOut);
  sheet.getRange(startRow, HEADERS.Actual_Profit, numRows, 1).setValues(actProfOut);
}

// ─────────────────────────────────────────────────────────────────────────────
//  8. APPEND TO ALLORDERS
// ─────────────────────────────────────────────────────────────────────────────

function getExistingAllOrderIds_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dest = ss.getSheetByName(SHEET_NAMES.ALL_ORDERS);
  if (!dest) return {};
  var lastRow = dest.getLastRow();
  if (lastRow < 2) return {};
  var ids = dest.getRange(2, 1, lastRow - 1, 1).getValues();
  var set = {};
  for (var i = 0; i < ids.length; i++) {
    var id = String(ids[i][0]).trim();
    if (id) set[id] = true;
  }
  return set;
}

function appendSelectedToAllOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (!sheet) return;

  var selection = ss.getActiveRange();
  if (!selection) {
    showStyledAlert_("No Selection", "Please select the rows you want to push to AllOrders.", "warning");
    return;
  }

  var startRow = Math.max(selection.getRow(), DATA_START_ROW);
  var endRow = selection.getLastRow();
  if (endRow < DATA_START_ROW) {
    showStyledAlert_("Invalid Selection", "Please select data rows (not the header).", "warning");
    return;
  }

  validateAndLogErrors_(sheet, startRow, endRow);

  var numRows = endRow - startRow + 1;
  var allData = sheet.getRange(startRow, 1, numRows, TOTAL_COLUMNS).getValues();
  var existingIds = getExistingAllOrderIds_();
  var duplicates = [], validRows = [], validData = [], duplicatesToMark = [];

  for (var i = 0; i < numRows; i++) {
    var portalId = String(allData[i][HEADERS.PORTAL_ORDER_ID - 1]).trim();
    var errLog = String(allData[i][HEADERS.what_to_fix - 1]).trim();

    if (!portalId) { duplicates.push({ row: startRow + i, id: "(empty)", reason: "No Portal Order ID" }); continue; }
    if (errLog !== "✅ All fields complete") { duplicates.push({ row: startRow + i, id: portalId, reason: "Incomplete: " + errLog }); continue; }
    if (existingIds[portalId]) { duplicates.push({ row: startRow + i, id: portalId, reason: "Already exists in AllOrders" }); duplicatesToMark.push(startRow + i); } 
    else { validRows.push(startRow + i); validData.push(allData[i]); }
  }

  for (var d = 0; d < duplicatesToMark.length; d++) {
    sheet.getRange(duplicatesToMark[d], HEADERS.what_to_fix).setValue("Duplicate");
  }

  if (duplicates.length > 0 && validData.length === 0) {
    showDuplicateBlockedPopup_(duplicates);
    return;
  }

  if (duplicates.length > 0) {
    PropertiesService.getScriptProperties().setProperty("PENDING_PUSH_START", String(startRow));
    PropertiesService.getScriptProperties().setProperty("PENDING_PUSH_END", String(endRow));
    showPartialDuplicatePopup_(validData, duplicates);
    return;
  }

  PropertiesService.getScriptProperties().setProperty("PENDING_PUSH_START", String(startRow));
  PropertiesService.getScriptProperties().setProperty("PENDING_PUSH_END", String(endRow));
  showPushConfirmationPopup_(validData, numRows);
}

function appendAllToAllOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) {
    showStyledAlert_("No Data", "No orders found to push.", "info");
    return;
  }

  var startRow = DATA_START_ROW;
  var endRow = lastRow;

  validateAndLogErrors_(sheet, startRow, endRow);

  var numRows = endRow - startRow + 1;
  var allData = sheet.getRange(startRow, 1, numRows, TOTAL_COLUMNS).getValues();

  var existingIds = getExistingAllOrderIds_();
  var duplicates = [], validRows = [], validData = [], duplicatesToMark = [];

  for (var i = 0; i < numRows; i++) {
    var hasData = allData[i].join("").trim() !== "";
    if (!hasData) continue; 

    var portalId = String(allData[i][HEADERS.PORTAL_ORDER_ID - 1]).trim();
    var errLog = String(allData[i][HEADERS.what_to_fix - 1]).trim();

    if (!portalId) { duplicates.push({ row: startRow + i, id: "(empty)", reason: "No Portal Order ID" }); continue; }
    if (errLog !== "✅ All fields complete") { duplicates.push({ row: startRow + i, id: portalId, reason: "Incomplete: " + errLog }); continue; }
    if (existingIds[portalId]) { duplicates.push({ row: startRow + i, id: portalId, reason: "Already exists in AllOrders" }); duplicatesToMark.push(startRow + i); } 
    else { validRows.push(startRow + i); validData.push(allData[i]); }
  }

  for (var d = 0; d < duplicatesToMark.length; d++) {
    sheet.getRange(duplicatesToMark[d], HEADERS.what_to_fix).setValue("Duplicate");
  }

  if (validData.length === 0) {
    if (duplicates.length > 0) showDuplicateBlockedPopup_(duplicates);
    else showStyledAlert_("No Valid Data", "No valid orders found to push.", "info");
    return;
  }

  PropertiesService.getScriptProperties().setProperty("PENDING_PUSH_START", String(startRow));
  PropertiesService.getScriptProperties().setProperty("PENDING_PUSH_END", String(endRow));

  if (duplicates.length > 0) {
    showPartialDuplicatePopup_(validData, duplicates);
    return;
  }
  showPushConfirmationPopup_(validData, validData.length);
}

function executePushToAllOrders() {
  var props = PropertiesService.getScriptProperties();
  var startRow = Number(props.getProperty("PENDING_PUSH_START"));
  var endRow = Number(props.getProperty("PENDING_PUSH_END"));

  if (!startRow || !endRow) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
  var dest = ss.getSheetByName(SHEET_NAMES.ALL_ORDERS);
  if (!sheet || !dest) return;

  var numRows = endRow - startRow + 1;
  var allData = sheet.getRange(startRow, 1, numRows, TOTAL_COLUMNS).getValues();
  var existingIds = getExistingAllOrderIds_();
  var allOrdersColCount = Math.max(dest.getLastColumn(), 42);
  var stockStatusCol = findHeaderColByAliases_(dest, HEADER_ROW, ["stock", "stock_availibilty", "stock_availability", "stockavailability"]);
  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var pushed = 0;

  for (var i = 0; i < numRows; i++) {
    var portalId = String(allData[i][HEADERS.PORTAL_ORDER_ID - 1]).trim();
    if (portalId && !existingIds[portalId]) {
      var outRow = buildAllOrdersRowFromNewOrderRow_(allData[i], portalMap, invMap, allOrdersColCount, stockStatusCol);
      dest.appendRow(outRow);
      pushed++;
    }
  }

  if (pushed > 0) {
    try { refreshAllOrdersStockAvailability_(); } catch (e1) { Logger.log(e1); }
    try { syncPendingOrdersToStockcheckRTS_(true); } catch (e2) { Logger.log(e2); }
  }

  props.deleteProperty("PENDING_PUSH_START");
  props.deleteProperty("PENDING_PUSH_END");
}

function executePushAll() { executePushToAllOrders(); }

function isAllFieldsCompleteMessage_(value) { return String(value || "").toLowerCase().indexOf("all fields complete") !== -1; }
function appendRowsToSheet_(sheet, rows, colCount) {
  if (!sheet || !rows || rows.length === 0) return;
  var appendStart = Math.max(sheet.getLastRow() + 1, DATA_START_ROW);
  sheet.getRange(appendStart, 1, rows.length, colCount).setValues(rows);
}

// [SECURITY UPDATE] Escape dynamically injected HTML strings here 
function showDuplicateBlockedPopup_(duplicates) {
  var dupRows = "";
  for (var i = 0; i < duplicates.length; i++) {
    dupRows += "<tr>" +
      '<td style="padding:8px 12px;border-bottom:1px solid #ffcdd2;font-size:13px;color:#c62828;">Row ' + escapeHtml_(duplicates[i].row) + "</td>" +
      '<td style="padding:8px 12px;border-bottom:1px solid #ffcdd2;font-size:13px;font-weight:600;">' + escapeHtml_(duplicates[i].id) + "</td>" +
      '<td style="padding:8px 12px;border-bottom:1px solid #ffcdd2;font-size:13px;max-width:200px;word-wrap:break-word;">' + escapeHtml_(duplicates[i].reason) + "</td></tr>";
  }

  var html = _buildPopupShell_(
    "🚫 Import Blocked — Validation Issues",
    "linear-gradient(135deg, #c62828, #e53935)",
    '<table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#ffebee;">' +
    '<th style="padding:8px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#c62828;">Row</th>' +
    '<th style="padding:8px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#c62828;">Order ID</th>' +
    '<th style="padding:8px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#c62828;">Reason</th>' +
    "</tr></thead><tbody>" + dupRows + "</tbody></table>" +
    '<div style="margin-top:14px;padding:12px 16px;background:#ffebee;border-left:4px solid #c62828;border-radius:6px;font-size:13px;color:#b71c1c;">❌ <strong>All selected orders have validation issues.</strong><br>Check reasons above. Import cancelled.</div>',
    '<button class="btn" style="background:linear-gradient(135deg,#757575,#9e9e9e);" onclick="google.script.host.close()">Close</button>'
  );

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(540).setHeight(380), " ");
}

// [SECURITY UPDATE] Escape dynamically injected HTML strings here 
function showPartialDuplicatePopup_(validData, duplicates) {
  var dupRows = "";
  for (var i = 0; i < duplicates.length; i++) {
    dupRows += "<tr>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #ffcdd2;font-size:12px;">Row ' + escapeHtml_(duplicates[i].row) + "</td>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #ffcdd2;font-size:12px;font-weight:600;">' + escapeHtml_(duplicates[i].id) + "</td>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #ffcdd2;font-size:12px;"><span title="' + escapeHtml_(duplicates[i].reason) + '" style="background:#ffcdd2;color:#c62828;padding:2px 8px;border-radius:10px;font-size:11px;cursor:help;display:inline-block;max-width:180px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' + escapeHtml_(duplicates[i].reason) + "</span></td></tr>";
  }

  var validSummary = '<div style="display:flex;gap:16px;margin-bottom:12px;"><div style="flex:1;background:#e8f5e9;padding:12px;border-radius:10px;text-align:center;"><div style="font-size:28px;font-weight:700;color:#2e7d32;">' + validData.length + '</div><div style="font-size:11px;color:#558b2f;text-transform:uppercase;letter-spacing:0.5px;">Will Push</div></div><div style="flex:1;background:#ffebee;padding:12px;border-radius:10px;text-align:center;"><div style="font-size:28px;font-weight:700;color:#c62828;">' + duplicates.length + '</div><div style="font-size:11px;color:#b71c1c;text-transform:uppercase;letter-spacing:0.5px;">Issues Found</div></div></div>';

  var html = _buildPopupShell_(
    "⚠️ Validation Issues — Partial Push",
    "linear-gradient(135deg, #e65100, #ff8f00)",
    validSummary + '<div style="font-size:13px;font-weight:600;color:#616161;margin-bottom:6px;">Issue rows (will be skipped):</div><table style="width:100%;border-collapse:collapse;table-layout:fixed;"><thead><tr><th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#9e9e9e;width:25%;">Row</th><th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#9e9e9e;width:30%;">Order ID</th><th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#9e9e9e;width:45%;">Status</th></tr></thead><tbody>' + dupRows + "</tbody></table>",
    '<button class="btn" style="background:linear-gradient(135deg,#757575,#9e9e9e);" onclick="google.script.host.close()">Cancel</button><button class="btn" style="background:linear-gradient(135deg,#43a047,#66bb6a);margin-left:10px;" id="pushValidBtn" onclick="this.disabled=true;this.style.cursor=\'not-allowed\';document.body.style.cursor=\'not-allowed\';this.textContent=\'Processing...\';google.script.run.withSuccessHandler(function(){document.body.style.cursor=\'auto\';google.script.host.close()}).withFailureHandler(function(e){console.error(\'Push error:\',e);document.body.style.cursor=\'auto\';google.script.host.close()}).executePushToAllOrders()">Push ' + validData.length + " Valid Row(s)</button>"
  );

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(540).setHeight(420), " ");
}

// [SECURITY UPDATE] Escape dynamically injected HTML strings here 
function showPushConfirmationPopup_(validData, count) {
  var previewRows = "";
  var previewCount = Math.min(count, 6);
  for (var i = 0; i < previewCount; i++) {
    var orderId = String(validData[i][HEADERS.PORTAL_ORDER_ID - 1]).trim() || "—";
    var brand = String(validData[i][HEADERS.BRAND_NAME - 1]).trim() || "—";
    var sku = String(validData[i][HEADERS.SKU - 1]).trim() || "—";
    previewRows += "<tr>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #e8eaf6;font-size:12px;">' + escapeHtml_(orderId) + "</td>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #e8eaf6;font-size:12px;">' + escapeHtml_(brand) + "</td>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #e8eaf6;font-size:12px;">' + escapeHtml_(sku) + "</td></tr>";
  }
  if (count > previewCount) {
    previewRows += '<tr><td colspan="3" style="padding:8px;text-align:center;color:#9e9e9e;font-size:12px;">...and ' + (count - previewCount) + " more</td></tr>";
  }

  var html = _buildPopupShell_(
    "📋 Push " + count + " Order(s) to AllOrders",
    "linear-gradient(135deg, #1565c0, #42a5f5)",
    '<div style="background:#e3f2fd;padding:12px;border-radius:10px;text-align:center;margin-bottom:12px;"><div style="font-size:32px;font-weight:700;color:#1565c0;">' + count + '</div><div style="font-size:12px;color:#1976d2;text-transform:uppercase;letter-spacing:0.5px;">Orders Ready</div></div><table style="width:100%;border-collapse:collapse;"><thead><tr style="background:#e8eaf6;"><th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#5c6bc0;">Order ID</th><th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#5c6bc0;">Brand</th><th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#5c6bc0;">SKU</th></tr></thead><tbody>' + previewRows + '</tbody></table><div style="margin-top:10px;padding:10px 14px;background:#e8f5e9;border-left:4px solid #43a047;border-radius:6px;font-size:12px;color:#2e7d32;">✅ No duplicates found. All orders are safe to push.</div>',
    '<button class="btn" style="background:linear-gradient(135deg,#757575,#9e9e9e);" onclick="google.script.host.close()">Cancel</button><button class="btn" style="background:linear-gradient(135deg,#1565c0,#42a5f5);margin-left:10px;" id="confirmPushBtn" onclick="this.disabled=true;this.style.cursor=\'not-allowed\';document.body.style.cursor=\'not-allowed\';this.textContent=\'Processing...\';google.script.run.withSuccessHandler(function(){document.body.style.cursor=\'auto\';google.script.host.close()}).withFailureHandler(function(e){console.error(\'Push error:\',e);document.body.style.cursor=\'auto\';google.script.host.close()}).executePushAll()">Confirm Push</button>'
  );

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(440), " ");
}

function _buildPopupShell_(title, gradient, bodyContent, buttonsHtml) {
  return "<!DOCTYPE html><html><head>" +
    '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">' +
    "<style>* { margin:0; padding:0; box-sizing:border-box; } body { font-family:\"Inter\",sans-serif; background:#f5f5f5; padding:0; } .card { background:#fff; border-radius:16px; overflow:hidden; box-shadow:0 8px 32px rgba(0,0,0,0.12); animation:slideUp 0.3s cubic-bezier(0.22,1,0.36,1); } .header { background:" + gradient + "; padding:18px 22px; color:#fff; } .header h2 { font-size:17px; font-weight:700; letter-spacing:-0.3px; } .body { padding:18px 22px; } .btn-row { padding:14px 22px; display:flex; gap:10px; justify-content:flex-end; background:#fafafa; border-top:1px solid #f0f0f0; } .btn { padding:10px 22px; border:none; border-radius:8px; font-family:\"Inter\",sans-serif; font-size:13px; font-weight:600; cursor:pointer; color:#fff; transition:transform 0.15s,box-shadow 0.15s; box-shadow:0 4px 12px rgba(0,0,0,0.12); } .btn:hover { transform:translateY(-1px); box-shadow:0 6px 18px rgba(0,0,0,0.18); } @keyframes slideUp { from { opacity:0; transform:translateY(16px); } to { opacity:1; transform:translateY(0); } }</style></head><body>" +
    '<div class="card"><div class="header"><h2>' + escapeHtml_(title) + "</h2></div><div class="body">" + bodyContent + "</div><div class="btn-row">" + buttonsHtml + "</div></div></body></html>";
}

// ─────────────────────────────────────────────────────────────────────────────
//  9. ONEDIT TRIGGER
// ─────────────────────────────────────────────────────────────────────────────

function fillBrandFromPortalId_(sheet, rows) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return false;

  var numRows = lastRow - DATA_START_ROW + 1;
  var allPortals = sheet.getRange(DATA_START_ROW, HEADERS.PORTAL_ORDER_ID, numRows, 1).getValues();
  var allBrands = sheet.getRange(DATA_START_ROW, HEADERS.BRAND_NAME, numRows, 1).getValues();

  var portalToBrandMap = {};
  for (var i = 0; i < numRows; i++) {
    var pid = String(allPortals[i][0]).trim();
    var b = String(allBrands[i][0]).trim();
    if (pid && b && !portalToBrandMap[pid]) portalToBrandMap[pid] = b;
  }

  var brandChanged = false;
  for (var j = 0; j < rows.length; j++) {
    var r = rows[j];
    var pid = String(sheet.getRange(r, HEADERS.PORTAL_ORDER_ID).getValue()).trim();
    var existingBrand = String(sheet.getRange(r, HEADERS.BRAND_NAME).getValue()).trim();
    if (pid && portalToBrandMap[pid] && !existingBrand) {
      sheet.getRange(r, HEADERS.BRAND_NAME).setValue(portalToBrandMap[pid]);
      brandChanged = true;
    }
  }
  return brandChanged;
}

var ROW_COUNT_KEY = "NEW_ORDERS_ROW_COUNT";

function onEdit(e) {
  if (isLimitedAuthEditEvent_(e) && PropertiesService.getScriptProperties().getProperty(NEW_ORDERS_EDIT_TRIGGER_INSTALLED_KEY) === "true") return;
  handleNewOrdersEdit_(e);
}

function handleNewOrdersAuthorizedEdit(e) { handleNewOrdersEdit_(e); }
function isLimitedAuthEditEvent_(e) { return e && e.authMode && String(e.authMode) === "LIMITED"; }

function installNewOrdersEditTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === NEW_ORDERS_EDIT_TRIGGER_HANDLER) ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger(NEW_ORDERS_EDIT_TRIGGER_HANDLER).forSpreadsheet(ss).onEdit().create();
  PropertiesService.getScriptProperties().setProperty(NEW_ORDERS_EDIT_TRIGGER_INSTALLED_KEY, "true");
  showStyledAlert_("Edit Trigger Installed", "New Orders edits can now read external databases.", "success");
}

function handleNewOrdersEdit_(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  if (sheetName === SHEET_NAMES.STOCKCHECK_RTS) { try { handleStockcheckWorkflowEdit_(e); } catch (e) {} return; }
  if (sheetName === SHEET_NAMES.DISPATCHED) { try { handleDispatchedWorkflowEdit_(e); } catch (e) {} return; }
  if (sheetName === SHEET_NAMES.CREATE_LABEL) {
    var clStartCol = e.range.getColumn();
    var clEndCol = e.range.getLastColumn();
    if (isColumnInRange_(CREATE_LABEL_COLS.LabelCreated, clStartCol, clEndCol) && typeof labelMoveDoneRowsToDispatched_ === "function") {
      try { labelMoveDoneRowsToDispatched_(); } catch (e) {}
    }
    return;
  }

  if (sheetName === SHEET_NAMES.MATCHING_TABLE) {
    var mtStartCol = e.range.getColumn();
    var mtEndCol = e.range.getLastColumn();
    if (isColumnInRange_(MATCHING_COLS.VendorSKU, mtStartCol, mtEndCol) || isColumnInRange_(MATCHING_COLS.PortalSKU, mtStartCol, mtEndCol)) {
      refreshAllNewOrdersFromCatalog_(); refreshAllOrdersStockAvailability_();
      try { syncPendingOrdersToStockcheckRTS_(true); } catch (e) {}
    }
    return;
  }

  if (sheetName === SHEET_NAMES.INVENTORY) {
    var invStartRow = Math.max(2, e.range.getRow());
    var invEndRow = e.range.getLastRow();
    var invStartCol = e.range.getColumn();
    var invEndCol = e.range.getLastColumn();

    if (invEndRow >= 2 && isColumnInRange_(INVENTORY_COLS.IMAGE_URL, invStartCol, invEndCol)) {
      updateInventoryImageFormulas_(sheet, invStartRow, invEndRow);
    }
    if (isColumnInRange_(INVENTORY_COLS.VendorSKU, invStartCol, invEndCol) || isColumnInRange_(INVENTORY_COLS.Product_Cost, invStartCol, invEndCol) || isColumnInRange_(INVENTORY_COLS.Stock, invStartCol, invEndCol)) {
      refreshAllNewOrdersFromCatalog_(); refreshAllOrdersStockAvailability_();
      try { syncPendingOrdersToStockcheckRTS_(true); } catch (e) {}
    }
    return;
  }

  if (sheetName === SHEET_NAMES.ALL_ORDERS) {
    var aoStartCol = e.range.getColumn();
    var aoEndCol = e.range.getLastColumn();
    if (isColumnInRange_(HEADERS.SKU, aoStartCol, aoEndCol) || isColumnInRange_(HEADERS.QTY, aoStartCol, aoEndCol) || isColumnInRange_(HEADERS.STAFF_NOTES, aoStartCol, aoEndCol) || isColumnInRange_(HEADERS.ORDER_STATUS, aoStartCol, aoEndCol)) {
      refreshAllOrdersStockAvailability_();
      try { syncPendingOrdersToStockcheckRTS_(true); } catch (e) {}
    }
    return;
  }

  if (sheetName === SHEET_NAMES.CONFIG) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var newOrdersSheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
      if (newOrdersSheet && newOrdersSheet.getLastRow() >= DATA_START_ROW) {
        applyValidationsToRange_(newOrdersSheet, DATA_START_ROW, newOrdersSheet.getLastRow());
      }
      applyWorkflowValidations_();
    } catch (err) {}
    return;
  }

  if (sheetName !== SHEET_NAMES.NEW_ORDERS) return;

  var startRow = e.range.getRow();
  var endRow = e.range.getLastRow();
  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();

  if (endRow < DATA_START_ROW) return;
  if (startRow < DATA_START_ROW) startRow = DATA_START_ROW;

  var rows = [];
  for (var r = startRow; r <= endRow; r++) rows.push(r);

  var props = PropertiesService.getScriptProperties();
  var prevRowCount = Number(props.getProperty(ROW_COUNT_KEY)) || 0;
  var currentRowCount = sheet.getLastRow();

  if (prevRowCount > 0 && currentRowCount < prevRowCount) handleRowDeletion_(sheet);
  props.setProperty(ROW_COUNT_KEY, String(currentRowCount));

  var skuEdited = isColumnInRange_(HEADERS.SKU, startCol, endCol);
  var brandEdited = isColumnInRange_(HEADERS.BRAND_NAME, startCol, endCol);
  var qtyEdited = isColumnInRange_(HEADERS.QTY, startCol, endCol);
  var imageUrlEdited = isColumnInRange_(HEADERS.IMAGE_URL, startCol, endCol);
  var internalOrderNoEdited = isColumnInRange_(HEADERS.INTERNAL_ORDERNO, startCol, endCol);
  var currencyEdited = isColumnInRange_(HEADERS.CURRENCY, startCol, endCol);
  var portalIdEdited = isColumnInRange_(HEADERS.PORTAL_ORDER_ID, startCol, endCol);
  var countryEdited = isColumnInRange_(HEADERS.COUNTRY, startCol, endCol);
  var fullNameEdited = isColumnInRange_(HEADERS.FULLNAME, startCol, endCol);
  var address1Edited = isColumnInRange_(HEADERS.ADDRESSLINE1, startCol, endCol);
  var cityEdited = isColumnInRange_(HEADERS.CITY, startCol, endCol);
  var stateEdited = isColumnInRange_(HEADERS.STATE, startCol, endCol);
  var pincodeEdited = isColumnInRange_(HEADERS.PINCODE, startCol, endCol);
  var editedRowCount = rows.length;
  var editedColCount = endCol - startCol + 1;
  var isBulkEdit = editedRowCount > 5 || editedColCount > 4;

  refreshEditedRows_(sheet, startRow, endRow, true);

  try { normalizeDropdownValuesInRange_(sheet, startRow, endRow, startCol, endCol); } catch (e) {}

  try {
    if (portalIdEdited && fillBrandFromPortalId_(sheet, rows)) brandEdited = true;
  } catch (e) {}

  try {
    if (internalOrderNoEdited && editedRowCount === 1) resequenceFromManualEdit_(sheet, startRow);
  } catch (e) {}

  try {
    if (currencyEdited) sheet.getRange(startRow, HEADERS.Conversion_Rate, rows.length, 1).clearContent();
  } catch (e) {}

  try {
    if (imageUrlEdited) updateImageFormulas_(sheet, startRow, endRow);
    var skuRowsToRefresh = (skuEdited || editedRowCount > 1) ? rows.slice() : getRowsNeedingSkuRefresh_(sheet, startRow, endRow);
    if (skuRowsToRefresh.length > 0) {
      if (skuRowsToRefresh.length > 3) {
        var minRow = Math.min.apply(null, skuRowsToRefresh);
        var maxRow = Math.max.apply(null, skuRowsToRefresh);
        batchUpdateProductInfo_(sheet, minRow, maxRow, true);
      } else updateProductInfoForRows_(sheet, skuRowsToRefresh, true);
    }
  } catch (e) {}

  var internalIdGroupingEdited = brandEdited || portalIdEdited || countryEdited || fullNameEdited || address1Edited || cityEdited || stateEdited || pincodeEdited;
  try {
    if (internalIdGroupingEdited) {
      assignInternalOrderNos_(sheet, rows);
      if (!isBulkEdit) regenerateInternalOrderNos_(sheet, null);
    }
  } catch (e) {}

  try {
    var statusData = sheet.getRange(startRow, 1, rows.length, 30).getValues();
    var statusOut = [];
    var statusChanged = false;
    for (var i = 0; i < rows.length; i++) {
      var rawData = statusData[i];
      var hasAnyData = false;
      for (var c = 0; c < 30; c++) {
        if (String(rawData[c]).trim() !== "") { hasAnyData = true; break; }
      }
      var ds = String(rawData[HEADERS.ORDER_STATUS - 1] || "").trim();
      if (hasAnyData && ds === "") {
        statusOut.push(["Pending"]); statusChanged = true;
      } else {
        statusOut.push([rawData[HEADERS.ORDER_STATUS - 1]]);
      }
    }
    if (statusChanged) sheet.getRange(startRow, HEADERS.ORDER_STATUS, rows.length, 1).setValues(statusOut);
  } catch (e) {}

  refreshEditedRows_(sheet, startRow, endRow, false);
}

function refreshEditedRows_(sheet, startRow, endRow, skipFinancials) {
  if (!sheet || startRow > endRow) return;
  try { applyValidationsToRange_(sheet, startRow, endRow); } catch (e) {}
  if (!skipFinancials) {
    try {
      var rowCount = endRow - startRow + 1;
      if (rowCount > 3) batchCalculateFinancials_(sheet, startRow, endRow);
      else {
        var rows = [];
        for (var r = startRow; r <= endRow; r++) rows.push(r);
        calculateFinancialsForRows_(sheet, rows);
      }
    } catch (e) {}
  }
  try { validateAndLogErrors_(sheet, startRow, endRow); } catch (e) {}
}

function refreshAllNewOrdersFromCatalog_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  batchUpdateProductInfo_(sheet, DATA_START_ROW, lastRow, true);
  refreshEditedRows_(sheet, DATA_START_ROW, lastRow, false);
}

function findHeaderColByAliases_(sheet, headerRow, aliases) {
  if (!sheet) return 0;
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 0;
  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var aliasSet = {};
  for (var i = 0; i < aliases.length; i++) aliasSet[String(aliases[i]).trim().toLowerCase().replace(/\s+/g, "")] = true;
  for (var c = 0; c < headers.length; c++) {
    var norm = String(headers[c] || "").trim().toLowerCase().replace(/\s+/g, "");
    if (aliasSet[norm]) return c + 1;
  }
  return 0;
}

function refreshAllOrdersStockAvailability_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.ALL_ORDERS);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  var stockStatusCol = findHeaderColByAliases_(sheet, HEADER_ROW, ["stock", "stock_availibilty", "stock_availability", "stockavailability"]);
  if (!stockStatusCol) return;

  var numRows = lastRow - DATA_START_ROW + 1;
  var skuVals = sheet.getRange(DATA_START_ROW, HEADERS.SKU, numRows, 1).getValues();
  var qtyVals = sheet.getRange(DATA_START_ROW, HEADERS.QTY, numRows, 1).getValues();
  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var out = [], bg = [];

  for (var i = 0; i < numRows; i++) {
    var sku = String(skuVals[i][0]).trim();
    var qty = Number(qtyVals[i][0]) || 0;
    if (!sku || !portalMap[sku] || !invMap[portalMap[sku]]) {
      out.push([""]); bg.push([CLEAR_COLOR]);
      continue;
    }
    var status = getStockAvailabilityStatus_(invMap[portalMap[sku]].Stock, qty);
    out.push([status]);
    bg.push([status.indexOf("Out of stock") === 0 ? "#ffe0b2" : CLEAR_COLOR]);
  }

  var stockStatusRange = sheet.getRange(DATA_START_ROW, stockStatusCol, numRows, 1);
  stockStatusRange.setValues(out).setBackgrounds(bg).setWrap(true);
}

function normalizeWorkflowKey_(value) { return String(value == null ? "" : value).trim().toLowerCase().replace(/\s+/g, " "); }
function normalizeWorkflowHeaderKey_(value) { return String(value == null ? "" : value).trim().toLowerCase().replace(/[^a-z0-9]/g, ""); }

function getWorkflowHeaderMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = normalizeWorkflowHeaderKey_(headers[i]);
    if (key && !Object.prototype.hasOwnProperty.call(map, key)) map[key] = i + 1; 
  }
  return map;
}

function getHeaderColFromMap_(headerMap, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var key = normalizeWorkflowHeaderKey_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(headerMap, key)) return headerMap[key];
  }
  return 0;
}

function getAllOrdersBundle_() {
  var ss = getWorkflowSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAMES.ALL_ORDERS);
  if (!sheet) return null;

  var headerRow = detectHeaderRowByAliases_(sheet, ["PORTAL_ORDER_ID", "Internal OrderNo", "ORDER_STATUS"], 5);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < headerRow || lastCol < 1) return null;

  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var dataStart = headerRow + 1;
  var numRows = Math.max(0, lastRow - headerRow);
  var data = numRows > 0 ? sheet.getRange(dataStart, 1, numRows, lastCol).getValues() : [];

  return { sheet: sheet, headerRow: headerRow, headers: headers, headerMap: getWorkflowHeaderMap_(headers), dataStartRow: dataStart, data: data };
}

function getWorkflowSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active && (active.getSheetByName(SHEET_NAMES.ALL_ORDERS) || active.getSheetByName(SHEET_NAMES.STOCKCHECK_RTS) || active.getSheetByName(SHEET_NAMES.NEW_ORDERS))) {
    return active;
  }
  return SpreadsheetApp.openById(SHIPPING_TARGET_SPREADSHEET_ID);
}

function parseAvailableStockQty_(stockValue) {
  if (stockValue === null || stockValue === undefined || stockValue === "") return NaN;
  var asNumber = Number(stockValue);
  if (!isNaN(asNumber)) return asNumber;
  var text = String(stockValue).trim();
  if (!text) return NaN;
  var match = text.match(/available\s*:?\s*(-?\d+(?:\.\d+)?)/i);
  if (match) return Number(match[1]);
  var digits = text.match(/-?\d+(?:\.\d+)?/);
  if (digits) return Number(digits[0]);
  return NaN;
}

function buildInternalSkuKey_(internalOrderNo, sku) { return normalizeWorkflowKey_(internalOrderNo) + "||" + normalizeWorkflowKey_(sku); }
function buildQueueKey_(portalOrderId, internalOrderNo, sku) { return buildInternalSkuKey_(internalOrderNo, sku) + "||" + normalizeWorkflowKey_(portalOrderId); }

function getActiveInternalSkuKeysFromSheet_(sheet, internalCol, skuCol, startRow) {
  var keys = {};
  if (!sheet) return keys;
  var firstDataRow = startRow || WORKFLOW_DATA_START_ROW;
  var lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) return keys;

  var numRows = lastRow - firstDataRow + 1;
  var values = sheet.getRange(firstDataRow, 1, numRows, Math.max(internalCol, skuCol)).getValues();
  for (var i = 0; i < values.length; i++) {
    var key = buildInternalSkuKey_(values[i][internalCol - 1], values[i][skuCol - 1]);
    if (key !== "||") keys[key] = true;
  }
  return keys;
}

function getInventoryLocationByPortalSkuMap_() {
  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var locationByPortalSku = {};
  for (var portalSku in portalMap) {
    if (!Object.prototype.hasOwnProperty.call(portalMap, portalSku)) continue;
    var vendorSku = portalMap[portalSku];
    var inv = invMap[vendorSku];
    if (inv) locationByPortalSku[normalizeWorkflowKey_(portalSku)] = String(inv.Location || "").trim();
  }
  return locationByPortalSku;
}

function syncPendingOrdersToStockcheckRTS() {
  try {
    var result = syncPendingOrdersToStockcheckRTS_(false);
    showStyledAlert_("Queue Synced", "Added: " + result.added + "<br>Removed: " + result.removed + "<br>Eligible pending rows: " + result.eligible, "success");
  } catch (e) {
    showStyledAlert_("Queue Sync Failed", escapeHtml_(String(e && e.message ? e.message : e)), "error");
  }
}

function syncPendingOrdersToStockcheckRTS_(suppressUi) {
  var ss = getWorkflowSpreadsheet_();
  var stockcheckSheet = ss.getSheetByName(SHEET_NAMES.STOCKCHECK_RTS);
  if (!stockcheckSheet) throw new Error("stockcheckRTS sheet not found.");

  var allBundle = getAllOrdersBundle_();
  if (!allBundle) return { added: 0, removed: 0, eligible: 0 };

  var hm = allBundle.headerMap;
  var portalCol = getHeaderColFromMap_(hm, ["PORTAL_ORDER_ID", "PortalOrderID", "Portal Order ID"]);
  var orderStatusCol = getHeaderColFromMap_(hm, ["ORDER_STATUS", "Order Status"]);
  var internalCol = getHeaderColFromMap_(hm, ["INTERNAL_ORDERNO", "Internal OrderNo", "InternalOrderNo"]);
  var staffNotesCol = getHeaderColFromMap_(hm, ["STAFF_NOTES", "Staff Notes"]);
  var qtyCol = getHeaderColFromMap_(hm, ["QTY", "Qty", "Quantity"]);
  var imageCol = getHeaderColFromMap_(hm, ["IMAGE", "Image"]);
  var imageUrlCol = getHeaderColFromMap_(hm, ["IMAGE_URL", "Image URL", "Image Url"]);
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  var stockCol = getHeaderColFromMap_(hm, ["stock", "stock_availibilty", "stock_availability", "stockavailability"]);

  if (!portalCol || !orderStatusCol || !internalCol || !qtyCol || !skuCol) {
    throw new Error("All Orders header mapping is incomplete. Please verify headers.");
  }

  var locationByPortalSku = getInventoryLocationByPortalSkuMap_();
  var activeCreateLabelKeys = getActiveInternalSkuKeysFromSheet_(ss.getSheetByName(SHEET_NAMES.CREATE_LABEL), CREATE_LABEL_COLS.INTERNAL_ORDERNO, CREATE_LABEL_COLS.SKU, WORKFLOW_DATA_START_ROW);
  var activeDispatchedKeys = getActiveInternalSkuKeysFromSheet_(ss.getSheetByName(SHEET_NAMES.DISPATCHED), DISPATCHED_COLS.INTERNAL_ORDERNO, DISPATCHED_COLS.SKU, WORKFLOW_DATA_START_ROW);

  var eligibleByKey = {}, eligibleRows = [];
  for (var i = 0; i < allBundle.data.length; i++) {
    var row = allBundle.data[i];
    if (String(row[orderStatusCol - 1] || "").trim().toLowerCase() !== "pending") continue;

    var staffNotes = staffNotesCol ? String(row[staffNotesCol - 1] || "") : "";
    if (staffNotes.toLowerCase().indexOf("halted:") !== -1) continue;

    var portalOrderId = String(row[portalCol - 1] || "").trim();
    var internalOrderNo = String(row[internalCol - 1] || "").trim();
    var sku = String(row[skuCol - 1] || "").trim();
    var qty = Number(row[qtyCol - 1]) || 0;
    if (!portalOrderId || !internalOrderNo || !sku || qty <= 0) continue;

    var stockQty = stockCol ? parseAvailableStockQty_(row[stockCol - 1]) : NaN;
    if (!(stockQty >= qty && stockQty > 0)) continue;

    var internalSkuKey = buildInternalSkuKey_(internalOrderNo, sku);
    if (activeCreateLabelKeys[internalSkuKey] || activeDispatchedKeys[internalSkuKey]) continue;

    var queueKey = buildQueueKey_(portalOrderId, internalOrderNo, sku);
    if (eligibleByKey[queueKey]) continue;

    var location = locationByPortalSku[normalizeWorkflowKey_(sku)] || "";
    eligibleByKey[queueKey] = true;

    var staffNotesVal = sanitizeInput_(staffNotesCol ? row[staffNotesCol - 1] : "");
    var imageVal = sanitizeInput_(imageCol ? row[imageCol - 1] : "");
    var imageUrlVal = String(sanitizeInput_(imageUrlCol ? row[imageUrlCol - 1] : "")).trim();
    
    if ((!imageVal || imageVal === "") && imageUrlVal && imageUrlVal.indexOf("http") === 0) {
      imageVal = '=IMAGE("' + imageUrlVal.replace(/"/g, '""') + '")';
    }

    eligibleRows.push({
      key: queueKey,
      values: [portalOrderId, internalOrderNo, staffNotesVal, qty, imageVal, imageUrlVal, sku, location, "", ""],
    });
  }

  var lastRow = stockcheckSheet.getLastRow();
  var removed = 0, existingKeys = {}, rowsToDelete = [];
  
  if (lastRow >= WORKFLOW_DATA_START_ROW) {
    var numRows = lastRow - WORKFLOW_DATA_START_ROW + 1;
    var current = stockcheckSheet.getRange(WORKFLOW_DATA_START_ROW, 1, numRows, STOCKCHECK_COLS.DoCreateLabels).getValues();
    for (var r = 0; r < current.length; r++) {
      var cur = current[r];
      var key = buildQueueKey_(cur[STOCKCHECK_COLS.PORTAL_ORDER_ID - 1], cur[STOCKCHECK_COLS.INTERNAL_ORDERNO - 1], cur[STOCKCHECK_COLS.SKU - 1]);
      if (key === "||||") continue;
      if (eligibleByKey[key]) existingKeys[key] = true;
      else rowsToDelete.push(WORKFLOW_DATA_START_ROW + r);
    }
  }

  if (rowsToDelete.length > 0) { batchDeleteRows_(stockcheckSheet, rowsToDelete); removed += rowsToDelete.length; }

  var rowsToAppend = [];
  for (var j = 0; j < eligibleRows.length; j++) {
    if (!existingKeys[eligibleRows[j].key]) rowsToAppend.push(eligibleRows[j].values);
  }

  if (rowsToAppend.length > 0) {
    var appendStart = Math.max(stockcheckSheet.getLastRow() + 1, WORKFLOW_DATA_START_ROW);
    stockcheckSheet.getRange(appendStart, 1, rowsToAppend.length, STOCKCHECK_COLS.DoCreateLabels).setValues(rowsToAppend);
  }

  refreshPickListSheetIfAvailable_(true);

  if (!suppressUi && rowsToAppend.length === 0 && removed === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("stockcheckRTS is already up to date.", "Queue Sync", 5);
  }

  return { added: rowsToAppend.length, removed: removed, eligible: eligibleRows.length };
}

function refreshPickListSheetIfAvailable_(suppressUi) {
  if (typeof refreshPickListFromStockcheckRTS_ !== "function") return;
  try { refreshPickListFromStockcheckRTS_(suppressUi); } catch (e) {}
}

function normalizeWorkflowChoice_(value, choices) {
  var raw = String(value == null ? "" : value).trim();
  if (!raw) return "";
  for (var i = 0; i < choices.length; i++) {
    if (raw.toLowerCase() === choices[i].toLowerCase()) return choices[i];
  }
  return raw;
}

function batchDeleteRows_(sheet, rowsToDelete) {
  if (!sheet || !rowsToDelete || rowsToDelete.length === 0) return;
  var uniqueRows = [], seen = {};
  for (var i = 0; i < rowsToDelete.length; i++) {
    var r = Number(rowsToDelete[i]);
    if (r > 0 && !seen[r]) { uniqueRows.push(r); seen[r] = true; }
  }
  uniqueRows.sort(function (a, b) { return b - a; });
  if (uniqueRows.length === 0) return;

  var start = uniqueRows[0], count = 1;
  for (var j = 1; j < uniqueRows.length; j++) {
    if (uniqueRows[j] === start - count) count++;
    else { sheet.deleteRows(start - count + 1, count); start = uniqueRows[j]; count = 1; }
  }
  sheet.deleteRows(start - count + 1, count);
}

function appendHaltReason_(currentValue, reason) {
  var trimmed = String(currentValue == null ? "" : currentValue).trim();
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "dd-MMM-yyyy HH:mm");
  var entry = "Halted: " + reason + " (" + stamp + ")";
  if (!trimmed) return entry;
  if (trimmed.toLowerCase().indexOf(("halted: " + reason).toLowerCase()) !== -1) return trimmed;
  return trimmed + " | " + entry;
}

function updateAllOrdersForHalt_(internalOrderNo, sku, reason, optionalBundle) {
  var bundle = optionalBundle || getAllOrdersBundle_();
  if (!bundle) return 0;

  var hm = bundle.headerMap;
  var internalCol = getHeaderColFromMap_(hm, ["INTERNAL_ORDERNO", "Internal OrderNo", "InternalOrderNo"]);
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  var orderStatusCol = getHeaderColFromMap_(hm, ["ORDER_STATUS", "Order Status"]);
  var staffNotesCol = getHeaderColFromMap_(hm, ["STAFF_NOTES", "Staff Notes"]);
  if (!internalCol || !orderStatusCol || !staffNotesCol) return 0;

  var targetInternal = normalizeWorkflowKey_(internalOrderNo);
  var targetSku = normalizeWorkflowKey_(sku);
  var touched = 0;
  for (var i = 0; i < bundle.data.length; i++) {
    var rowInternal = normalizeWorkflowKey_(bundle.data[i][internalCol - 1]);
    if (!rowInternal || rowInternal !== targetInternal) continue;
    if (targetSku && skuCol) {
      var rowSku = normalizeWorkflowKey_(bundle.data[i][skuCol - 1]);
      if (rowSku !== targetSku) continue;
    }

    var absRow = bundle.dataStartRow + i;
    bundle.sheet.getRange(absRow, orderStatusCol).setValue("Pending");
    var notesCell = bundle.sheet.getRange(absRow, staffNotesCol);
    notesCell.setValue(appendHaltReason_(notesCell.getValue(), reason));
    notesCell.setBackground("#fff3cd").setFontColor("#b71c1c").setFontWeight("bold");
    touched++;
  }
  return touched;
}

function removeRowsFromQueueByInternalSku_(sheet, internalCol, skuCol, internalOrderNo, sku) {
  if (!sheet) return 0;
  var targetKey = buildInternalSkuKey_(internalOrderNo, sku);
  var lastRow = sheet.getLastRow();
  if (lastRow < WORKFLOW_DATA_START_ROW) return 0;

  var rows = sheet.getRange(WORKFLOW_DATA_START_ROW, 1, lastRow - WORKFLOW_DATA_START_ROW + 1, Math.max(internalCol, skuCol)).getValues();
  var rowsToDelete = [];
  for (var i = 0; i < rows.length; i++) {
    var key = buildInternalSkuKey_(rows[i][internalCol - 1], rows[i][skuCol - 1]);
    if (key === targetKey) rowsToDelete.push(WORKFLOW_DATA_START_ROW + i);
  }
  batchDeleteRows_(sheet, rowsToDelete);
  return rowsToDelete.length;
}

function findAllOrderAddressRow_(internalOrderNo, sku, optionalBundle) {
  var bundle = optionalBundle || getAllOrdersBundle_();
  if (!bundle) return null;
  var hm = bundle.headerMap;
  var internalCol = getHeaderColFromMap_(hm, ["INTERNAL_ORDERNO", "Internal OrderNo", "InternalOrderNo"]);
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  if (!internalCol || !skuCol) return null;

  var targetInternal = normalizeWorkflowKey_(internalOrderNo);
  var targetSku = normalizeWorkflowKey_(sku);

  for (var i = 0; i < bundle.data.length; i++) {
    var r = bundle.data[i];
    if (normalizeWorkflowKey_(r[internalCol - 1]) !== targetInternal) continue;
    if (normalizeWorkflowKey_(r[skuCol - 1]) !== targetSku) continue;
    return { row: r, headerMap: hm };
  }

  for (var j = 0; j < bundle.data.length; j++) {
    var fallback = bundle.data[j];
    if (normalizeWorkflowKey_(fallback[internalCol - 1]) === targetInternal) {
      return { row: fallback, headerMap: hm };
    }
  }
  return null;
}

function moveStockcheckRowToCreateLabel_(stockcheckSheet, rowNumber, rowValues) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var createSheet = ss.getSheetByName(SHEET_NAMES.CREATE_LABEL);
  if (!createSheet) throw new Error("CreateLabels sheet not found.");

  var internalOrderNo = rowValues[STOCKCHECK_COLS.INTERNAL_ORDERNO - 1];
  var sku = rowValues[STOCKCHECK_COLS.SKU - 1];
  var addressSource = findAllOrderAddressRow_(internalOrderNo, sku);
  var headerMap = addressSource ? addressSource.headerMap : {};
  var sourceRow = addressSource ? addressSource.row : [];

  function fromAllOrders(aliases) {
    var col = getHeaderColFromMap_(headerMap, aliases);
    if (!col) return "";
    return sourceRow[col - 1];
  }

  var countryVal = String(fromAllOrders(["COUNTRY", "Country"]) || "").trim();
  var key = buildInternalSkuKey_(internalOrderNo, sku);

  if (countryVal && countryVal.toLowerCase() === "india") {
    var dispatchedSheet = ss.getSheetByName(SHEET_NAMES.DISPATCHED);
    if (!dispatchedSheet) throw new Error("Dispached sheet not found.");

    var dispatchPayload = [[internalOrderNo, sku, fromAllOrders(["FULLNAME", "Fullname", "Full Name"]), fromAllOrders(["ADDRESSLINE1", "AddressLine1"]), fromAllOrders(["ADDRESSLINE2", "AddressLine2"]), fromAllOrders(["CITY", "City"]), fromAllOrders(["STATE", "State"]), fromAllOrders(["PINCODE", "Pincode"]), fromAllOrders(["COUNTRY", "Country"]), fromAllOrders(["PHONE", "Phone"]), "Ready to Ship"]];

    var existingDispatch = getActiveInternalSkuKeysFromSheet_(dispatchedSheet, DISPATCHED_COLS.INTERNAL_ORDERNO, DISPATCHED_COLS.SKU, WORKFLOW_DATA_START_ROW);

    if (!existingDispatch[key]) {
      var appendRow = Math.max(dispatchedSheet.getLastRow() + 1, WORKFLOW_DATA_START_ROW);
      dispatchedSheet.getRange(appendRow, 1, 1, DISPATCHED_COLS.DISPACHED_STATUS).setValues(dispatchPayload);
    }

    try { updateAllOrdersStatusByInternal_(internalOrderNo, sku, "Ready to Ship"); } catch (e) {}
    removeRowsFromQueueByInternalSku_(ss.getSheetByName(SHEET_NAMES.CREATE_LABEL), CREATE_LABEL_COLS.INTERNAL_ORDERNO, CREATE_LABEL_COLS.SKU, internalOrderNo, sku);
    removeRowsFromQueueByInternalSku_(ss.getSheetByName(SHEET_NAMES.PICKLIST), 2, 7, internalOrderNo, sku);

    stockcheckSheet.deleteRow(rowNumber);
    return;
  }

  var payload = [[internalOrderNo, sku, fromAllOrders(["FULLNAME", "Fullname", "Full Name"]), fromAllOrders(["ADDRESSLINE1", "AddressLine1"]), fromAllOrders(["ADDRESSLINE2", "AddressLine2"]), fromAllOrders(["CITY", "City"]), fromAllOrders(["STATE", "State"]), fromAllOrders(["PINCODE", "Pincode"]), fromAllOrders(["COUNTRY", "Country"]), fromAllOrders(["PHONE", "Phone"]), "No"]];

  var existing = getActiveInternalSkuKeysFromSheet_(createSheet, CREATE_LABEL_COLS.INTERNAL_ORDERNO, CREATE_LABEL_COLS.SKU, WORKFLOW_DATA_START_ROW);

  if (!existing[key]) {
    var appendRow = Math.max(createSheet.getLastRow() + 1, WORKFLOW_DATA_START_ROW);
    createSheet.getRange(appendRow, 1, 1, CREATE_LABEL_COLS.LabelCreated).setValues(payload);
  }

  stockcheckSheet.deleteRow(rowNumber);
}

function haltStockcheckRow_(stockcheckSheet, rowNumber, rowValues, reason) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var internalOrderNo = rowValues[STOCKCHECK_COLS.INTERNAL_ORDERNO - 1];
  var sku = rowValues[STOCKCHECK_COLS.SKU - 1];

  updateAllOrdersForHalt_(internalOrderNo, sku, reason);
  stockcheckSheet.deleteRow(rowNumber);

  removeRowsFromQueueByInternalSku_(ss.getSheetByName(SHEET_NAMES.CREATE_LABEL), CREATE_LABEL_COLS.INTERNAL_ORDERNO, CREATE_LABEL_COLS.SKU, internalOrderNo, sku);
  removeRowsFromQueueByInternalSku_(ss.getSheetByName(SHEET_NAMES.PICKLIST), 2, 7, internalOrderNo, sku);
  removeRowsFromQueueByInternalSku_(ss.getSheetByName(SHEET_NAMES.DISPATCHED), DISPATCHED_COLS.INTERNAL_ORDERNO, DISPATCHED_COLS.SKU, internalOrderNo, sku);
}

function handleStockcheckWorkflowEdit_(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.STOCKCHECK_RTS) return;

  var startRow = Math.max(e.range.getRow(), WORKFLOW_DATA_START_ROW);
  var endRow = e.range.getLastRow();
  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();

  if (endRow < WORKFLOW_DATA_START_ROW) return;
  if (!isColumnInRange_(STOCKCHECK_COLS.ActualStockStatus, startCol, endCol) && !isColumnInRange_(STOCKCHECK_COLS.DoCreateLabels, startCol, endCol)) {
    refreshPickListSheetIfAvailable_(true); return;
  }

  var numRows = endRow - startRow + 1;
  var rangeData = sheet.getRange(startRow, 1, numRows, STOCKCHECK_COLS.DoCreateLabels).getValues();

  var actualOpts = getConfigValues_(["ActualStockStatus", "Actual Stock Status"]);
  if (!actualOpts || actualOpts.length === 0) actualOpts = ["Not Found", "Faulty", "OK"];

  var labelOpts = getConfigValues_(["DoCreateLabels", "Do Create Labels"]);
  if (!labelOpts || labelOpts.length === 0) labelOpts = ["Yes", "No"];

  var updates = [], rowsToProcess = [];

  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    var values = rangeData[i];
    var rawActual = values[STOCKCHECK_COLS.ActualStockStatus - 1];
    var rawDoCreate = values[STOCKCHECK_COLS.DoCreateLabels - 1];

    var actual = normalizeWorkflowChoice_(rawActual, actualOpts);
    var doCreate = normalizeWorkflowChoice_(rawDoCreate, labelOpts);

    var changed = false;
    if (rawActual !== actual) { values[STOCKCHECK_COLS.ActualStockStatus - 1] = actual; changed = true; }
    if (rawDoCreate !== doCreate) { values[STOCKCHECK_COLS.DoCreateLabels - 1] = doCreate; changed = true; }

    if (changed) updates.push({ row: row, actual: actual, doCreate: doCreate });

    if (actual === "OK" && doCreate === "Yes") rowsToProcess.push({ action: "move", values: values, row: row });
    else if (actual === "Faulty") rowsToProcess.push({ action: "halt", reason: "Faulty", values: values, row: row });
    else if (actual === "Not Found") rowsToProcess.push({ action: "halt", reason: "Stock Run Out", values: values, row: row });
  }

  for (var u = 0; u < updates.length; u++) {
    sheet.getRange(updates[u].row, STOCKCHECK_COLS.ActualStockStatus).setValue(updates[u].actual);
    sheet.getRange(updates[u].row, STOCKCHECK_COLS.DoCreateLabels).setValue(updates[u].doCreate);
  }

  if (rowsToProcess.length === 0) { refreshPickListSheetIfAvailable_(true); return; }
  batchProcessStockcheckRows_(sheet, rowsToProcess);

  try { syncPendingOrdersToStockcheckRTS_(true); } catch (e) {}
}

function batchProcessStockcheckRows_(stockcheckSheet, rowsToProcess) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bundle = getAllOrdersBundle_();
  if (!bundle) return;

  var createSheet = ss.getSheetByName(SHEET_NAMES.CREATE_LABEL);
  var dispatchedSheet = ss.getSheetByName(SHEET_NAMES.DISPATCHED);
  var picklistSheet = ss.getSheetByName(SHEET_NAMES.PICKLIST);

  var existingCreate = getActiveInternalSkuKeysFromSheet_(createSheet, CREATE_LABEL_COLS.INTERNAL_ORDERNO, CREATE_LABEL_COLS.SKU, WORKFLOW_DATA_START_ROW);
  var existingDispatch = getActiveInternalSkuKeysFromSheet_(dispatchedSheet, DISPATCHED_COLS.INTERNAL_ORDERNO, DISPATCHED_COLS.SKU, WORKFLOW_DATA_START_ROW);

  var createPayloads = [], dispatchPayloads = [], halts = [], readyToShipUpdates = [];

  for (var i = 0; i < rowsToProcess.length; i++) {
    var p = rowsToProcess[i];
    var internalOrderNo = p.values[STOCKCHECK_COLS.INTERNAL_ORDERNO - 1];
    var sku = p.values[STOCKCHECK_COLS.SKU - 1];
    var key = buildInternalSkuKey_(internalOrderNo, sku);

    if (p.action === "move") {
      var addressSource = findAllOrderAddressRow_(internalOrderNo, sku, bundle);
      var sourceRow = addressSource ? addressSource.row : [];

      var getField = function (aliases) {
        var c = getHeaderColFromMap_(addressSource ? addressSource.headerMap : bundle.headerMap, aliases);
        return c ? sourceRow[c - 1] : "";
      };

      var countryVal = String(getField(["COUNTRY", "Country"]) || "").trim();
      var isIndia = countryVal.toLowerCase() === "india";

      var payload = [internalOrderNo, sku, getField(["FULLNAME", "Fullname", "Full Name"]), getField(["ADDRESSLINE1", "AddressLine1"]), getField(["ADDRESSLINE2", "AddressLine2"]), getField(["CITY", "City"]), getField(["STATE", "State"]), getField(["PINCODE", "Pincode"]), getField(["COUNTRY", "Country"]), getField(["PHONE", "Phone"]), isIndia ? "Ready to Ship" : "No"];

      if (isIndia) {
        if (!existingDispatch[key]) { dispatchPayloads.push(payload); existingDispatch[key] = true; }
        readyToShipUpdates.push({ internalOrderNo: internalOrderNo, sku: sku });
        halts.push({ internalOrderNo: internalOrderNo, sku: sku, queuesOnly: true });
      } else {
        if (!existingCreate[key]) { createPayloads.push(payload); existingCreate[key] = true; }
      }
    } else if (p.action === "halt") {
      halts.push({ internalOrderNo: internalOrderNo, sku: sku, reason: p.reason });
    }
  }

  if (createPayloads.length > 0 && createSheet) {
    var appRow1 = Math.max(createSheet.getLastRow() + 1, WORKFLOW_DATA_START_ROW);
    createSheet.getRange(appRow1, 1, createPayloads.length, createPayloads[0].length).setValues(createPayloads);
  }
  if (dispatchPayloads.length > 0 && dispatchedSheet) {
    var appRow2 = Math.max(dispatchedSheet.getLastRow() + 1, WORKFLOW_DATA_START_ROW);
    dispatchedSheet.getRange(appRow2, 1, dispatchPayloads.length, dispatchPayloads[0].length).setValues(dispatchPayloads);
  }

  if (readyToShipUpdates.length > 0) {
    for (var r = 0; r < readyToShipUpdates.length; r++) updateAllOrdersStatusByInternal_(readyToShipUpdates[r].internalOrderNo, readyToShipUpdates[r].sku, "Ready to Ship", bundle);
  }

  for (var h = 0; h < halts.length; h++) {
    var hRec = halts[h];
    if (!hRec.queuesOnly) updateAllOrdersForHalt_(hRec.internalOrderNo, hRec.sku, hRec.reason, bundle);
    if (createSheet) removeRowsFromQueueByInternalSku_(createSheet, CREATE_LABEL_COLS.INTERNAL_ORDERNO, CREATE_LABEL_COLS.SKU, hRec.internalOrderNo, hRec.sku);
    if (picklistSheet) removeRowsFromQueueByInternalSku_(picklistSheet, 2, 7, hRec.internalOrderNo, hRec.sku);
    if (!hRec.queuesOnly && dispatchedSheet) removeRowsFromQueueByInternalSku_(dispatchedSheet, DISPATCHED_COLS.INTERNAL_ORDERNO, DISPATCHED_COLS.SKU, hRec.internalOrderNo, hRec.sku);
  }
}

function updateAllOrdersStatusByInternal_(internalOrderNo, sku, newStatus, optionalBundle) {
  var bundle = optionalBundle || getAllOrdersBundle_();
  if (!bundle) return 0;

  var hm = bundle.headerMap;
  var internalCol = getHeaderColFromMap_(hm, ["INTERNAL_ORDERNO", "Internal OrderNo", "InternalOrderNo"]);
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  var orderStatusCol = getHeaderColFromMap_(hm, ["ORDER_STATUS", "Order Status"]);
  if (!internalCol || !orderStatusCol) return 0;

  var targetInternal = normalizeWorkflowKey_(internalOrderNo);
  var targetSku = normalizeWorkflowKey_(sku);
  var count = 0;
  var background = getWorkflowStatusBackground_(newStatus);
  for (var i = 0; i < bundle.data.length; i++) {
    var row = bundle.data[i];
    if (normalizeWorkflowKey_(row[internalCol - 1]) !== targetInternal) continue;
    if (targetSku && skuCol && normalizeWorkflowKey_(row[skuCol - 1]) !== targetSku) continue;

    var absRow = bundle.dataStartRow + i;
    bundle.sheet.getRange(absRow, orderStatusCol).setValue(newStatus);
    if (background) bundle.sheet.getRange(absRow, 1, 1, bundle.headers.length).setBackground(background);
    count++;
  }
  return count;
}

function getWorkflowStatusBackground_(status) {
  var normalized = normalizeWorkflowKey_(status);
  if (normalized === "dispatched") return "#d9ead3";
  if (normalized === "cancel") return "#fdecea";
  if (normalized === "ready to ship") return "#fff2cc";
  return "";
}

function updateAllOrdersShippingStatusByInternal_(internalOrderNo, sku, newShippingStatus, optionalBundle) {
  var bundle = optionalBundle || getAllOrdersBundle_();
  if (!bundle) return 0;

  var hm = bundle.headerMap;
  var internalCol = getHeaderColFromMap_(hm, ["INTERNAL_ORDERNO", "Internal OrderNo", "InternalOrderNo"]);
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  var shippingCol = getHeaderColFromMap_(hm, ["STATUS", "Shipping Details Status", "Shipping Details", "Status"]);
  if (!internalCol || !shippingCol) return 0;

  var targetInternal = normalizeWorkflowKey_(internalOrderNo);
  var targetSku = normalizeWorkflowKey_(sku);
  var count = 0;
  for (var i = 0; i < bundle.data.length; i++) {
    var row = bundle.data[i];
    if (normalizeWorkflowKey_(row[internalCol - 1]) !== targetInternal) continue;
    if (targetSku && skuCol && normalizeWorkflowKey_(row[skuCol - 1]) !== targetSku) continue;

    var absRow = bundle.dataStartRow + i;
    bundle.sheet.getRange(absRow, shippingCol).setValue(newShippingStatus);
    count++;
  }
  return count;
}

function handleDispatchedWorkflowEdit_(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.DISPATCHED) return;

  var startRow = Math.max(e.range.getRow(), WORKFLOW_DATA_START_ROW);
  var endRow = e.range.getLastRow();
  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();

  if (endRow < WORKFLOW_DATA_START_ROW) return;
  if (!isColumnInRange_(DISPATCHED_COLS.DISPACHED_STATUS, startCol, endCol)) return;

  var needsSync = false;
  var numRows = endRow - startRow + 1;

  var statuses = sheet.getRange(startRow, DISPATCHED_COLS.DISPACHED_STATUS, numRows, 1).getValues();
  var internalNos = sheet.getRange(startRow, DISPATCHED_COLS.INTERNAL_ORDERNO, numRows, 1).getValues();
  var skus = sheet.getRange(startRow, DISPATCHED_COLS.SKU, numRows, 1).getValues();

  var statusUpdates = [], bgUpdates = [], allOrdersUpdates = [];
  var dispatchOpts = getConfigValues_(["DISPACHED_STATUS", "Dispatched Status", "DispatchedStatus", "DISPACHED_STAUTS", "Dispatched Stauts"]);
  if (!dispatchOpts || dispatchOpts.length === 0) dispatchOpts = ["Ready to Ship", "Pending", "Dispatched", "Cancel"];

  var rowsToDelete = [];
  for (var i = 0; i < numRows; i++) {
    var rawStatus = statuses[i][0];
    var status = normalizeWorkflowChoice_(rawStatus, dispatchOpts);
    statusUpdates.push([status]);

    if (status === "Dispatched") {
      bgUpdates.push(["#d9ead3"]);
      allOrdersUpdates.push({ internalOrderNo: internalNos[i][0], sku: skus[i][0], orderStatus: "Dispatched", shippingStatus: "In Transit" });
      rowsToDelete.push(startRow + i);
      needsSync = true;
    } else if (status === "Cancel") {
      bgUpdates.push(["#fdecea"]); needsSync = true;
    } else if (status === "Pending") {
      bgUpdates.push(["#fff9c4"]);
    } else {
      bgUpdates.push(["#ffffff"]);
    }
  }

  sheet.getRange(startRow, DISPATCHED_COLS.DISPACHED_STATUS, numRows, 1).setValues(statusUpdates).setBackgrounds(bgUpdates);

  if (allOrdersUpdates.length > 0) {
    var bundle = getAllOrdersBundle_();
    if (bundle) {
      for (var u = 0; u < allOrdersUpdates.length; u++) {
        var upd = allOrdersUpdates[u];
        if (upd.orderStatus) updateAllOrdersStatusByInternal_(upd.internalOrderNo, upd.sku, upd.orderStatus, bundle);
        if (upd.shippingStatus) updateAllOrdersShippingStatusByInternal_(upd.internalOrderNo, upd.sku, upd.shippingStatus, bundle);
      }
    }
  }

  if (rowsToDelete.length > 0) batchDeleteRows_(sheet, rowsToDelete);
  if (needsSync) { try { syncPendingOrdersToStockcheckRTS_(true); } catch (e) {} }
}

function applyWorkflowValidations_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stockcheckSheet = ss.getSheetByName(SHEET_NAMES.STOCKCHECK_RTS);
  var dispatchedSheet = ss.getSheetByName(SHEET_NAMES.DISPATCHED);

  var actualRange = getConfigColRange_(["ActualStockStatus", "Actual Stock Status"]);
  var labelRange = getConfigColRange_(["DoCreateLabels", "Do Create Labels"]);
  var dispatchRange = getConfigColRange_(["DISPACHED_STATUS", "Dispatched Status", "DispatchedStatus", "DISPACHED_STAUTS", "Dispatched Stauts"]);

  if (stockcheckSheet) {
    var actualRule = actualRange ? SpreadsheetApp.newDataValidation().requireValueInRange(actualRange, true).setAllowInvalid(false).build() : SpreadsheetApp.newDataValidation().requireValueInList(["Not Found", "Faulty", "OK"], true).setAllowInvalid(false).build();
    var labelRule = labelRange ? SpreadsheetApp.newDataValidation().requireValueInRange(labelRange, true).setAllowInvalid(false).build() : SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).setAllowInvalid(false).build();
    var stockRows = Math.max(1, stockcheckSheet.getMaxRows() - WORKFLOW_HEADER_ROW);
    stockcheckSheet.getRange(WORKFLOW_DATA_START_ROW, STOCKCHECK_COLS.ActualStockStatus, stockRows, 1).setDataValidation(actualRule);
    stockcheckSheet.getRange(WORKFLOW_DATA_START_ROW, STOCKCHECK_COLS.DoCreateLabels, stockRows, 1).setDataValidation(labelRule);
  }

  if (dispatchedSheet) {
    var dispatchedRule = dispatchRange ? SpreadsheetApp.newDataValidation().requireValueInRange(dispatchRange, true).setAllowInvalid(false).build() : SpreadsheetApp.newDataValidation().requireValueInList(["Ready to Ship", "Pending", "Dispatched", "Cancel"], true).setAllowInvalid(false).build();
    var dispatchRows = Math.max(1, dispatchedSheet.getMaxRows() - WORKFLOW_HEADER_ROW);
    dispatchedSheet.getRange(WORKFLOW_DATA_START_ROW, DISPATCHED_COLS.DISPACHED_STATUS, dispatchRows, 1).setDataValidation(dispatchedRule);
  }
}

function applyCreateLabelDefaults_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CREATE_LABEL);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < WORKFLOW_DATA_START_ROW) return 0;

  var numRows = lastRow - WORKFLOW_DATA_START_ROW + 1;
  var vals = sheet.getRange(WORKFLOW_DATA_START_ROW, CREATE_LABEL_COLS.LabelCreated, numRows, 1).getValues();
  var changed = 0;
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0] || "").trim();
    if (!v) { vals[i][0] = "No"; changed++; }
  }
  if (changed > 0) sheet.getRange(WORKFLOW_DATA_START_ROW, CREATE_LABEL_COLS.LabelCreated, numRows, 1).setValues(vals);
  return changed;
}

function applyDispatchedDefaults_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DISPATCHED);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < WORKFLOW_DATA_START_ROW) return 0;

  var numRows = lastRow - WORKFLOW_DATA_START_ROW + 1;
  var vals = sheet.getRange(WORKFLOW_DATA_START_ROW, DISPATCHED_COLS.DISPACHED_STATUS, numRows, 1).getValues();
  var changed = 0;
  for (var i = 0; i < vals.length; i++) {
    var v = String(vals[i][0] || "").trim();
    if (!v) { vals[i][0] = "Ready to Ship"; changed++; }
  }
  if (changed > 0) sheet.getRange(WORKFLOW_DATA_START_ROW, DISPATCHED_COLS.DISPACHED_STATUS, numRows, 1).setValues(vals);
  return changed;
}

function renameDeliveryStatusHeadersToOrderStatus_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var renamedCells = 0;

  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var maxRows = Math.min(5, sheet.getLastRow());
    var maxCols = sheet.getLastColumn();
    if (maxRows < 1 || maxCols < 1) continue;

    var range = sheet.getRange(1, 1, maxRows, maxCols);
    var values = range.getValues();
    var changed = false;

    for (var r = 0; r < values.length; r++) {
      for (var c = 0; c < values[r].length; c++) {
        var txt = String(values[r][c] || "").trim();
        var normalized = txt.toUpperCase().replace(/\s+/g, " ");
        if (normalized === "DELIVERY STATUS" || normalized === "DELIVERY_STATUS") {
          values[r][c] = "ORDER_STATUS"; renamedCells++; changed = true;
        }
      }
    }
    if (changed) range.setValues(values);
  }
  return renamedCells;
}

function setupSystem() {
  try {
    var renamedHeaders = renameDeliveryStatusHeadersToOrderStatus_();
    applyWorkflowValidations_();
    var labelDefaults = applyCreateLabelDefaults_();
    var dispatchedDefaults = applyDispatchedDefaults_();
    var syncStats = syncPendingOrdersToStockcheckRTS_(true);

    SpreadsheetApp.getActiveSpreadsheet().toast("Setup completed. Added: " + syncStats.added + ", Removed: " + syncStats.removed, "Setup System", 8);

    showStyledAlert_("System Setup Complete", "Header updates (legacy status -> ORDER_STATUS): " + renamedHeaders + "<br>Workflow validations applied.<br>Label defaults set: " + labelDefaults + "<br>Dispatched defaults set: " + dispatchedDefaults + "<br>stockcheckRTS sync added: " + syncStats.added + ", removed: " + syncStats.removed, "success");
  } catch (e) {
    showStyledAlert_("Setup Failed", escapeHtml_(String(e && e.message ? e.message : e)), "error");
  }
}

function buildAllOrdersRowFromNewOrderRow_(newOrderRow, portalMap, invMap, allOrdersColCount, stockStatusCol) {
  var out = [];
  for (var i = 0; i < allOrdersColCount; i++) out.push("");

  var copyCount = Math.min(40, newOrderRow.length, allOrdersColCount);
  for (var c = 0; c < copyCount; c++) {
    var val = newOrderRow[c];
    if (val !== null && typeof val === "object" && Object.prototype.toString.call(val) !== "[object Date]") val = "";
    out[c] = val;
  }

  var imageUrl = String(newOrderRow[HEADERS.IMAGE_URL - 1] || "").trim();
  if (imageUrl && imageUrl.indexOf("http") === 0) out[HEADERS.IMAGE - 1] = '=IMAGE("' + imageUrl.replace(/"/g, '""') + '")';
  else out[HEADERS.IMAGE - 1] = ""; 

  if (stockStatusCol > 0) {
    var sku = String(newOrderRow[HEADERS.SKU - 1] || "").trim();
    var qty = Number(newOrderRow[HEADERS.QTY - 1]) || 0;
    var status = "";
    if (sku) {
      var vendor = portalMap[sku];
      if (vendor && invMap[vendor]) status = getStockAvailabilityStatus_(invMap[vendor].Stock, qty);
    }
    out[stockStatusCol - 1] = status;
  }
  return out;
}

function normalizeHeaderKey_(value) { return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]/g, ""); }

function detectHeaderRowByAliases_(sheet, aliases, maxRowsToScan) {
  if (!sheet) return 1;
  var maxRows = Math.min(maxRowsToScan || 5, sheet.getLastRow());
  var maxCols = sheet.getLastColumn();
  if (maxRows < 1 || maxCols < 1) return 1;
  var normalizedAliases = {};
  for (var i = 0; i < aliases.length; i++) normalizedAliases[normalizeHeaderKey_(aliases[i])] = true;
  for (var r = 1; r <= maxRows; r++) {
    var row = sheet.getRange(r, 1, 1, maxCols).getValues()[0];
    for (var c = 0; c < row.length; c++) {
      if (normalizedAliases[normalizeHeaderKey_(row[c])]) return r;
    }
  }
  return 1;
}

function valueByAnyKey_(map, keys) {
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (Object.prototype.hasOwnProperty.call(map, k)) return map[k];
  }
  return "";
}

function resolveReadyFieldFromAllOrders_(destKey, sourceMap) {
  if (destKey === "orderstatus") return "Ready to Ship";
  if (Object.prototype.hasOwnProperty.call(sourceMap, destKey)) return sourceMap[destKey];

  var aliasMap = { deliveryorpickupdate: ["deliverypickupdate"], cateogry: ["category"], inr: ["conversionrate"], ecommerceexpensetaxes: ["actualexpense", "maximumexpense"], estimatedprofit: ["actualprofit", "maximumprofit"] };
  if (aliasMap[destKey]) return valueByAnyKey_(sourceMap, aliasMap[destKey]);

  if (destKey === "returnrequestdate" || destKey === "returncouriername" || destKey === "returntrackingcode" || destKey === "rtorecieveddate" || destKey === "wrongreturnclaims") return "";
  return "";
}

function buildReadyToShipRowFromAllOrders_(allHeaders, allOrderRow, readyHeaders) {
  var sourceMap = {};
  for (var i = 0; i < allHeaders.length; i++) sourceMap[normalizeHeaderKey_(allHeaders[i])] = allOrderRow[i];
  var out = [];
  for (var c = 0; c < readyHeaders.length; c++) {
    var destKey = normalizeHeaderKey_(readyHeaders[c]);
    out.push(resolveReadyFieldFromAllOrders_(destKey, sourceMap));
  }
  return out;
}

function getOrCreateInventoryHistoryLogSheet_(readyHeaders) {
  var logSs = getCatalogSpreadsheet_();
  var sh = logSs.getSheetByName(INVENTORY_HISTORY_LOG_SHEET_NAME);
  if (!sh) sh = logSs.insertSheet(INVENTORY_HISTORY_LOG_SHEET_NAME);

  var wantedHeaders = readyHeaders.slice();
  wantedHeaders.push("Stock before shipment");
  wantedHeaders.push("Stock after shipment");
  wantedHeaders.push("Shipped at");

  sh.getRange(1, 1, 1, wantedHeaders.length).setValues([wantedHeaders]);
  return sh;
}

function shipOrdersByRange_(allSheet, startRow, endRow) {
  var readySs = SpreadsheetApp.openById(SHIPPING_TARGET_SPREADSHEET_ID);
  var readySheet = readySs.getSheetByName("stockcheckRTS");
  if (!readySheet) { showStyledAlert_("Error", "ReadyToShip sheet not found in target spreadsheet.", "error"); return; }

  var invSheet = getInventorySheet_();
  if (!invSheet) { showStyledAlert_("Error", "Inventory sheet not found in catalog spreadsheet.", "error"); return; }

  var allCols = allSheet.getLastColumn();
  var numRows = endRow - startRow + 1;
  var allHeaders = allSheet.getRange(HEADER_ROW, 1, 1, allCols).getValues()[0];
  var selectedRows = allSheet.getRange(startRow, 1, numRows, allCols).getValues();

  var readyHeaderRow = detectHeaderRowByAliases_(readySheet, ["PortalOrderID", "PORTAL_ORDER_ID"], 5);
  var readyCols = readySheet.getLastColumn();
  var readyHeaders = readySheet.getRange(readyHeaderRow, 1, 1, readyCols).getValues()[0];

  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var workingStockByVendor = {};
  for (var v in invMap) {
    if (Object.prototype.hasOwnProperty.call(invMap, v)) workingStockByVendor[v] = Number(invMap[v].Stock) || 0;
  }

  var rowsToShip = [], logRows = [], stockUpdates = {}, statusOut = [];
  var shippedCount = 0, skippedCount = 0, skippedStatusCount = 0, skippedStockCount = 0;

  for (var i = 0; i < selectedRows.length; i++) {
    var row = selectedRows[i];
    statusOut.push([row[HEADERS.ORDER_STATUS - 1]]);

    if (!isPendingOrderStatus_(row[HEADERS.ORDER_STATUS - 1])) { skippedCount++; skippedStatusCount++; continue; }

    var sku = String(row[HEADERS.SKU - 1] || "").trim();
    var qty = Number(row[HEADERS.QTY - 1]) || 0;
    if (!sku || qty <= 0) { skippedCount++; skippedStockCount++; continue; }

    var vendorSKU = portalMap[sku];
    if (!vendorSKU || !invMap[vendorSKU]) { skippedCount++; skippedStockCount++; continue; }

    var beforeStock = Number(workingStockByVendor[vendorSKU]) || 0;
    if (!(beforeStock >= qty && beforeStock > 0)) { skippedCount++; skippedStockCount++; continue; }

    var afterStock = beforeStock - qty;
    workingStockByVendor[vendorSKU] = afterStock;
    stockUpdates[vendorSKU] = afterStock;

    var shipRow = row.slice();
    shipRow[HEADERS.ORDER_STATUS - 1] = "Ready to Ship";
    var readyRow = buildReadyToShipRowFromAllOrders_(allHeaders, shipRow, readyHeaders);
    rowsToShip.push(readyRow);
    logRows.push(readyRow.concat([beforeStock, afterStock, new Date()]));
    statusOut[i] = ["Ready to Ship"];
    shippedCount++;
  }

  if (rowsToShip.length === 0) {
    refreshAllOrdersStockAvailability_();
    showStyledAlert_("Nothing Shipped", "No selected rows are currently available in stock.", "warning");
    return;
  }

  var readyAppendStart = Math.max(readySheet.getLastRow() + 1, readyHeaderRow + 1);
  readySheet.getRange(readyAppendStart, 1, rowsToShip.length, readyHeaders.length).setValues(rowsToShip);

  for (var vs in stockUpdates) {
    if (!Object.prototype.hasOwnProperty.call(stockUpdates, vs)) continue;
    var inv = invMap[vs];
    if (!inv) continue;
    invSheet.getRange(inv.invRow, INVENTORY_COLS.Stock).setValue(stockUpdates[vs]);
  }

  var logSheet = getOrCreateInventoryHistoryLogSheet_(readyHeaders);
  var logStart = Math.max(logSheet.getLastRow() + 1, 2);
  logSheet.getRange(logStart, 1, logRows.length, readyHeaders.length + 3).setValues(logRows);

  allSheet.getRange(startRow, HEADERS.ORDER_STATUS, numRows, 1).setValues(statusOut);
  refreshAllNewOrdersFromCatalog_();
  refreshAllOrdersStockAvailability_();

  showStyledAlert_("Ship Orders Complete", "Shipped rows: " + shippedCount + "<br>Skipped rows: " + skippedCount + "<br>Inventory updated and history logged.", "success");
}

function shipOrders() { showStyledAlert_("Legacy Action Disabled", "Use the Master Workflow menu.", "info"); }
function shipAllOrders() { showStyledAlert_("Legacy Action Disabled", "Use the Master Workflow menu.", "info"); }

function normalizeDropdownValuesInRange_(sheet, startRow, endRow, startCol, endCol) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  for (var headerKey in VALIDATION_MAP) {
    var colIdx = HEADERS[headerKey];
    if (!isColumnInRange_(colIdx, startCol, endCol)) continue;

    var options = getConfigValues_(VALIDATION_MAP[headerKey]);
    var canonical = {};
    for (var i = 0; i < options.length; i++) {
      var opt = String(options[i]).trim();
      if (opt !== "") canonical[opt.toLowerCase()] = opt;
    }

    var range = sheet.getRange(startRow, colIdx, numRows, 1);
    var vals = range.getValues();
    var changed = false;

    for (var r = 0; r < vals.length; r++) {
      var raw = vals[r][0];
      var trimmed = String(raw).trim();
      if (trimmed === "") {
        if (raw !== "") { vals[r][0] = ""; changed = true; }
        continue;
      }
      var match = canonical[trimmed.toLowerCase()];
      if (match && raw !== match) { vals[r][0] = match; changed = true; }
    }
    if (changed) range.setValues(vals);
  }
}

function getRowsNeedingSkuRefresh_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return [];
  var skuVals = sheet.getRange(startRow, HEADERS.SKU, numRows, 1).getValues();
  var errVals = sheet.getRange(startRow, HEADERS.what_to_fix, numRows, 1).getValues();
  var rows = [];

  for (var i = 0; i < numRows; i++) {
    var sku = String(skuVals[i][0]).trim();
    var err = String(errVals[i][0]).toLowerCase();
    if (!sku) continue;
    if (err.indexOf("sku not found") !== -1 || err.indexOf("vendor sku not found") !== -1 || err.indexOf("vendor details not found in inventory") !== -1) {
      rows.push(startRow + i);
    }
  }
  return rows;
}

function isColumnInRange_(col, startCol, endCol) { return col >= startCol && col <= endCol; }

// ─────────────────────────────────────────────────────────────────────────────
//  10. REAL-TIME ERROR LOGGING
// ─────────────────────────────────────────────────────────────────────────────

function validateAndLogErrors_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var validOptions = {};
  for (var key in VALIDATION_MAP) {
    var rawValues = getConfigValues_(VALIDATION_MAP[key]);
    var validSet = {};
    for (var v = 0; v < rawValues.length; v++) validSet[String(rawValues[v]).trim().toLowerCase()] = true;
    validOptions[key] = validSet;
  }

  var allData = sheet.getRange(startRow, 1, numRows, TOTAL_COLUMNS).getValues();
  var errOutput = [];
  var bgByMandatoryKey = {};
  for (var mInit = 0; mInit < MANDATORY_FIELDS.length; mInit++) bgByMandatoryKey[MANDATORY_FIELDS[mInit].key] = [];

  for (var i = 0; i < numRows; i++) {
    var rowData = allData[i];
    var hasAnyData = false;
    for (var c = 0; c < 30; c++) {
      if (String(rowData[c]).trim() !== "") { hasAnyData = true; break; }
    }

    if (!hasAnyData) {
      errOutput.push([""]);
      for (var m = 0; m < MANDATORY_FIELDS.length; m++) bgByMandatoryKey[MANDATORY_FIELDS[m].key].push([CLEAR_COLOR]);
      continue;
    }

    var missing = [];
    for (var j = 0; j < MANDATORY_FIELDS.length; j++) {
      var field = MANDATORY_FIELDS[j];
      var colIdx = HEADERS[field.key];
      var val = String(rowData[colIdx - 1]).trim();

      if (val === "" || val === "0" || val === "undefined" || val === "null") {
        bgByMandatoryKey[field.key].push([ERROR_RED]);
        missing.push(field.label);
      } else {
        var limits = validOptions[field.key];
        if (limits && !limits[val.toLowerCase()]) {
          bgByMandatoryKey[field.key].push([ERROR_RED]);
          missing.push("Invalid " + field.label + " (Not in Dropdown)");
        } else {
          bgByMandatoryKey[field.key].push([CLEAR_COLOR]);
        }
      }
    }

    var currentError = String(rowData[HEADERS.what_to_fix - 1]).trim();
    var customErrors = currentError.split("|").map(function (s) { return s.trim(); }).filter(function (s) {
      return s.length > 0 && s.indexOf("⚠ Missing:") === -1 && s !== "✅ All fields complete" && s !== "Duplicate" && s !== "Stock is 0";
    });

    if (missing.length > 0) {
      var combinedOutput = "⚠ Missing: " + missing.join(", ");
      if (customErrors.length > 0) combinedOutput += " | " + customErrors.join(" | ");
      errOutput.push([combinedOutput]);
    } else {
      if (customErrors.length > 0) errOutput.push([customErrors.join(" | ")]);
      else errOutput.push(["✅ All fields complete"]);
    }
  }

  sheet.getRange(startRow, HEADERS.what_to_fix, numRows, 1).setValues(errOutput);

  for (var k = 0; k < MANDATORY_FIELDS.length; k++) {
    var key = MANDATORY_FIELDS[k].key;
    var col = HEADERS[key];
    var colBg = bgByMandatoryKey[key];
    try { sheet.getRange(startRow, col, numRows, 1).setBackgrounds(colBg); } 
    catch (bgErr) {
      for (var r = 0; r < numRows; r++) sheet.getRange(startRow + r, col).setBackground(colBg[r][0]);
    }
  }
}

function validateAllRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) { showStyledAlert_("No Data", "No data rows found to validate.", "info"); return; }
  validateAndLogErrors_(sheet, DATA_START_ROW, lastRow);
  showStyledAlert_("Validation Complete", "Error logs and highlights updated for " + (lastRow - 1) + " row(s).<br><br>🔴 Red cells = missing mandatory fields<br>✅ Green text in Error_Logs = all complete", "info");
}

function handleRowDeletion_(sheet) { regenerateInternalOrderNos_(sheet, null); }

function onChange(e) {
  if (!e) return;
  if (e.changeType === "REMOVE_ROW" || e.changeType === "INSERT_ROW") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
    if (!sheet) return;

    if (e.changeType === "REMOVE_ROW") regenerateInternalOrderNos_(sheet, null);

    var lastRow = sheet.getLastRow();
    if (lastRow >= DATA_START_ROW) applyValidationsToRange_(sheet, DATA_START_ROW, lastRow);
    PropertiesService.getScriptProperties().setProperty(ROW_COUNT_KEY, String(lastRow));

    try { applyWorkflowValidations_(); } catch (e) {}
  }
}

// [SECURITY UPDATE] Escape dynamically injected HTML strings here 
function showStyledAlert_(title, message, type) {
  var colors = {
    success: { bg: "#e8f5e9", accent: "#2e7d32", icon: "✅", gradient: "linear-gradient(135deg, #43a047, #66bb6a)" },
    error: { bg: "#ffebee", accent: "#c62828", icon: "❌", gradient: "linear-gradient(135deg, #e53935, #ef5350)" },
    warning: { bg: "#fff8e1", accent: "#f57f17", icon: "⚠️", gradient: "linear-gradient(135deg, #ff8f00, #ffb300)" },
    info: { bg: "#e3f2fd", accent: "#1565c0", icon: "ℹ️", gradient: "linear-gradient(135deg, #1e88e5, #42a5f5)" },
  };

  var c = colors[type] || colors.info;
  var safeMsg = escapeHtml_(message).replace(/\n/g, "<br>");

  var html = "<!DOCTYPE html><html><head>" +
    '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">' +
    "<style>* { margin: 0; padding: 0; box-sizing: border-box; } body { font-family: \"Inter\", sans-serif; background: transparent; display: flex; justify-content: center; padding: 4px; } .card { width: 100%; max-width: 420px; background: #fff; border-radius: 16px; box-shadow: 0 8px 32px rgba(0,0,0,0.12), 0 2px 8px rgba(0,0,0,0.06); overflow: hidden; animation: slideUp 0.35s cubic-bezier(0.22, 1, 0.36, 1); } .header { background: " + c.gradient + "; padding: 20px 24px; color: #fff; } .header .icon { font-size: 28px; margin-bottom: 6px; } .header h2 { font-size: 18px; font-weight: 700; letter-spacing: -0.3px; } .body { padding: 20px 24px; color: #37474f; font-size: 14px; line-height: 1.6; background: " + c.bg + "; border-top: 3px solid " + c.accent + "; } .btn-row { padding: 16px 24px; text-align: right; background: #fafafa; } .btn { background: " + c.gradient + "; color: #fff; border: none; padding: 10px 28px; border-radius: 8px; font-family: \"Inter\", sans-serif; font-size: 14px; font-weight: 600; cursor: pointer; transition: transform 0.15s, box-shadow 0.15s; box-shadow: 0 4px 12px rgba(0,0,0,0.15); } .btn:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(0,0,0,0.2); } @keyframes slideUp { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }</style></head><body>" +
    '<div class="card"><div class="header"><div class="icon">' + c.icon + "</div><h2>" + escapeHtml_(title) + "</h2></div><div class="body">" + safeMsg + "</div><div class=\"btn-row\"><button class=\"btn\" onclick=\"google.script.host.close()\">OK</button></div></div></body></html>";

  var output = HtmlService.createHtmlOutput(html).setWidth(460).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(output, " ");
}

function showStyledConfirm_(title, message) {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(title, message, ui.ButtonSet.YES_NO);
  return response === ui.Button.YES;
}

// ─────────────────────────────────────────────────────────────────────────────
//  12. MENU & INITIALIZATION
// ─────────────────────────────────────────────────────────────────────────────

function syncLastMaxFromAllOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  var allOrdersSheet = ss.getSheetByName(SHEET_NAMES.ALL_ORDERS);
  var newOrdersSheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);

  if (!configSheet || !allOrdersSheet || !newOrdersSheet) { showStyledAlert_("Error", "Required sheets not found!", "error"); return; }

  var configLastRow = configSheet.getLastRow();
  if (configLastRow < 2) return;
  var configData = configSheet.getRange(2, 1, configLastRow - 1, configSheet.getLastColumn()).getValues();

  var prefixMaxMap = {};
  for (var i = 0; i < configData.length; i++) {
    var initial = String(configData[i][CONFIG_COLS.Initial_code - 1]).trim().toUpperCase();
    if (initial) prefixMaxMap[initial] = 0;
  }

  function scanSheet(sheet) {
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < DATA_START_ROW) return; 
    var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    var internalColIdx = 0;

    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim().toUpperCase();
      if (h === "INTERNAL ORDERNO" || h === "INTERNAL_ORDERNO") { internalColIdx = i; break; }
    }

    if (internalColIdx < 0) return;
    var data = sheet.getRange(DATA_START_ROW, internalColIdx + 1, lastRow - DATA_START_ROW + 1, 1).getValues();
    for (var r = 0; r < data.length; r++) {
      var val = String(data[r][0]).trim();
      var match = val.match(/^([A-Za-z]*)(\d+)(?:-[A-Za-z0-9]+)?$/);
      if (match) {
        var prefix = match[1].toUpperCase();
        var num = parseInt(match[2], 10);
        if (prefixMaxMap.hasOwnProperty(prefix) && num > prefixMaxMap[prefix]) prefixMaxMap[prefix] = num;
      }
    }
  }

  scanSheet(allOrdersSheet);
  scanSheet(newOrdersSheet);

  var lastMaxOutput = [];
  for (var j = 0; j < configData.length; j++) {
    var initial = String(configData[j][CONFIG_COLS.Initial_code - 1]).trim().toUpperCase();
    if (initial && prefixMaxMap.hasOwnProperty(initial)) lastMaxOutput.push([prefixMaxMap[initial]]);
    else lastMaxOutput.push([""]);
  }

  configSheet.getRange(2, CONFIG_COLS.Last_Maximum, lastMaxOutput.length, 1).setValues(lastMaxOutput);
  showStyledAlert_("Success", "Last Maximum successfully synced from AllOrders and New_Orders!", "success");
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu("⚙ Master Workflow")
    .addItem("Setup System", "setupSystem")
    .addItem("Create Picklist", "generatePickList")
    .addItem("🎫 Create Labels", "openLabelGeneratorUI")
    .addSeparator()
    .addItem("Sync Pending to stockcheckRTS", "syncPendingOrdersToStockcheckRTS")
    .addToUi();

  ui.createMenu("📦 Order Manager")
    .addItem("🔄 Apply All Validations", "applyAllValidations")
    .addItem("🔢 Regenerate All Order IDs", "regenerateAllInternalOrderNos")
    .addItem("📊 Recalculate All Financials", "recalcAllFinancials")
    .addItem("🔍 Refresh Product Costs (All)", "refreshAllProductCosts")
    .addItem("🚨 Validate All Rows", "validateAllRows")
    .addSeparator()
    .addItem("📋 Push Selected to AllOrders", "appendSelectedToAllOrders")
    .addItem("🚀 Push All Orders", "appendAllToAllOrders")
    .addItem("Install New Orders edit trigger", "installNewOrdersEditTrigger")
    .addItem("🔄 Sync Last Maximum", "syncLastMaxFromAllOrders")
    .addSeparator()
    .addItem("⚙️ Initialize System", "initializeSystem")
    .addToUi();

  if (typeof installInventoryEditTrigger === "function") {
    ui.createMenu("Inventory Sync")
      .addItem("Install edit sync trigger", "installInventoryEditTrigger")
      .addItem("Refresh All Orders stock", "refreshAllOrdersStockAvailability")
      .addItem("Refresh New Orders SKUs", "refreshNewOrdersProductInfo")
      .addItem("Refresh Inventory images", "refreshAllInventoryImages")
      .addToUi();
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (sheet) PropertiesService.getScriptProperties().setProperty(ROW_COUNT_KEY, String(sheet.getLastRow()));

  try {
    refreshAllOrdersStockAvailability_();
    syncPendingOrdersToStockcheckRTS_(true);
  } catch (e) {}
}

function initializeSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (!sheet) { showStyledAlert_("Error", "New_Orders sheet not found!", "error"); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow >= DATA_START_ROW) applyValidationsToRange_(sheet, DATA_START_ROW, lastRow);

  PropertiesService.getScriptProperties().setProperty(ROW_COUNT_KEY, String(lastRow));
  showStyledAlert_("System Initialized", "Order Management System is ready!<br><br>• Validations applied to " + Math.max(0, lastRow - 1) + " rows<br>• Row count tracking enabled<br>• All triggers active", "success");
}

function recalcAllFinancials() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  batchCalculateFinancials_(sheet, DATA_START_ROW, lastRow);
  showStyledAlert_("Financials Updated", "All financial calculations have been recalculated for " + (lastRow - 1) + " rows.", "success");
}

function refreshAllProductCosts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  batchUpdateProductInfo_(sheet, DATA_START_ROW, lastRow);
  showStyledAlert_("Product Costs Refreshed", "Product costs have been refreshed from MatchingTable + Inventory for " + (lastRow - 1) + " rows.", "info");
}
