/* =============================================================================
 *  ORDER MANAGEMENT SYSTEM — new_order.gs
 *  Google Apps Script for New_Orders sheet automation
 *
 *  Modules:
 *    1. Constants & Column Mappings
 *    2. Config Helpers
 *    3. Validation Logic (6 fields)
 *    4. Product Cost Lookup (PortalSKU → VendorSKU → Inventory)
 *    5. Stock Logic
 *    6. Internal Order ID (continuous, gap-fill on delete)
 *    7. Financial Calculations
 *    8. Append to AllOrders
 *    9. onEdit Trigger (bulk-aware)
 *   10. Row Deletion Handler
 *   11. UI / Styled Popups
 *   12. Menu & Initialization
 * ============================================================================= */

// ─────────────────────────────────────────────────────────────────────────────
//  1. CONSTANTS & COLUMN MAPPINGS
// ─────────────────────────────────────────────────────────────────────────────

var SHEET_NAMES = {
  NEW_ORDERS: "[REDACTED_NEW_ORDERS_SHEET]",
  ALL_ORDERS: "[REDACTED_ALL_ORDERS_SHEET]",
  CONFIG: "[REDACTED_CONFIG_SHEET]",
  STOCKCHECK_RTS: "[REDACTED_STOCKCHECK_SHEET]",
  CREATE_LABEL: "[REDACTED_CREATELABELS_SHEET]",
  PICKLIST: "[REDACTED_PICKLIST_SHEET]",
  DISPATCHED: "[REDACTED_DISPATCHED_SHEET]",
  DISPATCH_HISTORY: "[REDACTED_DISPATCH_HISTORY_SHEET]",
  MATCHING_TABLE: "[REDACTED_MATCHING_TABLE_SHEET]",
  INVENTORY: "[REDACTED_INVENTORY_SHEET]",
};

var SHIPPING_TARGET_SPREADSHEET_ID =
  "[REDACTED_SHIPPING_SPREADSHEET_ID]";
var SHIPPING_TARGET_SHEET_NAME = "[REDACTED_STOCKCHECK_SHEET]";
var CATALOG_SPREADSHEET_ID = "[REDACTED_CATALOG_SPREADSHEET_ID]";
var INVENTORY_HISTORY_LOG_SHEET_NAME = "[REDACTED_LOG_SHEET]";
var NEW_ORDERS_EDIT_TRIGGER_HANDLER = "handleNewOrdersAuthorizedEdit";
var NEW_ORDERS_EDIT_TRIGGER_INSTALLED_KEY =
  "[REDACTED_EDIT_TRIGGER_KEY]";
var WORKFLOW_CHANGE_TRIGGER_HANDLER = "onChange";
var WORKFLOW_CHANGE_TRIGGER_INSTALLED_KEY = "[REDACTED_WORKFLOW_TRIGGER_KEY]";

/**
 * New_Orders column headers → 1-based column indices.
 * Keep this mapping in sync with the actual sheet header row.
 */
var HEADERS = {
  PORTAL_ORDER_ID: 1,
  DELIVERY_PICKUP_DATE: 2,
  ORDER_STATUS: 3,
  PAYMENT_STATUS: 4,
  STAFF_NOTES: 5,
  BRAND_NAME: 6,
  INTERNAL_ORDERNO: 7,
  PURCHASE_DATE: 8,
  SALES_CHANNEL: 9,
  SKU: 10,
  ITEM_NAME: 11,
  CATEGORY: 12,
  IMAGE: 13,
  ORDERNOTE: 14,
  QTY: 15,
  CURRENCY: 16,
  CURRENCY_PRICE: 17,
  FULLNAME: 18,
  ADDRESSLINE1: 19,
  ADDRESSLINE2: 20,
  CITY: 21,
  STATE: 22,
  PINCODE: 23,
  COUNTRY: 24,
  PHONE: 25,
  COURIER_NAME: 26,
  TRACKING_CODE: 27,
  STATUS: 28,
  SHIPPING_CHARGE: 29,
  IMAGE_URL: 30,
  LISTING_URL: 31,
  NEW_TRACKING: 32,
  Conversion_Rate: 33,
  Price_in_INR: 34,
  Shipping_Charge_Product: 35,
  Product_Cost: 36,
  Maximum_Expense: 37,
  Actual_Expense: 38,
  Maximum_Profit: 39,
  Actual_Profit: 40,
  STOCK: 41,
  what_to_fix: 42,
};

/** Total number of columns in New_Orders */
var TOTAL_COLUMNS = 42;

/** Header row in New_Orders (data starts at row 3) */
var HEADER_ROW = 2;
var DATA_START_ROW = 3;
var WORKFLOW_HEADER_ROW = 1;
var WORKFLOW_DATA_START_ROW = 2;

var STOCKCHECK_COLS = {
  PORTAL_ORDER_ID: 1,
  INTERNAL_ORDERNO: 2,
  STAFF_NOTES: 3,
  QTY: 4,
  IMAGE: 5,
  IMAGE_URL: 6,
  SKU: 7,
  SALES_CHANNEL: 8,
  COUNTRY: 9,
  ActualStockStatus: 10,
  DoCreateLabels: 11,
};

var CREATE_LABEL_COLS = {
  INTERNAL_ORDERNO: 1,
  SKU: 2,
  IMAGE: 3,
  IMAGE_URL: 4,
  FULLNAME: 5,
  ADDRESSLINE1: 6,
  ADDRESSLINE2: 7,
  CITY: 8,
  STATE: 9,
  PINCODE: 10,
  COUNTRY: 11,
  PHONE: 12,
  LabelCreated: 13,
  COURIER_NAME: 14,
  DELIVERY_PICKUP_DATE: 15,
  SALES_CHANNEL: 16,
  TRACKING_CODE: 17,
};

var DISPATCHED_COLS = {
  INTERNAL_ORDERNO: 1,
  SKU: 2,
  IMAGE: 3,
  IMAGE_URL: 4,
  FULLNAME: 5,
  ADDRESSLINE1: 6,
  ADDRESSLINE2: 7,
  CITY: 8,
  STATE: 9,
  PINCODE: 10,
  COUNTRY: 11,
  PHONE: 12,
  DISPACHED_STATUS: 13,
  PICKUP_DATE: 14,
  COURIER_NAME: 15,
  TRACKING_CODE: 16,
  SALES_CHANNEL: 17,
  PACKET_COUNT: 18,
  AMAZON_TRACKING_GEN: 19,
};

var DISPATCH_HISTORY_COLS = {
  DATE: 1,
  COURIER_NAME: 2,
  PACKET_COUNT: 3,
  INTERNAL_ORDERNO: 4,
  DISPACHED_STATUS: 5,
  IMAGE: 6,
  IMAGE_URL: 7,
  TRACKING_CODE: 8,
  DONE_BY: 9,
};

/**
 * Config sheet column headers → 1-based indices
 */
var CONFIG_COLS = {
  Country: 1,
  Charge: 2,
  Magic_Number: 3,
  Currency_Code: 4,
  Brand_Name: 5,
  Initial_code: 6,
  Last_Maximum: 7,
  Unique_Sales_Channel: 8,
  Category: 9,
  Sales_Channel: 10,
  Payment_Status: 11,
};

/**
 * MatchingTable columns
 */
var MATCHING_COLS = {
  VendorSKU: 1,
  PortalSKU: 2,
};

/**
 * Inventory columns
 */
var INVENTORY_COLS = {
  VendorSKU: 1,
  Image: 2,
  Item_Name: 3,
  Stock: 4,
  Product_Cost: 5,
  Location: 6,
  IMAGE_URL: 7,
};

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

/**
 * Validation mapping:  New_Orders header key → Config column key
 * These fields will get dropdown validation from Config values.
 */
var VALIDATION_MAP = {
  CATEGORY: "Category",
  COUNTRY: "Country",
  CURRENCY: "Currency_Code",
  BRAND_NAME: "Brand_Name",
  SALES_CHANNEL: "Sales_Channel",
  PAYMENT_STATUS: "Payment_Status",
};

/**
 * Mandatory fields that MUST be filled for a valid order.
 * If any of these are empty when the row has a PORTAL_ORDER_ID,
 * the cell is highlighted red and `what_to_fix` is populated.
 */
var MANDATORY_FIELDS = [
  { key: "PORTAL_ORDER_ID", label: "Portal Order ID" },
  { key: "BRAND_NAME", label: "Brand Name" },
  { key: "SKU", label: "SKU" },
  { key: "QTY", label: "Quantity" },
  { key: "CURRENCY", label: "Currency" },
  { key: "CURRENCY_PRICE", label: "Currency Price" },
  { key: "COUNTRY", label: "Country" },
  { key: "FULLNAME", label: "Full Name" },
  { key: "SALES_CHANNEL", label: "Sales Channel" },
  { key: "CATEGORY", label: "Category" },
  { key: "DELIVERY_PICKUP_DATE", label: "Delivery/Pickup Date" },
  { key: "PAYMENT_STATUS", label: "Payment Status" },
  { key: "ITEM_NAME", label: "Item Name" },
  { key: "ORDERNOTE", label: "Order Note" },
  { key: "ADDRESSLINE1", label: "Address Line1" },
  { key: "CITY", label: "City" },
  { key: "STATE", label: "State" },
  { key: "PINCODE", label: "Pincode" },
  { key: "IMAGE_URL", label: "Image url" },
  { key: "PURCHASE_DATE", label: "Purchase date" },
];

/** Colors for error highlighting */
var ERROR_RED = "#ffcdd2"; // light red for empty mandatory cells
var CLEAR_COLOR = "#ffffff"; // explicit white so pasted formatting is always overridden

// ─────────────────────────────────────────────────────────────────────────────
//  2. CONFIG HELPERS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns the Config sheet.
 */
function getConfigSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET_NAMES.CONFIG,
  );
}

/**
 * Reads a single column from Config and returns unique non-empty values.
 * Uses dynamic header lookup to allow columns to move.
 * @param {string|string[]} configColAliases – aliases for the config column header
 * @return {string[]}
 */
function getConfigValues_(configColAliases) {
  var sheet = getConfigSheet_();
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colIdx = -1;
  var aliases = Array.isArray(configColAliases)
    ? configColAliases
    : [configColAliases];

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
    // Fallback to hardcoded map if header not found but exists in CONFIG_COLS
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

/**
 * Returns the column range (excluding header) from Config for dynamic dropdowns.
 */
function getConfigColRange_(configColAliases) {
  var sheet = getConfigSheet_();
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return null;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colIdx = -1;
  var aliases = Array.isArray(configColAliases)
    ? configColAliases
    : [configColAliases];

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

/**
 * Returns the full Config data as an array of row objects (keyed by CONFIG_COLS).
 * Caches within a single script run via CacheService or simple closure.
 */
var _configCache = null;
function getConfigData_() {
  if (_configCache) return _configCache;
  var sheet = getConfigSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();
  _configCache = data;
  return data;
}

/**
 * Builds a lookup: Brand_Name → { Initial_code, Last_Maximum, Charge, Magic_Number }
 */
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
        configRow: i + 2, // 1-based row in Config sheet
      };
    }
  }
  return map;
}

/**
 * Builds a lookup: Country → { Charge, Magic_Number }
 */
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
      map[normalizeCountryKey_(country)] = info; // case/spacing-insensitive alias
    }
  }
  return map;
}

function normalizeCountryKey_(country) {
  return String(country || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

// ─────────────────────────────────────────────────────────────────────────────
//  3. VALIDATION LOGIC
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Applies dropdown validations for all 6 validated fields to a row range.
 * @param {Sheet} sheet – the New_Orders sheet
 * @param {number} startRow – first data row to validate
 * @param {number} endRow   – last data row to validate
 */
function applyValidationsToRange_(sheet, startRow, endRow) {
  if (startRow > endRow) return;
  var numRows = endRow - startRow + 1;

  var configSheet = getConfigSheet_();
  if (!configSheet) return;

  var maxConfigRows = Math.max(2, configSheet.getMaxRows());

  // Apply each validation
  for (var headerKey in VALIDATION_MAP) {
    var colIdx = HEADERS[headerKey];
    if (!colIdx) continue;

    var configColKey = VALIDATION_MAP[headerKey];
    var configColIdx = CONFIG_COLS[configColKey];
    if (!configColIdx) continue;

    // Build the dynamic range from Config (Row 2 to max)
    var listRange = configSheet.getRange(2, configColIdx, maxConfigRows - 1, 1);

    var targetRange = sheet.getRange(startRow, colIdx, numRows, 1);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(listRange, true) // true = show dropdown
      .setAllowInvalid(false)
      .build();
    targetRange.setDataValidation(rule);
  }
}

/**
 * Applies validations to all data rows in New_Orders.
 * Called on initialization or after structural changes.
 */
function applyAllValidations() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET_NAMES.NEW_ORDERS,
  );
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;
  applyValidationsToRange_(sheet, DATA_START_ROW, lastRow);
  showStyledAlert_(
    "Validations Applied",
    "Dropdown validations have been applied to all data rows.",
    "success",
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  4. PRODUCT COST LOOKUP  (SKU → PortalSKU → VendorSKU → Inventory)
// ─────────────────────────────────────────────────────────────────────────────
/**
 * Updates IMAGE column using IMAGE_URL column.
 * Converts URL into =IMAGE() formula.
 * @param {Sheet} sheet
 * @param {number} startRow
 * @param {number} endRow
 */
function updateImageFormulas_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  // Read IMAGE_URL values
  var urls = sheet
    .getRange(startRow, HEADERS.IMAGE_URL, numRows, 1)
    .getValues();

  var formulas = [];

  for (var i = 0; i < numRows; i++) {
    var url = String(urls[i][0]).trim();

    if (!url) {
      formulas.push([""]);
      continue;
    }

    // Build IMAGE formula
    formulas.push(['=IMAGE("' + url.replace(/"/g, '""') + '")']);
  }

  // Write formulas to IMAGE column
  sheet.getRange(startRow, HEADERS.IMAGE, numRows, 1).setFormulas(formulas);
}

/**
 * Updates Inventory IMAGE column using Inventory IMAGE_URL values.
 * Keeps Inventory logic separate from New_Orders image logic.
 * @param {Sheet} sheet
 * @param {number} startRow
 * @param {number} endRow
 */
function updateInventoryImageFormulas_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var urls = sheet
    .getRange(startRow, INVENTORY_COLS.IMAGE_URL, numRows, 1)
    .getValues();
  var formulas = [];
  for (var i = 0; i < numRows; i++) {
    var url = String(urls[i][0]).trim();
    if (!url) {
      formulas.push([""]);
    } else {
      formulas.push(['=IMAGE("' + url.replace(/"/g, '""') + '")']);
    }
  }

  sheet
    .getRange(startRow, INVENTORY_COLS.Image, numRows, 1)
    .setFormulas(formulas);
}

/**
 * Updates CreateLabels IMAGE column using CreateLabels IMAGE_URL values.
 * Converts IMAGE_URL into =IMAGE() formulas for immediate rendering.
 * @param {Sheet} sheet
 * @param {number} startRow
 * @param {number} endRow
 */
function updateCreateLabelsImageFormulas_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var urls = sheet
    .getRange(startRow, CREATE_LABEL_COLS.IMAGE_URL, numRows, 1)
    .getValues();
  var formulas = [];
  for (var i = 0; i < numRows; i++) {
    var url = String(urls[i][0]).trim();
    if (!url) {
      formulas.push([""]);
    } else {
      formulas.push(['=IMAGE("' + url.replace(/"/g, '""') + '")']);
    }
  }

  sheet
    .getRange(startRow, CREATE_LABEL_COLS.IMAGE, numRows, 1)
    .setFormulas(formulas);
}

/**
 * Updates Dispatched IMAGE column using Dispatched IMAGE_URL values.
 * Converts IMAGE_URL into =IMAGE() formulas for immediate rendering.
 * @param {Sheet} sheet
 * @param {number} startRow
 * @param {number} endRow
 */
function updateDispatchedImageFormulas_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var urls = sheet
    .getRange(startRow, DISPATCHED_COLS.IMAGE_URL, numRows, 1)
    .getValues();
  var formulas = [];
  for (var i = 0; i < numRows; i++) {
    var url = String(urls[i][0]).trim();
    if (!url) {
      formulas.push([""]);
    } else {
      formulas.push(['=IMAGE("' + url.replace(/"/g, '""') + '")']);
    }
  }

  sheet
    .getRange(startRow, DISPATCHED_COLS.IMAGE, numRows, 1)
    .setFormulas(formulas);
}

/**
 * Updates stockcheckRTS IMAGE column using IMAGE_URL values.
 * Converts IMAGE_URL into =IMAGE() formulas for immediate rendering.
 * @param {Sheet} sheet
 * @param {number} startRow
 * @param {number} endRow
 */
function updateStockcheckImageFormulas_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var urls = sheet
    .getRange(startRow, STOCKCHECK_COLS.IMAGE_URL, numRows, 1)
    .getValues();
  var formulas = [];
  for (var i = 0; i < numRows; i++) {
    var url = String(urls[i][0]).trim();
    if (!url) {
      formulas.push([""]);
    } else {
      formulas.push(['=IMAGE("' + url.replace(/"/g, '""') + '")']);
    }
  }

  sheet
    .getRange(startRow, STOCKCHECK_COLS.IMAGE, numRows, 1)
    .setFormulas(formulas);
}
/**
 * Builds PortalSKU → VendorSKU map from MatchingTable.
 * @return {Object.<string, string>}
 */
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

/**
 * Builds VendorSKU → { Product_Cost, Item_Name, Image, Stock, rowIndex } from Inventory.
 * @return {Object}
 */
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
        invRow: i + 2, // 1-based row in Inventory
      };
    }
  }
  return map;
}

/**
 * Looks up product info for a single SKU with explicit status.
 * @param {string} sku
 * @param {Object} portalMap
 * @param {Object} invMap
 * @return {Object}
 */
function lookupProductInfo_(sku, portalMap, invMap) {
  var vendorSKU = portalMap[sku];
  if (!vendorSKU) {
    return { ok: false, reason: "vendor_sku_missing", sku: sku };
  }
  var inv = invMap[vendorSKU];
  if (!inv) {
    return {
      ok: false,
      reason: "vendor_details_missing",
      sku: sku,
      vendorSKU: vendorSKU,
    };
  }
  return {
    ok: true,
    reason: "ok",
    Product_Cost: inv.Product_Cost,
    Item_Name: inv.Item_Name,
    Image: inv.Image,
    VendorSKU: vendorSKU,
    Stock: inv.Stock,
    invRow: inv.invRow,
  };
}

function isSkuLookupErrorMessage_(msg) {
  var m = String(msg || "").toLowerCase();
  return (
    m.indexOf("sku not found") !== -1 ||
    m.indexOf("vendor sku not found") !== -1 ||
    m.indexOf("vendor details not found in inventory") !== -1
  );
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
  return (
    String(status || "")
      .trim()
      .toLowerCase() === "pending"
  );
}

/**
 * Updates Product_Cost, ITEM_NAME, IMAGE for specified rows based on their SKU.
 * @param {Sheet} sheet – New_Orders sheet
 * @param {number[]} rows – array of 1-based row numbers
 * @param {boolean=} suppressAlerts – true to avoid UI popups (safe for triggers)
 */
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

    // Filter out previous custom errors to start fresh for this validation
    var errorsArr = curVal
      .split("|")
      .map(function (e) {
        return e.trim();
      })
      .filter(function (e) {
        return (
          e.length > 0 &&
          !isSkuLookupErrorMessage_(e) &&
          e !== "Stock is 0" &&
          e !== "✅ All fields complete"
        );
      });

    if (info.ok) {
      sheet.getRange(row, HEADERS.Product_Cost).setValue(info.Product_Cost);
      sheet.getRange(row, HEADERS.STOCK).setValue(info.Stock);
      // sheet.getRange(row, HEADERS.ITEM_NAME).setValue(info.Item_Name); // Explicitly preserved to prevent overwrite
      errCell.setValue(errorsArr.join(" | "));
    } else {
      sheet.getRange(row, HEADERS.STOCK).setValue("");
      if (info.reason === "vendor_sku_missing") {
        errorsArr.push("Vendor SKU not found: " + sku);
      } else if (info.reason === "vendor_details_missing") {
        errorsArr.push(
          "Vendor details not found in Inventory: " + info.vendorSKU,
        );
      } else {
        errorsArr.push("SKU not found: " + sku);
      }
      errCell.setValue(errorsArr.join(" | "));
      errors.push("Row " + row + ": " + errorsArr[errorsArr.length - 1]);
    }
  }

  if (errors.length > 0 && !suppressAlerts) {
    showStyledAlert_("SKU Lookup Warning", errors.join("\n"), "warning");
  }
}

/**
 * Batch-optimized version: reads all SKUs at once, writes results in batch.
 * @param {Sheet} sheet
 * @param {number} startRow
 * @param {number} endRow
 * @param {boolean=} suppressAlerts – true to avoid UI popups (safe for triggers)
 */
function batchUpdateProductInfo_(sheet, startRow, endRow, suppressAlerts) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();

  // Batch-read SKU column
  var skuValues = sheet.getRange(startRow, HEADERS.SKU, numRows, 1).getValues();
  var existingErrVals = sheet
    .getRange(startRow, HEADERS.what_to_fix, numRows, 1)
    .getValues();

  // Prepare output arrays
  var costOut = [];
  var stockOut = [];
  var errOut = [];
  var errors = [];

  for (var i = 0; i < numRows; i++) {
    var sku = String(skuValues[i][0]).trim();
    var curVal = String(existingErrVals[i][0]).trim();
    var errorsArr = curVal
      .split("|")
      .map(function (e) {
        return e.trim();
      })
      .filter(function (e) {
        return (
          e.length > 0 &&
          !isSkuLookupErrorMessage_(e) &&
          e !== "Stock is 0" &&
          e !== "✅ All fields complete"
        );
      });

    if (!sku) {
      costOut.push([""]);
      stockOut.push([""]);
      errOut.push([errorsArr.join(" | ")]);
      continue;
    }
    var info = lookupProductInfo_(sku, portalMap, invMap);
    if (info.ok) {
      costOut.push([info.Product_Cost]);
      stockOut.push([info.Stock]);
      errOut.push([errorsArr.join(" | ")]);
    } else {
      costOut.push([""]);
      stockOut.push([""]);
      if (info.reason === "vendor_sku_missing") {
        errorsArr.push("Vendor SKU not found: " + sku);
      } else if (info.reason === "vendor_details_missing") {
        errorsArr.push(
          "Vendor details not found in Inventory: " + info.vendorSKU,
        );
      } else {
        errorsArr.push("SKU not found: " + sku);
      }
      errOut.push([errorsArr.join(" | ")]);
      errors.push(
        "Row " + (startRow + i) + ": " + errorsArr[errorsArr.length - 1],
      );
    }
  }

  // Batch-write
  sheet.getRange(startRow, HEADERS.Product_Cost, numRows, 1).setValues(costOut);
  sheet.getRange(startRow, HEADERS.STOCK, numRows, 1).setValues(stockOut);
  // sheet.getRange(startRow, HEADERS.ITEM_NAME, numRows, 1).setValues(nameOut); // Explicitly preserved to prevent overwrite
  sheet.getRange(startRow, HEADERS.what_to_fix, numRows, 1).setValues(errOut);

  if (errors.length > 0 && !suppressAlerts) {
    showStyledAlert_(
      "SKU Lookup Issues",
      errors.length + " SKU(s) not found:\n" + errors.join("\n"),
      "warning",
    );
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  5. STOCK LOGIC
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Decrements stock in Inventory for given rows.
 * Uses: SKU → PortalSKU → VendorSKU → Inventory.Stock
 * @param {Sheet} sheet – New_Orders
 * @param {number[]} rows
 */
function decrementStockForRows_(sheet, rows) {
  if (!rows || rows.length === 0) return;

  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var invSheet = getInventorySheet_();
  if (!invSheet) return;

  // Aggregate qty changes per VendorSKU
  var changes = {}; // vendorSKU → totalQty

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var sku = String(sheet.getRange(row, HEADERS.SKU).getValue()).trim();
    var qty = Number(sheet.getRange(row, HEADERS.QTY).getValue()) || 0;
    if (!sku || qty <= 0) continue;

    var vendorSKU = portalMap[sku];
    if (!vendorSKU || !invMap[vendorSKU]) continue;

    changes[vendorSKU] = (changes[vendorSKU] || 0) + qty;
  }

  // Apply stock decrements
  for (var vsku in changes) {
    var inv = invMap[vsku];
    var newStock = Math.max(0, inv.Stock - changes[vsku]);
    invSheet.getRange(inv.invRow, INVENTORY_COLS.Stock).setValue(newStock);
  }
}

/**
 * Restores stock in Inventory (used when orders are cancelled/deleted).
 * @param {string} sku – the portal SKU
 * @param {number} qty – quantity to restore
 */
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

/**
 * Helper: Finds the highest sequence number historically used for a brand prefix
 * by scanning the AllOrders sheet.
 */
function getMaxFromAllOrders_(prefix) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.ALL_ORDERS);
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return 0;

  var headers = sheet
    .getRange(HEADER_ROW, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  var internalColIdx = -1;
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim().toUpperCase();
    if (h === "INTERNAL ORDERNO" || h === "INTERNAL_ORDERNO") {
      internalColIdx = i;
      break;
    }
  }
  if (internalColIdx < 0) return 0;

  var data = sheet
    .getRange(
      DATA_START_ROW,
      internalColIdx + 1,
      lastRow - DATA_START_ROW + 1,
      1,
    )
    .getValues();

  // Robust handling: prefixes in Config may include digits (e.g. "SY0").
  // Instead of relying on a strict regex that only captures letters,
  // check whether the cell value starts with the configured prefix
  // (case-insensitive) and then parse the trailing numeric portion.
  var max = 0;
  var pfxUpper = String(prefix || "").toUpperCase();
  for (var j = 0; j < data.length; j++) {
    var val = String(data[j][0] || "").trim();
    if (!val) continue;
    if (pfxUpper && val.toUpperCase().indexOf(pfxUpper) !== 0) continue;

    // Remove the prefix and any leading hyphen-separated suffix, then extract digits
    var rest = val.substring(pfxUpper.length);
    // If rest begins with a hyphen (e.g. "SY0-ABC123"), strip it
    if (rest.indexOf("-") === 0) rest = rest.substring(1);
    var digitsMatch = rest.match(/(\d+)/);
    if (digitsMatch) {
      var num = parseInt(digitsMatch[1], 10);
      if (!isNaN(num)) max = Math.max(max, num);
    }
  }
  return max;
}

// ─────────────────────────────────────────────────────────────────────────────
//  6. INTERNAL ORDER ID  (continuous numbering, gap-fill on delete)
// ─────────────────────────────────────────────────────────────────────────────

function normalizeKeyPart_(v) {
  return String(v || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function isIndiaCountry_(country) {
  var c = normalizeKeyPart_(country);
  return c === "india" || c === "in";
}

/**
 * For NON-India orders only, build a grouping key:
 * BRAND_NAME + FULLNAME + ADDRESSLINE1 + CITY + STATE + PINCODE + COUNTRY.
 * Returns '' if not eligible.
 */
function buildForeignAddressKey_(
  brand,
  fullName,
  address1,
  city,
  state,
  pincode,
  country,
) {
  if (isIndiaCountry_(country)) return "";
  var parts = [
    normalizeKeyPart_(brand),
    normalizeKeyPart_(fullName),
    normalizeKeyPart_(address1),
    normalizeKeyPart_(city),
    normalizeKeyPart_(state),
    normalizeKeyPart_(pincode),
    normalizeKeyPart_(country),
  ];
  for (var i = 0; i < parts.length; i++) {
    if (!parts[i]) return "";
  }
  return parts.join("||");
}

/**
 * Assigns INTERNAL_ORDERNO for specified rows based on BRAND_NAME.
 * Format: {Initial_code}{sequential_number}
 * @param {Sheet} sheet – New_Orders
 * @param {number[]} rows – rows to assign
 */
function assignInternalOrderNos_(sheet, rows) {
  if (!rows || rows.length === 0) return;

  var brandMap = getBrandConfigMap_();
  var configSheet = getConfigSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  // Read all data across needed columns
  var rowCount = lastRow - DATA_START_ROW + 1;
  var allBrands = sheet
    .getRange(DATA_START_ROW, HEADERS.BRAND_NAME, rowCount, 1)
    .getValues();
  var allOrderNos = sheet
    .getRange(DATA_START_ROW, HEADERS.INTERNAL_ORDERNO, rowCount, 1)
    .getValues();
  var allPortalIds = sheet
    .getRange(DATA_START_ROW, HEADERS.PORTAL_ORDER_ID, rowCount, 1)
    .getValues();
  var allCountries = sheet
    .getRange(DATA_START_ROW, HEADERS.COUNTRY, rowCount, 1)
    .getValues();
  var allFullNames = sheet
    .getRange(DATA_START_ROW, HEADERS.FULLNAME, rowCount, 1)
    .getValues();
  var allAddress1 = sheet
    .getRange(DATA_START_ROW, HEADERS.ADDRESSLINE1, rowCount, 1)
    .getValues();
  var allCities = sheet
    .getRange(DATA_START_ROW, HEADERS.CITY, rowCount, 1)
    .getValues();
  var allStates = sheet
    .getRange(DATA_START_ROW, HEADERS.STATE, rowCount, 1)
    .getValues();
  var allPincodes = sheet
    .getRange(DATA_START_ROW, HEADERS.PINCODE, rowCount, 1)
    .getValues();

  // Find current max per brand and grouping maps
  var brandMaxMap = {};
  var portalIdToOrderNo = {};
  var foreignKeyToOrderNo = {};

  for (var i = 0; i < allBrands.length; i++) {
    var b = String(allBrands[i][0]).trim();
    var pId = String(allPortalIds[i][0]).trim();
    var orderNo = String(allOrderNos[i][0]).trim();

    if (!b || !brandMap[b]) continue;
    var prefix = brandMap[b].Initial_code;

    if (pId && orderNo) {
      portalIdToOrderNo[pId] = orderNo;
    }
    var foreignKey = buildForeignAddressKey_(
      b,
      allFullNames[i][0],
      allAddress1[i][0],
      allCities[i][0],
      allStates[i][0],
      allPincodes[i][0],
      allCountries[i][0],
    );
    if (foreignKey && orderNo) {
      foreignKeyToOrderNo[foreignKey] = orderNo;
    }

    if (orderNo.indexOf(prefix) === 0) {
      var num = parseInt(orderNo.substring(prefix.length), 10);
      if (!isNaN(num)) {
        brandMaxMap[b] = Math.max(brandMaxMap[b] || 0, num);
      }
    }
  }

  // Assign new order numbers for requested rows
  var configUpdates = {}; // brand -> new max

  var writeRows = [];
  var writeVals = [];

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

    var foreignKeyRow = buildForeignAddressKey_(
      brand,
      fullName,
      address1,
      city,
      state,
      pincode,
      country,
    );
    var finalOrderNo = "";
    if (foreignKeyRow && foreignKeyToOrderNo[foreignKeyRow]) {
      finalOrderNo = foreignKeyToOrderNo[foreignKeyRow];
    } else if (pId && portalIdToOrderNo[pId]) {
      finalOrderNo = portalIdToOrderNo[pId];
    } else {
      var prefix = brandMap[brand].Initial_code;
      // Get the absolute max starting point from AllOrders if missing
      if (brandMaxMap[brand] === undefined) {
        brandMaxMap[brand] = Math.max(
          getMaxFromAllOrders_(prefix),
          brandMap[brand].Last_Maximum,
        );
      }
      var nextNum = brandMaxMap[brand] + 1;
      brandMaxMap[brand] = nextNum;

      finalOrderNo = prefix + nextNum;
      if (foreignKeyRow) {
        foreignKeyToOrderNo[foreignKeyRow] = finalOrderNo;
      }
      if (pId) {
        portalIdToOrderNo[pId] = finalOrderNo;
      }
      configUpdates[brand] = nextNum;
    }
    writeRows.push(row);
    writeVals.push([finalOrderNo]);
  }

  for (var w = 0; w < writeRows.length; w++) {
    sheet
      .getRange(writeRows[w], HEADERS.INTERNAL_ORDERNO)
      .setValue(writeVals[w][0]);
  }

  // Batch update Config Last_Maximums
  for (var changedBrand in configUpdates) {
    configSheet
      .getRange(brandMap[changedBrand].configRow, CONFIG_COLS.Last_Maximum)
      .setValue(configUpdates[changedBrand]);
  }
}

/**
 * Regenerates all INTERNAL_ORDERNO values for a given brand.
 * Called after row deletion to fill gaps.
 * @param {Sheet} sheet – New_Orders
 * @param {string} brandName – the brand to resequence (or null for all brands)
 */
function regenerateInternalOrderNos_(sheet, brandName) {
  var brandMap = getBrandConfigMap_();
  var configSheet = getConfigSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  var numRows = lastRow - DATA_START_ROW + 1;
  var allBrands = sheet
    .getRange(DATA_START_ROW, HEADERS.BRAND_NAME, numRows, 1)
    .getValues();
  var orderNos = sheet
    .getRange(DATA_START_ROW, HEADERS.INTERNAL_ORDERNO, numRows, 1)
    .getValues();
  var allPortalIds = sheet
    .getRange(DATA_START_ROW, HEADERS.PORTAL_ORDER_ID, numRows, 1)
    .getValues();
  var allCountries = sheet
    .getRange(DATA_START_ROW, HEADERS.COUNTRY, numRows, 1)
    .getValues();
  var allFullNames = sheet
    .getRange(DATA_START_ROW, HEADERS.FULLNAME, numRows, 1)
    .getValues();
  var allAddress1 = sheet
    .getRange(DATA_START_ROW, HEADERS.ADDRESSLINE1, numRows, 1)
    .getValues();
  var allCities = sheet
    .getRange(DATA_START_ROW, HEADERS.CITY, numRows, 1)
    .getValues();
  var allStates = sheet
    .getRange(DATA_START_ROW, HEADERS.STATE, numRows, 1)
    .getValues();
  var allPincodes = sheet
    .getRange(DATA_START_ROW, HEADERS.PINCODE, numRows, 1)
    .getValues();

  // Group rows by brand
  var brandRows = {};
  for (var i = 0; i < numRows; i++) {
    var b = String(allBrands[i][0]).trim();
    if (!b || !brandMap[b]) continue;
    if (brandName && b !== brandName) continue;
    if (!brandRows[b]) brandRows[b] = [];
    brandRows[b].push(i); // 0-based index relative to DATA_START_ROW
  }

  // Re-sequence each brand
  for (var brand in brandRows) {
    var prefix = brandMap[brand].Initial_code;
    var indices = brandRows[brand];

    // Find current min assigned sequence in New_Orders for this brand to protect against AllOrders missing
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
      var foreignKey = buildForeignAddressKey_(
        brand,
        allFullNames[rIdx][0],
        allAddress1[rIdx][0],
        allCities[rIdx][0],
        allStates[rIdx][0],
        allPincodes[rIdx][0],
        allCountries[rIdx][0],
      );
      var seqNum;

      if (foreignKey && foreignMap[foreignKey]) {
        seqNum = foreignMap[foreignKey];
      } else if (pId && pIdMap[pId]) {
        seqNum = pIdMap[pId];
      } else {
        maxNum++;
        seqNum = maxNum;
        if (foreignKey) {
          foreignMap[foreignKey] = seqNum;
        }
        if (pId) {
          pIdMap[pId] = seqNum;
        }
      }

      orderNos[rIdx][0] = prefix + seqNum;
    }

    // Update Last_Maximum in Config
    configSheet
      .getRange(brandMap[brand].configRow, CONFIG_COLS.Last_Maximum)
      .setValue(maxNum);
  }

  // Batch write order numbers back
  sheet
    .getRange(DATA_START_ROW, HEADERS.INTERNAL_ORDERNO, numRows, 1)
    .setValues(orderNos);
}

/**
 * Full regeneration for all brands — menu callable.
 */
function regenerateAllInternalOrderNos() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    SHEET_NAMES.NEW_ORDERS,
  );
  if (!sheet) return;
  regenerateInternalOrderNos_(sheet, null);
  showStyledAlert_(
    "Order IDs Regenerated",
    "All internal order numbers have been re-sequenced without gaps.",
    "success",
  );
}

/**
 * Automatically resequences internal order IDs after a manual user edit.
 * Reads the newly entered number, and re-sequences rows BELOW it sequentially.
 */
function resequenceFromManualEdit_(sheet, editedRow) {
  var brandMap = getBrandConfigMap_();
  var thisBrand = String(
    sheet.getRange(editedRow, HEADERS.BRAND_NAME).getValue(),
  ).trim();
  if (!thisBrand || !brandMap[thisBrand]) return;

  var newOrderNo = String(
    sheet.getRange(editedRow, HEADERS.INTERNAL_ORDERNO).getValue(),
  ).trim();
  var prefix = brandMap[thisBrand].Initial_code;
  if (newOrderNo.indexOf(prefix) !== 0) return; // Ignore invalid edits

  var startNum = parseInt(newOrderNo.substring(prefix.length), 10);
  if (isNaN(startNum)) return;

  var lastRow = sheet.getLastRow();
  if (lastRow <= editedRow) {
    // Just update the config if they edited the last row
    var configSheet = getConfigSheet_();
    var currentConfigMax =
      Number(
        configSheet
          .getRange(brandMap[thisBrand].configRow, CONFIG_COLS.Last_Maximum)
          .getValue(),
      ) || 0;
    if (startNum > currentConfigMax) {
      configSheet
        .getRange(brandMap[thisBrand].configRow, CONFIG_COLS.Last_Maximum)
        .setValue(startNum);
    }
    return;
  }

  var numRows = lastRow - editedRow;
  var allBrands = sheet
    .getRange(editedRow + 1, HEADERS.BRAND_NAME, numRows, 1)
    .getValues();
  var allPortalIds = sheet
    .getRange(editedRow + 1, HEADERS.PORTAL_ORDER_ID, numRows, 1)
    .getValues();
  var allCountries = sheet
    .getRange(editedRow + 1, HEADERS.COUNTRY, numRows, 1)
    .getValues();
  var allFullNames = sheet
    .getRange(editedRow + 1, HEADERS.FULLNAME, numRows, 1)
    .getValues();
  var allAddress1 = sheet
    .getRange(editedRow + 1, HEADERS.ADDRESSLINE1, numRows, 1)
    .getValues();
  var allCities = sheet
    .getRange(editedRow + 1, HEADERS.CITY, numRows, 1)
    .getValues();
  var allStates = sheet
    .getRange(editedRow + 1, HEADERS.STATE, numRows, 1)
    .getValues();
  var allPincodes = sheet
    .getRange(editedRow + 1, HEADERS.PINCODE, numRows, 1)
    .getValues();
  var orderNos = sheet
    .getRange(editedRow + 1, HEADERS.INTERNAL_ORDERNO, numRows, 1)
    .getValues();

  var currentMax = startNum;
  var pIdMap = {};
  var foreignMap = {};
  var editedPortalId = String(
    sheet.getRange(editedRow, HEADERS.PORTAL_ORDER_ID).getValue(),
  ).trim();
  if (editedPortalId) {
    pIdMap[editedPortalId] = startNum;
  }
  var editedForeignKey = buildForeignAddressKey_(
    thisBrand,
    sheet.getRange(editedRow, HEADERS.FULLNAME).getValue(),
    sheet.getRange(editedRow, HEADERS.ADDRESSLINE1).getValue(),
    sheet.getRange(editedRow, HEADERS.CITY).getValue(),
    sheet.getRange(editedRow, HEADERS.STATE).getValue(),
    sheet.getRange(editedRow, HEADERS.PINCODE).getValue(),
    sheet.getRange(editedRow, HEADERS.COUNTRY).getValue(),
  );
  if (editedForeignKey) {
    foreignMap[editedForeignKey] = startNum;
  }

  for (var i = 0; i < numRows; i++) {
    var b = String(allBrands[i][0]).trim();
    if (b !== thisBrand) continue;

    var pId = String(allPortalIds[i][0]).trim();
    var foreignKey = buildForeignAddressKey_(
      thisBrand,
      allFullNames[i][0],
      allAddress1[i][0],
      allCities[i][0],
      allStates[i][0],
      allPincodes[i][0],
      allCountries[i][0],
    );
    var seqNum;

    if (foreignKey && foreignMap[foreignKey]) {
      seqNum = foreignMap[foreignKey];
    } else if (pId && pIdMap[pId]) {
      seqNum = pIdMap[pId];
    } else {
      currentMax++;
      seqNum = currentMax;
      if (foreignKey) {
        foreignMap[foreignKey] = seqNum;
      }
      if (pId) {
        pIdMap[pId] = seqNum;
      }
    }
    orderNos[i][0] = prefix + seqNum;
  }

  // Write the updated sequence back to sheet
  sheet
    .getRange(editedRow + 1, HEADERS.INTERNAL_ORDERNO, numRows, 1)
    .setValues(orderNos);

  // Update Config
  var configSheet = getConfigSheet_();
  var currentConfigMax =
    Number(
      configSheet
        .getRange(brandMap[thisBrand].configRow, CONFIG_COLS.Last_Maximum)
        .getValue(),
    ) || 0;
  if (currentMax > currentConfigMax) {
    configSheet
      .getRange(brandMap[thisBrand].configRow, CONFIG_COLS.Last_Maximum)
      .setValue(currentMax);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
//  7. FINANCIAL CALCULATIONS
// ─────────────────────────────────────────────────────────────────────────────

var _exchangeRatesCache = {};

/**
 * Fetches conversion rate to INR for a given currency code.
 * Returns cached value when available.
 */
function getConversionRateFromAPI_(currency) {
  var cur = String(currency || "")
    .trim()
    .toUpperCase();
  if (!cur) return 0;
  if (cur === "INR") return 1;
  if (_exchangeRatesCache[cur]) return _exchangeRatesCache[cur];

  try {
    var response = UrlFetchApp.fetch(
      "[REDACTED_API_URL]" + cur,
      { muteHttpExceptions: true },
    );
    if (response.getResponseCode() !== 200) return 0;
    var data = JSON.parse(response.getContentText());
    var rate = data && data.rates ? Number(data.rates.INR) : 0;
    if (rate && !isNaN(rate)) {
      _exchangeRatesCache[cur] = rate;
      return rate;
    }
  } catch (err) {
    Logger.log("getConversionRateFromAPI_ error for " + cur + ": " + err);
  }
  return 0;
}

/**
 * Calculates financial fields for specified rows.
 * Uses OLD Magic Number logic with SHIPPING_CHARGE_PRODUCT fix.
 * @param {Sheet} sheet
 * @param {number[]} rows
 */
function calculateFinancialsForRows_(sheet, rows) {
  if (!rows || rows.length === 0) return;

  var countryMap = getCountryConfigMap_();

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];

    var currency = String(
      sheet.getRange(row, HEADERS.CURRENCY).getValue(),
    ).trim();
    var country = String(
      sheet.getRange(row, HEADERS.COUNTRY).getValue(),
    ).trim();
    var productCost =
      Number(sheet.getRange(row, HEADERS.Product_Cost).getValue()) || 0;

    // Skip visually empty rows
    if (!currency && !country && productCost === 0) {
      sheet.getRange(row, HEADERS.Price_in_INR).setValue("");
      sheet.getRange(row, HEADERS.Shipping_Charge_Product).setValue("");
      sheet.getRange(row, HEADERS.Maximum_Expense).setValue("");
      sheet.getRange(row, HEADERS.Actual_Expense).setValue("");
      sheet.getRange(row, HEADERS.Maximum_Profit).setValue("");
      sheet.getRange(row, HEADERS.Actual_Profit).setValue("");

      continue;
    }

    var currencyPrice =
      Number(sheet.getRange(row, HEADERS.CURRENCY_PRICE).getValue()) || 0;

    var convRateRaw = sheet.getRange(row, HEADERS.Conversion_Rate).getValue();
    var convRate = Number(convRateRaw);

    if (!convRate || isNaN(convRate)) {
      convRate = getConversionRateFromAPI_(currency);

      sheet.getRange(row, HEADERS.Conversion_Rate).setValue(convRate);
    }

    var shippingCharge =
      Number(sheet.getRange(row, HEADERS.SHIPPING_CHARGE).getValue()) || 0;

    var qty = Number(sheet.getRange(row, HEADERS.QTY).getValue()) || 1;

    // Country Config
    var countryInfo = countryMap[country] ||
      countryMap[normalizeCountryKey_(country)] || {
        Charge: 0,
        Magic_Number: 0,
      };

    var shippingChargeProduct = Number(countryInfo.Charge) || 0;

    var magicNumber = Number(countryInfo.Magic_Number) || 0;

    // =========================
    // CALCULATIONS
    // =========================

    var priceInINR = currencyPrice * convRate;

    var maxExpense = priceInINR * magicNumber;

    // FIXED
    var actualExpense = productCost + shippingChargeProduct;

    var maxProfit = priceInINR * 0.2;

    var actualProfit = maxExpense - actualExpense + maxProfit;

    // Write
    sheet.getRange(row, HEADERS.Price_in_INR).setValue(priceInINR);

    sheet
      .getRange(row, HEADERS.Shipping_Charge_Product)
      .setValue(shippingChargeProduct);

    sheet.getRange(row, HEADERS.Maximum_Expense).setValue(maxExpense);

    sheet.getRange(row, HEADERS.Actual_Expense).setValue(actualExpense);

    sheet.getRange(row, HEADERS.Maximum_Profit).setValue(maxProfit);

    sheet.getRange(row, HEADERS.Actual_Profit).setValue(actualProfit);
  }
}

/**
 * Batch version using OLD Magic Number logic
 * with SHIPPING_CHARGE_PRODUCT fix.
 * @param {Sheet} sheet
 * @param {number} startRow
 * @param {number} endRow
 */
function batchCalculateFinancials_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;

  if (numRows <= 0) return;

  var countryMap = getCountryConfigMap_();

  // Batch Read
  var currencyPrices = sheet
    .getRange(startRow, HEADERS.CURRENCY_PRICE, numRows, 1)
    .getValues();

  var convRates = sheet
    .getRange(startRow, HEADERS.Conversion_Rate, numRows, 1)
    .getValues();

  var productCosts = sheet
    .getRange(startRow, HEADERS.Product_Cost, numRows, 1)
    .getValues();

  var shippingCharges = sheet
    .getRange(startRow, HEADERS.SHIPPING_CHARGE, numRows, 1)
    .getValues();

  var countries = sheet
    .getRange(startRow, HEADERS.COUNTRY, numRows, 1)
    .getValues();

  var qtys = sheet.getRange(startRow, HEADERS.QTY, numRows, 1).getValues();

  var currencies = sheet
    .getRange(startRow, HEADERS.CURRENCY, numRows, 1)
    .getValues();

  // Output Arrays
  var convRateUpdates = [];

  var priceInINROut = [];
  var shipProdOut = [];
  var maxExpOut = [];
  var actExpOut = [];
  var maxProfOut = [];
  var actProfOut = [];

  for (var i = 0; i < numRows; i++) {
    var cur = String(currencies[i][0]).trim();
    var ctry = String(countries[i][0]).trim();
    var pc = Number(productCosts[i][0]) || 0;

    // Skip visually empty rows
    if (!cur && !ctry && pc === 0) {
      convRateUpdates.push([""]);

      priceInINROut.push([""]);
      shipProdOut.push([""]);
      maxExpOut.push([""]);
      actExpOut.push([""]);
      maxProfOut.push([""]);
      actProfOut.push([""]);

      continue;
    }

    var cp = Number(currencyPrices[i][0]) || 0;

    var cr = Number(convRates[i][0]);

    if (!cr || isNaN(cr)) {
      cr = getConversionRateFromAPI_(cur);
    }

    convRateUpdates.push([cr]);

    var sc = Number(shippingCharges[i][0]) || 0;

    var qty = Number(qtys[i][0]) || 1;

    var countryInfo = countryMap[ctry] ||
      countryMap[normalizeCountryKey_(ctry)] || {
        Charge: 0,
        Magic_Number: 0,
      };

    var shippingChargeProduct = Number(countryInfo.Charge) || 0;

    var magicNumber = Number(countryInfo.Magic_Number) || 0;

    // =========================
    // CALCULATIONS
    // =========================

    var prInINR = cp * cr;

    var maxExp = prInINR * magicNumber;

    // FIXED
    var actExp = pc + shippingChargeProduct;

    var maxProf = prInINR * 0.2;

    var actProf = maxExp - actExp + maxProf;

    // Push Outputs
    priceInINROut.push([prInINR]);

    shipProdOut.push([shippingChargeProduct]);

    maxExpOut.push([maxExp]);

    actExpOut.push([actExp]);

    maxProfOut.push([maxProf]);

    actProfOut.push([actProf]);
  }

  // Batch Write
  sheet
    .getRange(startRow, HEADERS.Conversion_Rate, numRows, 1)
    .setValues(convRateUpdates);

  sheet
    .getRange(startRow, HEADERS.Price_in_INR, numRows, 1)
    .setValues(priceInINROut);

  sheet
    .getRange(startRow, HEADERS.Shipping_Charge_Product, numRows, 1)
    .setValues(shipProdOut);

  sheet
    .getRange(startRow, HEADERS.Maximum_Expense, numRows, 1)
    .setValues(maxExpOut);

  sheet
    .getRange(startRow, HEADERS.Actual_Expense, numRows, 1)
    .setValues(actExpOut);

  sheet
    .getRange(startRow, HEADERS.Maximum_Profit, numRows, 1)
    .setValues(maxProfOut);

  sheet
    .getRange(startRow, HEADERS.Actual_Profit, numRows, 1)
    .setValues(actProfOut);
}
// ─────────────────────────────────────────────────────────────────────────────
//  8. APPEND TO ALLORDERS (with duplicate check & styled popup)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Returns a Set of all PORTAL_ORDER_IDs already in AllOrders.
 * @return {Object} – keys are order IDs (strings)
 */
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

/**
 * Menu-callable: push selected rows to AllOrders with duplicate check.
 * Shows a styled confirmation popup before pushing.
 */
function appendSelectedToAllOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
  if (!sheet) return;

  var selection = ss.getActiveRange();
  if (!selection) {
    showStyledAlert_(
      "No Selection",
      "Please select the rows you want to push to AllOrders.",
      "warning",
    );
    return;
  }

  var startRow = Math.max(selection.getRow(), DATA_START_ROW);
  var endRow = selection.getLastRow();
  if (endRow < DATA_START_ROW) {
    showStyledAlert_(
      "Invalid Selection",
      "Please select data rows (not the header).",
      "warning",
    );
    return;
  }

  // Eradicate obsolete 'Duplicate' stamps by forcing a live validation refresh before pulling
  validateAndLogErrors_(sheet, startRow, endRow);

  var numRows = endRow - startRow + 1;
  var allData = sheet.getRange(startRow, 1, numRows, TOTAL_COLUMNS).getValues();

  // ── Duplicate check ──
  var existingIds = getExistingAllOrderIds_();
  var duplicates = [];
  var validRows = [];
  var validData = [];

  var duplicatesToMark = [];

  for (var i = 0; i < numRows; i++) {
    var portalId = String(allData[i][HEADERS.PORTAL_ORDER_ID - 1]).trim();
    var errLog = String(allData[i][HEADERS.what_to_fix - 1]).trim();

    if (!portalId) {
      duplicates.push({
        row: startRow + i,
        id: "(empty)",
        reason: "No Portal Order ID",
      });
      continue;
    }
    if (errLog !== "✅ All fields complete") {
      duplicates.push({
        row: startRow + i,
        id: portalId,
        reason: "Incomplete: " + errLog,
      });
      continue;
    }

    if (existingIds[portalId]) {
      duplicates.push({
        row: startRow + i,
        id: portalId,
        reason: "Already exists in AllOrders",
      });
      duplicatesToMark.push(startRow + i);
    } else {
      validRows.push(startRow + i);
      validData.push(allData[i]);
    }
  }

  // Stamp duplicate in sheet for user feedback
  for (var d = 0; d < duplicatesToMark.length; d++) {
    sheet
      .getRange(duplicatesToMark[d], HEADERS.what_to_fix)
      .setValue("Duplicate");
  }

  // ── If ALL are duplicates, block entirely ──
  if (duplicates.length > 0 && validData.length === 0) {
    showDuplicateBlockedPopup_(duplicates);
    return;
  }

  // ── If SOME are duplicates, warn and show which will be skipped ──
  if (duplicates.length > 0) {
    // Store valid data for callback
    PropertiesService.getScriptProperties().setProperty(
      "[REDACTED_PUSH_START_KEY]",
      String(startRow),
    );
    PropertiesService.getScriptProperties().setProperty(
      "[REDACTED_PUSH_END_KEY]",
      String(endRow),
    );
    showPartialDuplicatePopup_(validData, duplicates);
    return;
  }

  // ── No duplicates → store range and show confirmation ──
  PropertiesService.getScriptProperties().setProperty(
    "[REDACTED_PUSH_START_KEY]",
    String(startRow),
  );
  PropertiesService.getScriptProperties().setProperty(
    "[REDACTED_PUSH_END_KEY]",
    String(endRow),
  );
  showPushConfirmationPopup_(validData, numRows);
}

/**
 * Menu-callable: push ALL valid rows from New_Orders to AllOrders.
 */
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

  // Force a validation refresh to cleanse obsolete Duplicate tags natively
  validateAndLogErrors_(sheet, startRow, endRow);

  var numRows = endRow - startRow + 1;
  var allData = sheet.getRange(startRow, 1, numRows, TOTAL_COLUMNS).getValues();

  // ── Duplicate check ──
  var existingIds = getExistingAllOrderIds_();
  var duplicates = [];
  var validRows = [];
  var validData = [];

  var duplicatesToMark = [];

  for (var i = 0; i < numRows; i++) {
    // Only process if there's some data
    var hasData = allData[i].join("").trim() !== "";
    if (!hasData) continue; // skip entirely empty rows

    var portalId = String(allData[i][HEADERS.PORTAL_ORDER_ID - 1]).trim();
    var errLog = String(allData[i][HEADERS.what_to_fix - 1]).trim();

    if (!portalId) {
      duplicates.push({
        row: startRow + i,
        id: "(empty)",
        reason: "No Portal Order ID",
      });
      continue;
    }
    if (errLog !== "✅ All fields complete") {
      duplicates.push({
        row: startRow + i,
        id: portalId,
        reason: "Incomplete: " + errLog,
      });
      continue;
    }

    if (existingIds[portalId]) {
      duplicates.push({
        row: startRow + i,
        id: portalId,
        reason: "Already exists in AllOrders",
      });
      duplicatesToMark.push(startRow + i);
    } else {
      validRows.push(startRow + i);
      validData.push(allData[i]);
    }
  }

  // Stamp duplicate in sheet for user feedback
  for (var d = 0; d < duplicatesToMark.length; d++) {
    sheet
      .getRange(duplicatesToMark[d], HEADERS.what_to_fix)
      .setValue("Duplicate");
  }

  if (validData.length === 0) {
    if (duplicates.length > 0) {
      showDuplicateBlockedPopup_(duplicates);
    } else {
      showStyledAlert_(
        "No Valid Data",
        "No valid orders found to push.",
        "info",
      );
    }
    return;
  }

  PropertiesService.getScriptProperties().setProperty(
    "[REDACTED_PUSH_START_KEY]",
    String(startRow),
  );
  PropertiesService.getScriptProperties().setProperty(
    "[REDACTED_PUSH_END_KEY]",
    String(endRow),
  );

  if (duplicates.length > 0) {
    showPartialDuplicatePopup_(validData, duplicates);
    return;
  }

  showPushConfirmationPopup_(validData, validData.length);
}

/**
 * Called from popup: push only the valid (non-duplicate) rows.
 */
function executePushToAllOrders() {
  var props = PropertiesService.getScriptProperties();
  var startRow = Number(props.getProperty("[REDACTED_PUSH_START_KEY]"));
  var endRow = Number(props.getProperty("[REDACTED_PUSH_END_KEY]"));

  if (!startRow || !endRow) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
  var dest = ss.getSheetByName(SHEET_NAMES.ALL_ORDERS);
  if (!sheet || !dest) return;

  var numRows = endRow - startRow + 1;
  var allData = sheet.getRange(startRow, 1, numRows, TOTAL_COLUMNS).getValues();
  var existingIds = getExistingAllOrderIds_();
  var allOrdersColCount = Math.max(dest.getLastColumn(), 42);
  var stockStatusCol = findHeaderColByAliases_(dest, HEADER_ROW, [
    "stock",
    "stock_availibilty",
    "stock_availability",
    "stockavailability",
  ]);
  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var pushed = 0;

  for (var i = 0; i < numRows; i++) {
    var portalId = String(allData[i][HEADERS.PORTAL_ORDER_ID - 1]).trim();
    if (portalId && !existingIds[portalId]) {
      var outRow = buildAllOrdersRowFromNewOrderRow_(
        allData[i],
        portalMap,
        invMap,
        allOrdersColCount,
        stockStatusCol,
      );
      dest.appendRow(outRow);
      pushed++;
    }
  }

  if (pushed > 0) {
    try {
      refreshAllOrdersStockAvailability_();
    } catch (e1) {
      Logger.log("refreshAllOrdersStockAvailability_ error: " + e1);
    }
  }

  props.deleteProperty("[REDACTED_PUSH_START_KEY]");
  props.deleteProperty("[REDACTED_PUSH_END_KEY]");
}

/**
 * Called from confirmation popup when NO duplicates exist.
 */
function executePushAll() {
  var props = PropertiesService.getScriptProperties();
  var startRow = Number(props.getProperty("[REDACTED_PUSH_START_KEY]"));
  var endRow = Number(props.getProperty("[REDACTED_PUSH_END_KEY]"));

  if (!startRow || !endRow) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
  var dest = ss.getSheetByName(SHEET_NAMES.ALL_ORDERS);
  if (!sheet || !dest) return;

  var numRows = endRow - startRow + 1;
  var allData = sheet.getRange(startRow, 1, numRows, TOTAL_COLUMNS).getValues();
  var existingIds = getExistingAllOrderIds_();
  var allOrdersColCount = Math.max(dest.getLastColumn(), 42);
  var stockStatusCol = findHeaderColByAliases_(dest, HEADER_ROW, [
    "stock",
    "stock_availibilty",
    "stock_availability",
    "stockavailability",
  ]);
  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var pushed = 0;

  for (var i = 0; i < numRows; i++) {
    var portalId = String(allData[i][HEADERS.PORTAL_ORDER_ID - 1]).trim();
    if (portalId && !existingIds[portalId]) {
      var outRow = buildAllOrdersRowFromNewOrderRow_(
        allData[i],
        portalMap,
        invMap,
        allOrdersColCount,
        stockStatusCol,
      );
      dest.appendRow(outRow);
      pushed++;
    }
  }

  if (pushed > 0) {
    try {
      refreshAllOrdersStockAvailability_();
    } catch (e1) {
      Logger.log("refreshAllOrdersStockAvailability_ error: " + e1);
    }
  }

  props.deleteProperty("[REDACTED_PUSH_START_KEY]");
  props.deleteProperty("[REDACTED_PUSH_END_KEY]");
}

function isAllFieldsCompleteMessage_(value) {
  return (
    String(value || "")
      .toLowerCase()
      .indexOf("all fields complete") !== -1
  );
}

function appendRowsToSheet_(sheet, rows, colCount) {
  if (!sheet || !rows || rows.length === 0) return;
  var appendStart = Math.max(sheet.getLastRow() + 1, DATA_START_ROW);
  sheet.getRange(appendStart, 1, rows.length, colCount).setValues(rows);
}

// ── Push Popup: ALL duplicates / incomplete (blocked) ──
function showDuplicateBlockedPopup_(duplicates) {
  var dupRows = "";
  for (var i = 0; i < duplicates.length; i++) {
    dupRows +=
      "<tr>" +
      '<td style="padding:8px 12px;border-bottom:1px solid #ffcdd2;font-size:13px;color:#c62828;">Row ' +
      duplicates[i].row +
      "</td>" +
      '<td style="padding:8px 12px;border-bottom:1px solid #ffcdd2;font-size:13px;font-weight:600;">' +
      duplicates[i].id +
      "</td>" +
      '<td style="padding:8px 12px;border-bottom:1px solid #ffcdd2;font-size:13px;max-width:200px;word-wrap:break-word;">' +
      duplicates[i].reason +
      "</td>" +
      "</tr>";
  }

  var html = _buildPopupShell_(
    "🚫 Import Blocked — Validation Issues",
    "linear-gradient(135deg, #c62828, #e53935)",
    '<table style="width:100%;border-collapse:collapse;">' +
      '<thead><tr style="background:#ffebee;">' +
      '<th style="padding:8px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#c62828;">Row</th>' +
      '<th style="padding:8px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#c62828;">Order ID</th>' +
      '<th style="padding:8px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#c62828;">Reason</th>' +
      "</tr></thead><tbody>" +
      dupRows +
      "</tbody></table>" +
      '<div style="margin-top:14px;padding:12px 16px;background:#ffebee;border-left:4px solid #c62828;' +
      'border-radius:6px;font-size:13px;color:#b71c1c;">' +
      "❌ <strong>All selected orders have validation issues.</strong><br>Check reasons above (incomplete fields or duplicated data). Import cancelled.</div>",
    '<button class="btn" style="background:linear-gradient(135deg,#757575,#9e9e9e);" onclick="google.script.host.close()">Close</button>',
  );

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(540).setHeight(380),
    " ",
  );
}

// ── Push Popup: PARTIAL duplicates / errors (some valid) ──
function showPartialDuplicatePopup_(validData, duplicates) {
  var dupRows = "";
  for (var i = 0; i < duplicates.length; i++) {
    dupRows +=
      "<tr>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #ffcdd2;font-size:12px;">Row ' +
      duplicates[i].row +
      "</td>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #ffcdd2;font-size:12px;font-weight:600;">' +
      duplicates[i].id +
      "</td>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #ffcdd2;font-size:12px;">' +
      '<span title="' +
      duplicates[i].reason.replace(/"/g, "&quot;") +
      '" style="background:#ffcdd2;color:#c62828;padding:2px 8px;border-radius:10px;font-size:11px;cursor:help;display:inline-block;max-width:180px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' +
      duplicates[i].reason +
      "</span></td>" +
      "</tr>";
  }

  var validSummary =
    '<div style="display:flex;gap:16px;margin-bottom:12px;">' +
    '<div style="flex:1;background:#e8f5e9;padding:12px;border-radius:10px;text-align:center;">' +
    '<div style="font-size:28px;font-weight:700;color:#2e7d32;">' +
    validData.length +
    "</div>" +
    '<div style="font-size:11px;color:#558b2f;text-transform:uppercase;letter-spacing:0.5px;">Will Push</div></div>' +
    '<div style="flex:1;background:#ffebee;padding:12px;border-radius:10px;text-align:center;">' +
    '<div style="font-size:28px;font-weight:700;color:#c62828;">' +
    duplicates.length +
    "</div>" +
    '<div style="font-size:11px;color:#b71c1c;text-transform:uppercase;letter-spacing:0.5px;">Issues Found</div></div></div>';

  var html = _buildPopupShell_(
    "⚠️ Validation Issues — Partial Push",
    "linear-gradient(135deg, #e65100, #ff8f00)",
    validSummary +
      '<div style="font-size:13px;font-weight:600;color:#616161;margin-bottom:6px;">Issue rows (will be skipped):</div>' +
      '<table style="width:100%;border-collapse:collapse;table-layout:fixed;">' +
      '<thead><tr><th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#9e9e9e;width:25%;">Row</th>' +
      '<th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#9e9e9e;width:30%;">Order ID</th>' +
      '<th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#9e9e9e;width:45%;">Status</th>' +
      "</tr></thead><tbody>" +
      dupRows +
      "</tbody></table>",
    '<button class="btn" style="background:linear-gradient(135deg,#757575,#9e9e9e);" onclick="google.script.host.close()">Cancel</button>' +
      '<button class="btn" style="background:linear-gradient(135deg,#43a047,#66bb6a);margin-left:10px;" ' +
      "id=\"pushValidBtn\" onclick=\"this.disabled=true;this.style.cursor='not-allowed';document.body.style.cursor='not-allowed';this.textContent='Processing...';google.script.run.withSuccessHandler(function(){document.body.style.cursor='auto';google.script.host.close()}).withFailureHandler(function(e){console.error('Push error:',e);document.body.style.cursor='auto';google.script.host.close()}).executePushToAllOrders()\">" +
      "Push " +
      validData.length +
      " Valid Row(s)</button>",
  );

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(540).setHeight(420),
    " ",
  );
}

// ── Push Popup: NO duplicates (confirm push) ──
function showPushConfirmationPopup_(validData, count) {
  // Build preview table
  var previewRows = "";
  var previewCount = Math.min(count, 6);
  for (var i = 0; i < previewCount; i++) {
    var orderId =
      String(validData[i][HEADERS.PORTAL_ORDER_ID - 1]).trim() || "—";
    var brand = String(validData[i][HEADERS.BRAND_NAME - 1]).trim() || "—";
    var sku = String(validData[i][HEADERS.SKU - 1]).trim() || "—";
    previewRows +=
      "<tr>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #e8eaf6;font-size:12px;">' +
      orderId +
      "</td>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #e8eaf6;font-size:12px;">' +
      brand +
      "</td>" +
      '<td style="padding:6px 10px;border-bottom:1px solid #e8eaf6;font-size:12px;">' +
      sku +
      "</td>" +
      "</tr>";
  }
  if (count > previewCount) {
    previewRows +=
      '<tr><td colspan="3" style="padding:8px;text-align:center;color:#9e9e9e;font-size:12px;">' +
      "...and " +
      (count - previewCount) +
      " more</td></tr>";
  }

  var html = _buildPopupShell_(
    "📋 Push " + count + " Order(s) to AllOrders",
    "linear-gradient(135deg, #1565c0, #42a5f5)",
    '<div style="background:#e3f2fd;padding:12px;border-radius:10px;text-align:center;margin-bottom:12px;">' +
      '<div style="font-size:32px;font-weight:700;color:#1565c0;">' +
      count +
      "</div>" +
      '<div style="font-size:12px;color:#1976d2;text-transform:uppercase;letter-spacing:0.5px;">Orders Ready</div></div>' +
      '<table style="width:100%;border-collapse:collapse;">' +
      '<thead><tr style="background:#e8eaf6;">' +
      '<th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#5c6bc0;">Order ID</th>' +
      '<th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#5c6bc0;">Brand</th>' +
      '<th style="padding:6px 10px;text-align:left;font-size:10px;text-transform:uppercase;color:#5c6bc0;">SKU</th>' +
      "</tr></thead><tbody>" +
      previewRows +
      "</tbody></table>" +
      '<div style="margin-top:10px;padding:10px 14px;background:#e8f5e9;border-left:4px solid #43a047;' +
      'border-radius:6px;font-size:12px;color:#2e7d32;">' +
      "✅ No duplicates found. All orders are safe to push.</div>",
    '<button class="btn" style="background:linear-gradient(135deg,#757575,#9e9e9e);" onclick="google.script.host.close()">Cancel</button>' +
      '<button class="btn" style="background:linear-gradient(135deg,#1565c0,#42a5f5);margin-left:10px;" ' +
      "id=\"confirmPushBtn\" onclick=\"this.disabled=true;this.style.cursor='not-allowed';document.body.style.cursor='not-allowed';this.textContent='Processing...';google.script.run.withSuccessHandler(function(){document.body.style.cursor='auto';google.script.host.close()}).withFailureHandler(function(e){console.error('Push error:',e);document.body.style.cursor='auto';google.script.host.close()}).executePushAll()\">" +
      "Confirm Push</button>",
  );

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(520).setHeight(440),
    " ",
  );
}

/**
 * Reusable HTML popup shell builder.
 */
function _buildPopupShell_(title, gradient, bodyContent, buttonsHtml) {
  return (
    "<!DOCTYPE html><html><head>" +
    '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">' +
    "<style>" +
    "* { margin:0; padding:0; box-sizing:border-box; }" +
    'body { font-family:"Inter",sans-serif; background:#f5f5f5; padding:0; }' +
    ".card { background:#fff; border-radius:16px; overflow:hidden;" +
    "  box-shadow:0 8px 32px rgba(0,0,0,0.12); animation:slideUp 0.3s cubic-bezier(0.22,1,0.36,1); }" +
    ".header { background:" +
    gradient +
    "; padding:18px 22px; color:#fff; }" +
    ".header h2 { font-size:17px; font-weight:700; letter-spacing:-0.3px; }" +
    ".body { padding:18px 22px; }" +
    ".btn-row { padding:14px 22px; display:flex; gap:10px; justify-content:flex-end; background:#fafafa;" +
    "  border-top:1px solid #f0f0f0; }" +
    '.btn { padding:10px 22px; border:none; border-radius:8px; font-family:"Inter",sans-serif;' +
    "  font-size:13px; font-weight:600; cursor:pointer; color:#fff;" +
    "  transition:transform 0.15s,box-shadow 0.15s; box-shadow:0 4px 12px rgba(0,0,0,0.12); }" +
    ".btn:hover { transform:translateY(-1px); box-shadow:0 6px 18px rgba(0,0,0,0.18); }" +
    "@keyframes slideUp { from { opacity:0; transform:translateY(16px); }" +
    "  to { opacity:1; transform:translateY(0); } }" +
    "</style></head><body>" +
    '<div class="card">' +
    '  <div class="header"><h2>' +
    title +
    "</h2></div>" +
    '  <div class="body">' +
    bodyContent +
    "</div>" +
    '  <div class="btn-row">' +
    buttonsHtml +
    "</div>" +
    "</div></body></html>"
  );
}

// ─────────────────────────────────────────────────────────────────────────────
//  9. ONEDIT TRIGGER  (bulk-aware: paste, drag-fill, multi-cell)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Automatically populates BRAND_NAME if the PORTAL_ORDER_ID matches an existing row.
 */
function fillBrandFromPortalId_(sheet, rows) {
  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return false;

  var numRows = lastRow - DATA_START_ROW + 1;
  var allPortals = sheet
    .getRange(DATA_START_ROW, HEADERS.PORTAL_ORDER_ID, numRows, 1)
    .getValues();
  var allBrands = sheet
    .getRange(DATA_START_ROW, HEADERS.BRAND_NAME, numRows, 1)
    .getValues();

  var portalToBrandMap = {};
  for (var i = 0; i < numRows; i++) {
    var pid = String(allPortals[i][0]).trim();
    var b = String(allBrands[i][0]).trim();
    if (pid && b && !portalToBrandMap[pid]) {
      portalToBrandMap[pid] = b;
    }
  }

  var brandChanged = false;
  for (var j = 0; j < rows.length; j++) {
    var r = rows[j];
    var pid = String(
      sheet.getRange(r, HEADERS.PORTAL_ORDER_ID).getValue(),
    ).trim();
    var existingBrand = String(
      sheet.getRange(r, HEADERS.BRAND_NAME).getValue(),
    ).trim();
    if (pid && portalToBrandMap[pid] && !existingBrand) {
      sheet.getRange(r, HEADERS.BRAND_NAME).setValue(portalToBrandMap[pid]);
      brandChanged = true;
    }
  }
  return brandChanged;
}

/**
 * Property key to track row count for deletion detection.
 */
var ROW_COUNT_KEY = "[REDACTED_ROW_COUNT_KEY]";

function onEdit(e) {
  if (
    isLimitedAuthEditEvent_(e) &&
    PropertiesService.getScriptProperties().getProperty(
      NEW_ORDERS_EDIT_TRIGGER_INSTALLED_KEY,
    ) === "true"
  ) {
    return;
  }
  handleNewOrdersEdit_(e);
}

function handleNewOrdersAuthorizedEdit(e) {
  handleNewOrdersEdit_(e);
}

function isLimitedAuthEditEvent_(e) {
  return e && e.authMode && String(e.authMode) === "LIMITED";
}

function installNewOrdersEditTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === NEW_ORDERS_EDIT_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger(NEW_ORDERS_EDIT_TRIGGER_HANDLER)
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  PropertiesService.getScriptProperties().setProperty(
    NEW_ORDERS_EDIT_TRIGGER_INSTALLED_KEY,
    "true",
  );
  showStyledAlert_(
    "Edit Trigger Installed",
    "New Orders edits can now read the external MatchingTable and Inventory spreadsheet.",
    "success",
  );
}

function installWorkflowChangeTrigger(suppressUi) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === WORKFLOW_CHANGE_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger(WORKFLOW_CHANGE_TRIGGER_HANDLER)
    .forSpreadsheet(ss)
    .onChange()
    .create();

  PropertiesService.getScriptProperties().setProperty(
    WORKFLOW_CHANGE_TRIGGER_INSTALLED_KEY,
    "true",
  );

  if (!suppressUi) {
    showStyledAlert_(
      "Change Trigger Installed",
      "Row insert/delete changes will now auto-sync All Orders with stockcheckRTS.",
      "success",
    );
  }
}

/**
 * Main onEdit trigger — handles all edit types.
 */
function handleNewOrdersEdit_(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // STOCKCHECK_RTS onEdit handling is intentionally disabled to avoid
  // laggy behavior when users type 'OK' / 'Faulty' / 'Not Found'.
  // Use the `StockChecked` menu item or the `stockCheckedButton` drawing
  // to process stockcheck rows in batch instead.
  if (sheetName === SHEET_NAMES.STOCKCHECK_RTS) {
    return;
  }

  // Route both spellings of the Dispatched sheet ('Dispatched' and 'Dispached')
  if (sheetName === SHEET_NAMES.DISPATCHED || sheetName === "[REDACTED_DISPACHED_TYPO_SHEET]") {
    try {
      handleDispatchedWorkflowEdit_(e);
    } catch (dispatchErr) {
      Logger.log("handleDispatchedWorkflowEdit_ error: " + dispatchErr);
    }
    return;
  }

  if (sheetName === SHEET_NAMES.CREATE_LABEL) {
    var clStartRow = Math.max(e.range.getRow(), WORKFLOW_DATA_START_ROW);
    var clEndRow = e.range.getLastRow();
    var clStartCol = e.range.getColumn();
    var clEndCol = e.range.getLastColumn();
    // Disabled automatic CreateLabels → AllOrders sync on edit because
    // the CreateLabels UI is handled explicitly via `stockChecked()`
    // and the menu/button actions. To sync CreateLabels to All Orders,
    // run StockChecked (menu) or use the 'Send to Dispatch' action.
    // NOTE: Previously this auto-moved rows from CreateLabels to Dispatched
    // on every edit. Disable automatic movement; provide a menu action
    // `Send to Dispatch` which the user can run manually.
    // The legacy function `labelMoveDoneRowsToDispatched_` remains available
    // and can be invoked via the new menu item.
    // However, handle IMAGE_URL edits so images render immediately.
    try {
      if (
        clEndRow >= WORKFLOW_DATA_START_ROW &&
        isColumnInRange_(CREATE_LABEL_COLS.IMAGE_URL, clStartCol, clEndCol)
      ) {
        try {
          updateCreateLabelsImageFormulas_(sheet, clStartRow, clEndRow);
        } catch (imgErr) {
          Logger.log(
            "updateCreateLabelsImageFormulas_ (on-edit) error: " + imgErr,
          );
        }
      }
    } catch (innerErr) {
      Logger.log("CreateLabels on-edit handler error: " + innerErr);
    }
    return;
  }

  // MatchingTable edits: refresh SKU mapping/costs in New_Orders
  if (sheetName === SHEET_NAMES.MATCHING_TABLE) {
    var mtStartCol = e.range.getColumn();
    var mtEndCol = e.range.getLastColumn();
    if (
      isColumnInRange_(MATCHING_COLS.VendorSKU, mtStartCol, mtEndCol) ||
      isColumnInRange_(MATCHING_COLS.PortalSKU, mtStartCol, mtEndCol)
    ) {
      refreshAllNewOrdersFromCatalog_();
      refreshAllOrdersStockAvailability_();
    }
    return;
  }

  // Inventory edits:
  // 1) IMAGE_URL edits -> update Inventory IMAGE formulas
  // 2) SKU/cost/stock edits -> refresh New_Orders product lookup outputs
  if (sheetName === SHEET_NAMES.INVENTORY) {
    var invStartRow = Math.max(2, e.range.getRow());
    var invEndRow = e.range.getLastRow();
    var invStartCol = e.range.getColumn();
    var invEndCol = e.range.getLastColumn();

    if (
      invEndRow >= 2 &&
      isColumnInRange_(INVENTORY_COLS.IMAGE_URL, invStartCol, invEndCol)
    ) {
      updateInventoryImageFormulas_(sheet, invStartRow, invEndRow);
    }

    if (
      isColumnInRange_(INVENTORY_COLS.VendorSKU, invStartCol, invEndCol) ||
      isColumnInRange_(INVENTORY_COLS.Product_Cost, invStartCol, invEndCol) ||
      isColumnInRange_(INVENTORY_COLS.Stock, invStartCol, invEndCol)
    ) {
      refreshAllNewOrdersFromCatalog_();
      refreshAllOrdersStockAvailability_();
    }
    return;
  }

  // All Orders edits: keep stock availability live on SKU/QTY changes and pastes.
  if (sheetName === SHEET_NAMES.ALL_ORDERS) {
    var aoStartCol = e.range.getColumn();
    var aoEndCol = e.range.getLastColumn();
    if (
      isColumnInRange_(HEADERS.SKU, aoStartCol, aoEndCol) ||
      isColumnInRange_(HEADERS.QTY, aoStartCol, aoEndCol)
    ) {
      refreshAllOrdersStockAvailability_();
    }
    return;
  }

  // Only process New_Orders edits below
  if (sheetName === SHEET_NAMES.CONFIG) {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var newOrdersSheet = ss.getSheetByName(SHEET_NAMES.NEW_ORDERS);
      if (newOrdersSheet && newOrdersSheet.getLastRow() >= DATA_START_ROW) {
        applyValidationsToRange_(
          newOrdersSheet,
          DATA_START_ROW,
          newOrdersSheet.getLastRow(),
        );
      }
      applyWorkflowValidations_();
    } catch (err) {
      Logger.log("Config onEdit update error: " + err);
    }
    return;
  }

  if (sheetName !== SHEET_NAMES.NEW_ORDERS) return;

  var startRow = e.range.getRow();
  var endRow = e.range.getLastRow();
  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();

  // Ignore header row edits
  if (endRow < DATA_START_ROW) return;
  if (startRow < DATA_START_ROW) startRow = DATA_START_ROW;

  // Collect all affected rows
  var rows = [];
  for (var r = startRow; r <= endRow; r++) {
    rows.push(r);
  }

  // ── Row deletion detection ──
  var props = PropertiesService.getScriptProperties();
  var prevRowCount = Number(props.getProperty(ROW_COUNT_KEY)) || 0;
  var currentRowCount = sheet.getLastRow();

  if (prevRowCount > 0 && currentRowCount < prevRowCount) {
    handleRowDeletion_(sheet);
  }
  props.setProperty(ROW_COUNT_KEY, String(currentRowCount));

  // ── Detect which columns were edited ──
  var skuEdited = isColumnInRange_(HEADERS.SKU, startCol, endCol);
  var brandEdited = isColumnInRange_(HEADERS.BRAND_NAME, startCol, endCol);
  var qtyEdited = isColumnInRange_(HEADERS.QTY, startCol, endCol);
  var imageUrlEdited = isColumnInRange_(HEADERS.IMAGE_URL, startCol, endCol);
  var internalOrderNoEdited = isColumnInRange_(
    HEADERS.INTERNAL_ORDERNO,
    startCol,
    endCol,
  );
  var currencyEdited = isColumnInRange_(HEADERS.CURRENCY, startCol, endCol);
  var portalIdEdited = isColumnInRange_(
    HEADERS.PORTAL_ORDER_ID,
    startCol,
    endCol,
  );
  var countryEdited = isColumnInRange_(HEADERS.COUNTRY, startCol, endCol);
  var fullNameEdited = isColumnInRange_(HEADERS.FULLNAME, startCol, endCol);
  var address1Edited = isColumnInRange_(HEADERS.ADDRESSLINE1, startCol, endCol);
  var cityEdited = isColumnInRange_(HEADERS.CITY, startCol, endCol);
  var stateEdited = isColumnInRange_(HEADERS.STATE, startCol, endCol);
  var pincodeEdited = isColumnInRange_(HEADERS.PINCODE, startCol, endCol);
  var editedRowCount = rows.length;
  var editedColCount = endCol - startCol + 1;
  var isBulkEdit = editedRowCount > 5 || editedColCount > 4;

  // Quick first pass so pasted rows immediately get dropdown rules + red/missing markers.
  refreshEditedRows_(sheet, startRow, endRow, true);

  // Normalize dropdown text case-insensitively (e.g. "usd" -> "USD") for pasted/manual values.
  try {
    normalizeDropdownValuesInRange_(sheet, startRow, endRow, startCol, endCol);
  } catch (normErr) {
    Logger.log("normalizeDropdownValuesInRange_ error: " + normErr);
  }

  // ── PORTAL ID changed → populate BRAND_NAME from historic matching row ──
  try {
    if (portalIdEdited) {
      if (fillBrandFromPortalId_(sheet, rows)) {
        brandEdited = true; // explicitly trigger the brand assignments down below seamlessly!
      }
    }
  } catch (portalErr) {
    Logger.log("fillBrandFromPortalId_ error: " + portalErr);
  }

  // ── Manual Internal Order ID edit → resequence below rows ──
  try {
    if (internalOrderNoEdited && editedRowCount === 1) {
      resequenceFromManualEdit_(sheet, startRow);
    }
  } catch (manualReseqErr) {
    Logger.log("resequenceFromManualEdit_ error: " + manualReseqErr);
  }

  // ── Currency changed → clear Conversion Rate so it auto-fetches ──
  try {
    if (currencyEdited) {
      sheet
        .getRange(startRow, HEADERS.Conversion_Rate, rows.length, 1)
        .clearContent();
    }
  } catch (currencyErr) {
    Logger.log("clear conversion rate error: " + currencyErr);
  }

  // ── SKU changed → update product info ──
  // ── IMAGE_URL changed → generate IMAGE formula ──
  try {
    if (imageUrlEdited) {
      updateImageFormulas_(sheet, startRow, endRow);
    }
    var skuRowsToRefresh = [];
    if (skuEdited || editedRowCount > 1) {
      skuRowsToRefresh = rows.slice();
    } else {
      skuRowsToRefresh = getRowsNeedingSkuRefresh_(sheet, startRow, endRow);
    }
    if (skuRowsToRefresh.length > 0) {
      if (skuRowsToRefresh.length > 3) {
        var minRow = Math.min.apply(null, skuRowsToRefresh);
        var maxRow = Math.max.apply(null, skuRowsToRefresh);
        batchUpdateProductInfo_(sheet, minRow, maxRow, true);
      } else {
        updateProductInfoForRows_(sheet, skuRowsToRefresh, true);
      }
    }
  } catch (skuErr) {
    Logger.log("SKU refresh error: " + skuErr);
  }

  // ── BRAND_NAME changed → assign internal order nos ──
  var internalIdGroupingEdited =
    brandEdited ||
    portalIdEdited ||
    countryEdited ||
    fullNameEdited ||
    address1Edited ||
    cityEdited ||
    stateEdited ||
    pincodeEdited;
  try {
    if (internalIdGroupingEdited) {
      assignInternalOrderNos_(sheet, rows);
      // Full resequence on every large paste is expensive and can kill trigger execution.
      // Keep it for smaller edits; deletions are still handled by onChange/onDelete logic.
      if (!isBulkEdit) {
        regenerateInternalOrderNos_(sheet, null);
      }
    }
  } catch (internalIdErr) {
    Logger.log("Internal order assignment error: " + internalIdErr);
  }

  // ── Set ORDER_STATUS = 'Pending' for new rows (batch mode for paste safety) ──
  try {
    var statusData = sheet.getRange(startRow, 1, rows.length, 30).getValues();
    var statusOut = [];
    var statusChanged = false;
    for (var i = 0; i < rows.length; i++) {
      var rawData = statusData[i];
      var hasAnyData = false;
      for (var c = 0; c < 30; c++) {
        if (String(rawData[c]).trim() !== "") {
          hasAnyData = true;
          break;
        }
      }
      var ds = String(rawData[HEADERS.ORDER_STATUS - 1] || "").trim();
      if (hasAnyData && ds === "") {
        statusOut.push(["Pending"]);
        statusChanged = true;
      } else {
        statusOut.push([rawData[HEADERS.ORDER_STATUS - 1]]);
      }
    }
    if (statusChanged) {
      sheet
        .getRange(startRow, HEADERS.ORDER_STATUS, rows.length, 1)
        .setValues(statusOut);
    }
  } catch (statusErr) {
    Logger.log("Order status defaulting error: " + statusErr);
  }

  // ── Stock decrement on SKU + QTY entry ──
  // if (skuEdited || qtyEdited) {
  //   decrementStockForRows_(sheet, rows);
  // }

  // ═══════════════════════════════════════════════════════════════════════
  //  AUTO-RUN ON EVERY EDIT: validations, financials, order IDs, errors
  // ═══════════════════════════════════════════════════════════════════════

  // Final pass: ensure corrected pasted rows clear red/errors instantly after all updates.
  refreshEditedRows_(sheet, startRow, endRow, false);
}

/**
 * Centralized range refresh for onEdit:
 * - Applies dropdown validations
 * - Recalculates financials (optional)
 * - Rebuilds mandatory-field errors + red highlights
 */
function refreshEditedRows_(sheet, startRow, endRow, skipFinancials) {
  if (!sheet || startRow > endRow) return;

  try {
    applyValidationsToRange_(sheet, startRow, endRow);
  } catch (valRuleErr) {
    Logger.log("applyValidationsToRange_ error: " + valRuleErr);
  }

  if (!skipFinancials) {
    try {
      var rowCount = endRow - startRow + 1;
      if (rowCount > 3) {
        batchCalculateFinancials_(sheet, startRow, endRow);
      } else {
        var rows = [];
        for (var r = startRow; r <= endRow; r++) rows.push(r);
        calculateFinancialsForRows_(sheet, rows);
      }
    } catch (finErr) {
      Logger.log("financial refresh error: " + finErr);
    }
  }

  try {
    validateAndLogErrors_(sheet, startRow, endRow);
  } catch (validateErr) {
    Logger.log("validateAndLogErrors_ error: " + validateErr);
  }
}

/**
 * Refreshes all New_Orders rows after MatchingTable/Inventory updates.
 * This clears stale SKU not found errors once mapping/cost rows are fixed.
 */
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
  for (var i = 0; i < aliases.length; i++) {
    aliasSet[String(aliases[i]).trim().toLowerCase().replace(/\s+/g, "")] =
      true;
  }
  for (var c = 0; c < headers.length; c++) {
    var norm = String(headers[c] || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "");
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

  var stockStatusCol = findHeaderColByAliases_(sheet, HEADER_ROW, [
    "stock",
    "stock_availibilty",
    "stock_availability",
    "stockavailability",
  ]);
  if (!stockStatusCol) return;

  var numRows = lastRow - DATA_START_ROW + 1;
  var skuVals = sheet
    .getRange(DATA_START_ROW, HEADERS.SKU, numRows, 1)
    .getValues();
  var qtyVals = sheet
    .getRange(DATA_START_ROW, HEADERS.QTY, numRows, 1)
    .getValues();

  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var out = [];
  var bg = [];

  for (var i = 0; i < numRows; i++) {
    var sku = String(skuVals[i][0]).trim();
    var qty = Number(qtyVals[i][0]) || 0;
    if (!sku) {
      out.push([""]);
      bg.push([CLEAR_COLOR]);
      continue;
    }
    var vendor = portalMap[sku];
    if (!vendor || !invMap[vendor]) {
      out.push([""]);
      bg.push([CLEAR_COLOR]);
      continue;
    }
    var status = getStockAvailabilityStatus_(invMap[vendor].Stock, qty);
    out.push([status]);
    if (status.indexOf("Out of stock") === 0) {
      bg.push(["#ffe0b2"]); // orange for out of stock
    } else {
      bg.push([CLEAR_COLOR]);
    }
  }

  var stockStatusRange = sheet.getRange(
    DATA_START_ROW,
    stockStatusCol,
    numRows,
    1,
  );
  stockStatusRange.setValues(out);
  stockStatusRange.setBackgrounds(bg);
  stockStatusRange.setWrap(true);
}

function normalizeWorkflowKey_(value) {
  return String(value == null ? "" : value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function normalizeWorkflowHeaderKey_(value) {
  return String(value == null ? "" : value)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getWorkflowHeaderMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = normalizeWorkflowHeaderKey_(headers[i]);
    if (key && !Object.prototype.hasOwnProperty.call(map, key)) {
      map[key] = i + 1; // 1-based
    }
  }
  return map;
}

function getHeaderColFromMap_(headerMap, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var key = normalizeWorkflowHeaderKey_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(headerMap, key)) {
      return headerMap[key];
    }
  }
  return 0;
}

function getAllOrdersBundle_() {
  var ss = getWorkflowSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_NAMES.ALL_ORDERS);
  if (!sheet) return null;

  var headerRow = detectHeaderRowByAliases_(
    sheet,
    ["PORTAL_ORDER_ID", "Internal OrderNo", "ORDER_STATUS"],
    5,
  );
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < headerRow || lastCol < 1) return null;

  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var dataStart = headerRow + 1;
  var numRows = Math.max(0, lastRow - headerRow);
  var data =
    numRows > 0
      ? sheet.getRange(dataStart, 1, numRows, lastCol).getValues()
      : [];

  return {
    sheet: sheet,
    headerRow: headerRow,
    headers: headers,
    headerMap: getWorkflowHeaderMap_(headers),
    dataStartRow: dataStart,
    data: data,
  };
}

function getWorkflowSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  var ss = null;
  if (
    active &&
    (active.getSheetByName(SHEET_NAMES.ALL_ORDERS) ||
      active.getSheetByName(SHEET_NAMES.STOCKCHECK_RTS) ||
      active.getSheetByName(SHEET_NAMES.NEW_ORDERS))
  ) {
    ss = active;
  } else {
    ss = SpreadsheetApp.openById(SHIPPING_TARGET_SPREADSHEET_ID);
  }

  // Normalize dispatched sheet name to prefer correct spelling when present.
  // If 'Dispatched' exists, use that; otherwise fall back to existing value.
  try {
    if (ss && ss.getSheetByName && ss.getSheetByName("[REDACTED_DISPATCHED_SHEET]")) {
      SHEET_NAMES.DISPATCHED = "[REDACTED_DISPATCHED_SHEET]";
    } else if (ss && ss.getSheetByName && ss.getSheetByName("[REDACTED_DISPACHED_TYPO_SHEET]")) {
      SHEET_NAMES.DISPATCHED = "[REDACTED_DISPACHED_TYPO_SHEET]";
    }
  } catch (e) {
    // ignore errors during normalization
  }

  return ss;
}

function parseAvailableStockQty_(stockValue) {
  if (stockValue === null || stockValue === undefined || stockValue === "")
    return NaN;

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

function buildInternalSkuKey_(internalOrderNo, sku) {
  return (
    normalizeWorkflowKey_(internalOrderNo) + "||" + normalizeWorkflowKey_(sku)
  );
}

function buildQueueKey_(portalOrderId, internalOrderNo, sku) {
  var internalSku = buildInternalSkuKey_(internalOrderNo, sku);
  var portal = normalizeWorkflowKey_(portalOrderId);
  return internalSku + "||" + portal;
}

function getActiveInternalSkuKeysFromSheet_(
  sheet,
  internalCol,
  skuCol,
  startRow,
) {
  var keys = {};
  if (!sheet) return keys;

  var firstDataRow = startRow || WORKFLOW_DATA_START_ROW;
  var lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) return keys;

  var numRows = lastRow - firstDataRow + 1;
  var values = sheet
    .getRange(firstDataRow, 1, numRows, Math.max(internalCol, skuCol))
    .getValues();
  for (var i = 0; i < values.length; i++) {
    var internalOrderNo = values[i][internalCol - 1];
    var sku = values[i][skuCol - 1];
    var key = buildInternalSkuKey_(internalOrderNo, sku);
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
    if (!inv) continue;
    locationByPortalSku[normalizeWorkflowKey_(portalSku)] = String(
      inv.Location || "",
    ).trim();
  }
  return locationByPortalSku;
}

var _workflowCourierConfigCache = null;
function getWorkflowCourierConfig_() {
  if (_workflowCourierConfigCache) return _workflowCourierConfigCache;

  var cfg = {
    shipglobalCountries: {},
    shiprocketCountries: {},
    salesChannelCourierMap: {},
  };

  var sheet = getConfigSheet_();
  if (!sheet) return cfg;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return cfg;

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0];

  var shipglobalCol = 0;
  var shiprocketCol = 0;
  var salesChannelCol = 0;
  var courierNameCol = 0;

  for (var c = 0; c < headers.length; c++) {
    var key = normalizeWorkflowHeaderKey_(headers[c]);
    if (key === "shipglobal" && !shipglobalCol) shipglobalCol = c + 1;
    if (key === "shiprocket" && !shiprocketCol) shiprocketCol = c + 1;
    if (key === "saleschannel") salesChannelCol = c + 1; // keep latest
    if (key === "couriername") courierNameCol = c + 1; // keep latest
  }

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (shipglobalCol) {
      var g = normalizeWorkflowKey_(row[shipglobalCol - 1]);
      if (g) cfg.shipglobalCountries[g] = true;
    }
    if (shiprocketCol) {
      var sr = normalizeWorkflowKey_(row[shiprocketCol - 1]);
      if (sr) cfg.shiprocketCountries[sr] = true;
    }
    if (salesChannelCol && courierNameCol) {
      var ch = normalizeWorkflowKey_(row[salesChannelCol - 1]);
      var cn = String(row[courierNameCol - 1] || "").trim();
      if (ch && cn) cfg.salesChannelCourierMap[ch] = cn;
  }
  }

  _workflowCourierConfigCache = cfg;
  return cfg;
}

function resolveCourierNameForOrder_(country, salesChannel) {
  var cfg = getWorkflowCourierConfig_();
  var countryKey = normalizeWorkflowKey_(country);
  if (countryKey === "india" || countryKey === "in") {
    var channelKey = normalizeWorkflowKey_(salesChannel);
    return channelKey && cfg.salesChannelCourierMap[channelKey]
      ? cfg.salesChannelCourierMap[channelKey]
      : "";
  }
  if (cfg.shipglobalCountries[countryKey]) return "Shipglobal";
  if (cfg.shiprocketCountries[countryKey]) return "Shiprocket";
  return "";
}

function syncPendingOrdersToStockcheckRTS() {
  try {
    var result = syncPendingOrdersToStockcheckRTSUsingReserve_(false, true);
    showStyledAlert_(
      "Queue Synced",
      "Added: " +
        result.added +
        "<br>" +
        "Removed: " +
        result.removed +
        "<br>" +
        "Eligible pending rows: " +
        result.eligible,
      "success",
    );
  } catch (e) {
    Logger.log("syncPendingOrdersToStockcheckRTS error: " + e);
    showStyledAlert_(
      "Queue Sync Failed",
      String(e && e.message ? e.message : e),
      "error",
    );
  }
}

function syncPendingOrdersToStockcheckRTS_(suppressUi) {
  var ss = getWorkflowSpreadsheet_();
  var stockcheckSheet = ss.getSheetByName(SHEET_NAMES.STOCKCHECK_RTS);
  if (!stockcheckSheet) throw new Error("stockcheckRTS sheet not found.");

  var allBundle = getAllOrdersBundle_();
  if (!allBundle) return { added: 0, removed: 0, eligible: 0 };

  var hm = allBundle.headerMap;
  var portalCol = getHeaderColFromMap_(hm, [
    "PORTAL_ORDER_ID",
    "PortalOrderID",
    "Portal Order ID",
  ]);
  var orderStatusCol = getHeaderColFromMap_(hm, [
    "ORDER_STATUS",
    "Order Status",
  ]);
  var internalCol = getHeaderColFromMap_(hm, [
    "INTERNAL_ORDERNO",
    "Internal OrderNo",
    "InternalOrderNo",
  ]);
  var staffNotesCol = getHeaderColFromMap_(hm, ["STAFF_NOTES", "Staff Notes"]);
  var qtyCol = getHeaderColFromMap_(hm, ["QTY", "Qty", "Quantity"]);
  var imageCol = getHeaderColFromMap_(hm, ["IMAGE", "Image"]);
  var imageUrlCol = getHeaderColFromMap_(hm, [
    "IMAGE_URL",
    "Image URL",
    "Image Url",
  ]);
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  var stockCol = getHeaderColFromMap_(hm, [
    "stock",
    "stock_availibilty",
    "stock_availability",
    "stockavailability",
  ]);
  var salesChannelCol = getHeaderColFromMap_(hm, [
    "SALES_CHANNEL",
    "Sales Channel",
    "SalesChannel",
  ]);
  var countryCol = getHeaderColFromMap_(hm, ["COUNTRY", "Country"]);

  if (!portalCol || !orderStatusCol || !internalCol || !qtyCol || !skuCol) {
    throw new Error(
      "All Orders header mapping is incomplete. Please verify headers.",
    );
  }

  var locationByPortalSku = getInventoryLocationByPortalSkuMap_();
  var activeCreateLabelKeys = getActiveInternalSkuKeysFromSheet_(
    ss.getSheetByName(SHEET_NAMES.CREATE_LABEL),
    CREATE_LABEL_COLS.INTERNAL_ORDERNO,
    CREATE_LABEL_COLS.SKU,
    WORKFLOW_DATA_START_ROW,
  );
  var activeDispatchedKeys = getActiveInternalSkuKeysFromSheet_(
    ss.getSheetByName(SHEET_NAMES.DISPATCHED),
    DISPATCHED_COLS.INTERNAL_ORDERNO,
    DISPATCHED_COLS.SKU,
    WORKFLOW_DATA_START_ROW,
  );

  // Build a map of quantities already queued in stockcheckRTS per portal SKU
  // so we don't queue more than available stock when multiple orders request
  // the same SKU. The map keys are normalized SKU strings.
  var queuedBySku = {};
  try {
    var scLastRow = stockcheckSheet.getLastRow();
    if (scLastRow >= WORKFLOW_DATA_START_ROW) {
      var scNumRows = scLastRow - WORKFLOW_DATA_START_ROW + 1;
      var scData = stockcheckSheet
        .getRange(
          WORKFLOW_DATA_START_ROW,
          1,
          scNumRows,
          STOCKCHECK_COLS.DoCreateLabels,
        )
        .getValues();
      for (var si = 0; si < scData.length; si++) {
        var scSku = String(scData[si][STOCKCHECK_COLS.SKU - 1] || "").trim();
        var scQty = Number(scData[si][STOCKCHECK_COLS.QTY - 1]) || 0;
        if (scSku && scQty > 0) {
          var k = normalizeWorkflowKey_(scSku);
          queuedBySku[k] = (queuedBySku[k] || 0) + scQty;
        }
      }
    }
  } catch (qbErr) {
    Logger.log("queuedBySku build error: " + qbErr);
  }

  function isEligibleForStockcheckStatus_(statusValue) {
    var s = normalizeWorkflowKey_(statusValue);
    return (
      s === "pending" ||
      s === "ready to ship" ||
      s === "faulty" ||
      s === "not found" ||
      s === "stock run out"
    );
  }

  var eligibleByKey = {};
  var eligibleRows = [];
  var allOrderRowsToMarkReady = [];
  for (var i = 0; i < allBundle.data.length; i++) {
    var row = allBundle.data[i];
    var orderStatusRaw = String(row[orderStatusCol - 1] || "").trim();
    var orderStatus = normalizeWorkflowKey_(orderStatusRaw);
    if (!isEligibleForStockcheckStatus_(orderStatus)) continue;

    var staffNotes = staffNotesCol ? String(row[staffNotesCol - 1] || "") : "";
    if (staffNotes.toLowerCase().indexOf("halted:") !== -1) continue;

    var portalOrderId = String(row[portalCol - 1] || "").trim();
    var internalOrderNo = String(row[internalCol - 1] || "").trim();
    var sku = String(row[skuCol - 1] || "").trim();
    var qty = Number(row[qtyCol - 1]) || 0;
    if (!portalOrderId || !internalOrderNo || !sku || qty <= 0) continue;

    var isStockException =
      orderStatus === "faulty" ||
      orderStatus === "not found" ||
      orderStatus === "stock run out";
    var stockQty = stockCol ? parseAvailableStockQty_(row[stockCol - 1]) : NaN;
    // Require available stock to be greater than or equal to order quantity
    // considering quantities already queued in stockcheckRTS. This prevents
    // queuing multiple pending rows that together exceed available stock.
    var skuNorm = normalizeWorkflowKey_(sku);
    var alreadyQueued = queuedBySku[skuNorm] || 0;
    if (!isStockException && !(stockQty >= qty + alreadyQueued)) continue;

    // Reserve the qty in the local queued map so subsequent rows in this
    // scan account for it when making decisions.
    queuedBySku[skuNorm] = alreadyQueued + qty;

    var internalSkuKey = buildInternalSkuKey_(internalOrderNo, sku);
    if (
      activeCreateLabelKeys[internalSkuKey] ||
      activeDispatchedKeys[internalSkuKey]
    )
      continue;

    var queueKey = buildQueueKey_(portalOrderId, internalOrderNo, sku);
    if (eligibleByKey[queueKey]) continue;

    var location = locationByPortalSku[normalizeWorkflowKey_(sku)] || "";
    var salesChannelVal = salesChannelCol ? row[salesChannelCol - 1] : "";
    var countryVal = countryCol ? row[countryCol - 1] : "";
    eligibleByKey[queueKey] = true;
    if (orderStatus === "pending") {
      allOrderRowsToMarkReady.push(allBundle.dataStartRow + i);
    }

    // Sanitize values before queuing to ensure setValues() never receives
    // unsupported object types (e.g. CellImage objects). Prefer IMAGE() formula
    // when an image URL is available.
    var rawStaffNotes = staffNotesCol ? row[staffNotesCol - 1] : "";
    var rawImageVal = imageCol ? row[imageCol - 1] : "";
    var rawImageUrlVal = imageUrlCol ? String(row[imageUrlCol - 1] || "") : "";

    function _scrub_(v) {
      if (v === null || v === undefined) return "";
      if (
        typeof v === "object" &&
        Object.prototype.toString.call(v) !== "[object Date]"
      )
        return "";
      return v;
    }

    var staffNotesVal = _scrub_(rawStaffNotes);
    var imageUrlVal = String(_scrub_(rawImageUrlVal || "")).trim();
    // Do not send image objects or formulas here — only send the IMAGE_URL.
    var imageValOut = "";

    eligibleRows.push({
      key: queueKey,
      values: [
        portalOrderId,
        internalOrderNo,
        staffNotesVal,
        qty,
        imageValOut,
        imageUrlVal,
        sku,
        salesChannelVal || "",
        countryVal || "",
        "",
        "",
      ],
    });
  }

  var lastRow = stockcheckSheet.getLastRow();
  var removed = 0;
  var existingKeys = {};
  var rowsToDelete = [];
  if (lastRow >= WORKFLOW_DATA_START_ROW) {
    var numRows = lastRow - WORKFLOW_DATA_START_ROW + 1;
    var current = stockcheckSheet
      .getRange(
        WORKFLOW_DATA_START_ROW,
        1,
        numRows,
        STOCKCHECK_COLS.DoCreateLabels,
      )
      .getValues();
    for (var r = 0; r < current.length; r++) {
      var cur = current[r];
      var key = buildQueueKey_(
        cur[STOCKCHECK_COLS.PORTAL_ORDER_ID - 1],
        cur[STOCKCHECK_COLS.INTERNAL_ORDERNO - 1],
        cur[STOCKCHECK_COLS.SKU - 1],
      );
      if (key === "||||") continue;
      if (eligibleByKey[key]) {
        existingKeys[key] = true;
      } else {
        rowsToDelete.push(WORKFLOW_DATA_START_ROW + r);
      }
    }
  }

  if (rowsToDelete.length > 0) {
    batchDeleteRows_(stockcheckSheet, rowsToDelete);
    removed += rowsToDelete.length;
  }

  var rowsToAppend = [];
  for (var j = 0; j < eligibleRows.length; j++) {
    if (!existingKeys[eligibleRows[j].key]) {
      rowsToAppend.push(eligibleRows[j].values);
    }
  }

  if (rowsToAppend.length > 0) {
    var appendStart = Math.max(
      stockcheckSheet.getLastRow() + 1,
      WORKFLOW_DATA_START_ROW,
    );
    stockcheckSheet
      .getRange(
        appendStart,
        1,
        rowsToAppend.length,
        STOCKCHECK_COLS.DoCreateLabels,
      )
      .setValues(rowsToAppend);
    try {
      if (typeof updateStockcheckImageFormulas_ === "function") {
        updateStockcheckImageFormulas_(
          stockcheckSheet,
          appendStart,
          appendStart + rowsToAppend.length - 1,
        );
      }
    } catch (e) {
      Logger.log("updateStockcheckImageFormulas_ error: " + e);
    }
  }

  if (allOrderRowsToMarkReady.length > 0) {
    var readyBg = getWorkflowStatusBackground_("Ready to Ship");
    for (var m = 0; m < allOrderRowsToMarkReady.length; m++) {
      var absRow = allOrderRowsToMarkReady[m];
      allBundle.sheet
        .getRange(absRow, orderStatusCol)
        .setValue("Ready to Ship");
      if (readyBg) {
        allBundle.sheet
          .getRange(absRow, 1, 1, allBundle.headers.length)
          .setBackground(readyBg);
      }
    }
  }

  refreshPickListSheetIfAvailable_(true);

  if (!suppressUi && rowsToAppend.length === 0 && removed === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "stockcheckRTS is already up to date.",
      "Queue Sync",
      5,
    );
  }

  return {
    added: rowsToAppend.length,
    removed: removed,
    eligible: eligibleRows.length,
  };
}

function refreshPickListSheetIfAvailable_(suppressUi) {
  if (typeof refreshPickListFromStockcheckRTS_ !== "function") return;
  try {
    refreshPickListFromStockcheckRTS_(suppressUi);
  } catch (pickErr) {
    Logger.log("refreshPickListFromStockcheckRTS_ error: " + pickErr);
  }
}

/**
 * Ensure the All Orders sheet has a `PENDING_RESERVE` header and return its col index.
 * Creates the column at the end of the header row if missing.
 */
function ensurePendingReserveHeaderCol_(bundle) {
  if (!bundle || !bundle.sheet) return 0;
  var col = getHeaderColFromMap_(bundle.headerMap, [
    "PENDING_RESERVE",
    "Pending Reserve",
    "Pending_Reserve",
    "PENDINGRESERVE",
  ]);
  if (col > 0) return col;

  try {
    var sheet = bundle.sheet;
    var headerRow = bundle.headerRow;
    var lastCol = sheet.getLastColumn();
    sheet.insertColumnsAfter(lastCol, 1);
    sheet.getRange(headerRow, lastCol + 1).setValue("PENDING_RESERVE");
    // Refresh bundle header map
    var headers = sheet
      .getRange(headerRow, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    bundle.headers = headers;
    bundle.headerMap = getWorkflowHeaderMap_(headers);
    return getHeaderColFromMap_(bundle.headerMap, [
      "PENDING_RESERVE",
      "Pending Reserve",
      "Pending_Reserve",
      "PENDINGRESERVE",
    ]);
  } catch (e) {
    Logger.log("ensurePendingReserveHeaderCol_ error: " + e);
    return 0;
  }
}

/**
 * Calculates per-row pending reserve values for All Orders.
 * Algorithm (per SKU): iterate rows top-to-bottom; for each reserve-active row,
 *  - if available >= qty: write pendingReserve = available_before and decrease available by qty
 *  - else: write pendingReserve = available_before and mark ORDER_STATUS = 'Low Stock' (do not consume stock)
 * This is a manual, on-click operation and does not mutate the Inventory sheet.
 */
function updatePendingReservesForAllOrders_(optionalBundle) {
  var bundle = optionalBundle || getAllOrdersBundle_();
  if (!bundle) return { updated: 0 };

  var hm = bundle.headerMap;
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  var qtyCol = getHeaderColFromMap_(hm, ["QTY", "Qty", "Quantity"]);
  var orderStatusCol = getHeaderColFromMap_(hm, [
    "ORDER_STATUS",
    "Order Status",
  ]);
  var stockCol = getHeaderColFromMap_(hm, [
    "stock",
    "stock_availibilty",
    "stock_availability",
    "stockavailability",
  ]);
  if (!skuCol || !qtyCol || !orderStatusCol) {
    throw new Error(
      "All Orders header mapping incomplete for reserve computation.",
    );
  }

  var pendingCol = ensurePendingReserveHeaderCol_(bundle);
  if (!pendingCol) throw new Error("Could not ensure PENDING_RESERVE header.");

  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();

  // Build normalized portal->vendor map for robust lookup
  var portalNorm = {};
  for (var p in portalMap) {
    if (!Object.prototype.hasOwnProperty.call(portalMap, p)) continue;
    portalNorm[normalizeWorkflowKey_(p)] = portalMap[p];
  }

  // Current available per vendor (normalized key) — clone inventory stock
  var currentAvail = {};
  for (var v in invMap) {
    if (!Object.prototype.hasOwnProperty.call(invMap, v)) continue;
    currentAvail[normalizeWorkflowKey_(v)] = Number(invMap[v].Stock) || 0;
  }

  var outPending = [];
  var outStatus = [];
  var rows = bundle.data || [];

  function isReserveActiveStatus_(statusValue) {
    var s = normalizeWorkflowKey_(statusValue);
    return (
      s === "pending" ||
      s === "low stock" ||
      s === "lowstock" ||
      s === "stock run out" ||
      s === "not found" ||
      s === "faulty" ||
      s === "ready to ship" ||
      s === "readytoship"
    );
  }

  function isRecoverableToPendingStatus_(statusValue) {
    var s = normalizeWorkflowKey_(statusValue);
    return (
      s === "low stock" ||
      s === "lowstock" ||
      s === "stock run out" ||
      s === "not found" ||
      s === "faulty"
    );
  }

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var statusRaw = row[orderStatusCol - 1];
    var statusNorm = normalizeWorkflowKey_(statusRaw);

    // Only recalc rows that can still feed stockcheckRTS; preserve completed flow states.
    if (!isReserveActiveStatus_(statusRaw)) {
      outPending.push([""]);
      outStatus.push([row[orderStatusCol - 1]]);
      continue;
    }

    var sku = String(row[skuCol - 1] || "").trim();
    var qty = Number(row[qtyCol - 1]) || 0;
    if (!sku || qty <= 0) {
      outPending.push([""]);
      outStatus.push([row[orderStatusCol - 1]]);
      continue;
    }

    var v = portalNorm[normalizeWorkflowKey_(sku)];
    var vNorm = normalizeWorkflowKey_(v || "");
    var reserveKey = vNorm || normalizeWorkflowKey_(sku);
    var stockTextAvail = stockCol
      ? parseAvailableStockQty_(row[stockCol - 1])
      : NaN;
    var hasAvailability =
      (vNorm && typeof currentAvail[vNorm] !== "undefined") ||
      !!invMap[v] ||
      !isNaN(stockTextAvail);
    if (!hasAvailability) {
      outPending.push([""]);
      outStatus.push([row[orderStatusCol - 1]]);
      continue;
    }

    var availBefore =
      typeof currentAvail[reserveKey] !== "undefined"
        ? currentAvail[reserveKey]
        : invMap[v]
          ? Number(invMap[v].Stock) || 0
          : !isNaN(stockTextAvail)
            ? stockTextAvail
            : 0;

    // If enough available, reserve (decrement in-memory); otherwise mark low stock and skip
    if (availBefore >= qty) {
      outPending.push([availBefore]);
      currentAvail[reserveKey] = availBefore - qty;
      if (isRecoverableToPendingStatus_(statusRaw)) {
        outStatus.push(["Pending"]);
      } else {
        outStatus.push([row[orderStatusCol - 1]]);
      }
    } else {
      outPending.push([availBefore]);
      if (
        statusNorm === "pending" ||
        statusNorm === "low stock" ||
        statusNorm === "lowstock" ||
        statusNorm === "stock run out" ||
        statusNorm === "ready to ship" ||
        statusNorm === "readytoship"
      ) {
        outStatus.push(["Low Stock"]);
      } else {
        outStatus.push([row[orderStatusCol - 1]]);
      }
    }
  }

  // Batch write pending reserves and updated statuses
  var sheet = bundle.sheet;
  var start = bundle.dataStartRow;
  if (outPending.length > 0) {
    try {
      sheet
        .getRange(start, pendingCol, outPending.length, 1)
        .setValues(outPending);
    } catch (e) {
      Logger.log(
        "updatePendingReservesForAllOrders_ pending write error: " + e,
      );
    }
  }
  if (outStatus.length > 0) {
    try {
      sheet
        .getRange(start, orderStatusCol, outStatus.length, 1)
        .setValues(outStatus);
    } catch (e2) {
      Logger.log(
        "updatePendingReservesForAllOrders_ status write error: " + e2,
      );
    }
  }

  return { updated: outPending.length };
}

/**
 * Button wrapper: compute pending reserves for All Orders.
 */
function updateReserveButton() {
  try {
    var res = updatePendingReservesForAllOrders_();
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Pending reserves updated: " + (res && res.updated ? res.updated : 0),
      "Update Reserve",
      6,
    );
    return res;
  } catch (err) {
    Logger.log("updateReserveButton error: " + err);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Reserve update failed: " + String(err),
      "Update Reserve",
      6,
    );
    throw err;
  }
}

/**
 * Sync to stockcheckRTS using the computed PENDING_RESERVE field.
 * If requireFullQty is true (default) a row is eligible only when both
 *   - inventory has enough (considering already queued in stockcheckRTS)
 *   - PENDING_RESERVE for the row is >= qty
 */
function syncPendingOrdersToStockcheckRTSUsingReserve_(
  suppressUi,
  requireFullQty,
) {
  requireFullQty =
    typeof requireFullQty === "undefined" ? true : !!requireFullQty;
  var ss = getWorkflowSpreadsheet_();
  var stockcheckSheet = ss.getSheetByName(SHEET_NAMES.STOCKCHECK_RTS);
  if (!stockcheckSheet) throw new Error("stockcheckRTS sheet not found.");

  var allBundle = getAllOrdersBundle_();
  if (!allBundle) return { added: 0, removed: 0, eligible: 0 };

  var hm = allBundle.headerMap;
  var portalCol = getHeaderColFromMap_(hm, [
    "PORTAL_ORDER_ID",
    "PortalOrderID",
    "Portal Order ID",
  ]);
  var orderStatusCol = getHeaderColFromMap_(hm, [
    "ORDER_STATUS",
    "Order Status",
  ]);
  var internalCol = getHeaderColFromMap_(hm, [
    "INTERNAL_ORDERNO",
    "Internal OrderNo",
    "InternalOrderNo",
  ]);
  var qtyCol = getHeaderColFromMap_(hm, ["QTY", "Qty", "Quantity"]);
  var imageCol = getHeaderColFromMap_(hm, ["IMAGE", "Image"]);
  var imageUrlCol = getHeaderColFromMap_(hm, [
    "IMAGE_URL",
    "Image URL",
    "Image Url",
  ]);
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  var pendingCol = getHeaderColFromMap_(hm, [
    "PENDING_RESERVE",
    "Pending Reserve",
    "Pending_Reserve",
  ]);
  var salesChannelCol = getHeaderColFromMap_(hm, [
    "SALES_CHANNEL",
    "Sales Channel",
    "SalesChannel",
  ]);
  var countryCol = getHeaderColFromMap_(hm, ["COUNTRY", "Country"]);

  if (!portalCol || !orderStatusCol || !internalCol || !qtyCol || !skuCol) {
    throw new Error(
      "All Orders header mapping is incomplete. Please verify headers.",
    );
  }
  if (!pendingCol)
    throw new Error(
      "PENDING_RESERVE column not found. Run Update Reserve first.",
    );

  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();

  // Build queuedBySku from existing stockcheckRTS rows (same as main sync)
  var queuedBySku = {};
  try {
    var scLastRow = stockcheckSheet.getLastRow();
    if (scLastRow >= WORKFLOW_DATA_START_ROW) {
      var scNumRows = scLastRow - WORKFLOW_DATA_START_ROW + 1;
      var scData = stockcheckSheet
        .getRange(
          WORKFLOW_DATA_START_ROW,
          1,
          scNumRows,
          STOCKCHECK_COLS.DoCreateLabels,
        )
        .getValues();
      for (var si = 0; si < scData.length; si++) {
        var scSku = String(scData[si][STOCKCHECK_COLS.SKU - 1] || "").trim();
        var scQty = Number(scData[si][STOCKCHECK_COLS.QTY - 1]) || 0;
        if (scSku && scQty > 0) {
          var k = normalizeWorkflowKey_(scSku);
          queuedBySku[k] = (queuedBySku[k] || 0) + scQty;
        }
      }
    }
  } catch (qbErr) {
    Logger.log("queuedBySku build error (reserve-sync): " + qbErr);
  }

  var eligibleByKey = {};
  var eligibleRows = [];
  var allOrderRowsToMarkReady = [];

  for (var i = 0; i < allBundle.data.length; i++) {
    var row = allBundle.data[i];
    var orderStatusRaw = String(row[orderStatusCol - 1] || "").trim();
    var orderStatus = normalizeWorkflowKey_(orderStatusRaw);
    if (orderStatus !== "pending") continue;

    var portalOrderId = String(row[portalCol - 1] || "").trim();
    var internalOrderNo = String(row[internalCol - 1] || "").trim();
    var sku = String(row[skuCol - 1] || "").trim();
    var qty = Number(row[qtyCol - 1]) || 0;
    if (!portalOrderId || !internalOrderNo || !sku || qty <= 0) continue;

    var pendingVal = Number(row[pendingCol - 1]) || 0;
    var skuNorm = normalizeWorkflowKey_(sku);
    var vendorSku =
      portalMap[sku] || portalMap[normalizeWorkflowKey_(sku)] || null;
    var invStock =
      vendorSku && invMap[vendorSku] ? Number(invMap[vendorSku].Stock) : NaN;
    var alreadyQueued = queuedBySku[skuNorm] || 0;

    // Require inventory availability (considering already queued) and pending reserve
    if (isNaN(invStock)) continue;
    if (requireFullQty) {
      if (!(invStock >= qty + alreadyQueued && pendingVal >= qty)) continue;
    } else {
      if (!(pendingVal > 0 && invStock > alreadyQueued)) continue;
    }

    // Reserve locally for subsequent rows
    queuedBySku[skuNorm] = alreadyQueued + qty;

    var internalSkuKey = buildInternalSkuKey_(internalOrderNo, sku);
    var queueKey = buildQueueKey_(portalOrderId, internalOrderNo, sku);
    if (eligibleByKey[queueKey]) continue;

    // Sanitize values similar to main sync function
    function _scrub_(v) {
      if (v === null || v === undefined) return "";
      if (
        typeof v === "object" &&
        Object.prototype.toString.call(v) !== "[object Date]"
      )
        return "";
      return v;
    }

    var imageUrlVal = String(
      _scrub_(imageUrlCol ? row[imageUrlCol - 1] : "") || "",
    ).trim();
    var imageValOut = ""; // do not send image objects or inline formulas
    var salesChannelVal = salesChannelCol ? row[salesChannelCol - 1] : "";
    var countryVal = countryCol ? row[countryCol - 1] : "";

    eligibleByKey[queueKey] = true;
    allOrderRowsToMarkReady.push(allBundle.dataStartRow + i);

    eligibleRows.push({
      key: queueKey,
      values: [
        portalOrderId,
        internalOrderNo,
        "",
        qty,
        imageValOut,
        imageUrlVal,
        sku,
        salesChannelVal || "",
        countryVal || "",
        "",
        "",
      ],
    });
  }

  // Keep existing stockcheckRTS rows in place. This button only appends newly
  // eligible Pending rows; it does not auto-clean/sync the queue.
  var lastRow = stockcheckSheet.getLastRow();
  var removed = 0;
  var existingKeys = {};
  if (lastRow >= WORKFLOW_DATA_START_ROW) {
    var numRows = lastRow - WORKFLOW_DATA_START_ROW + 1;
    var current = stockcheckSheet
      .getRange(
        WORKFLOW_DATA_START_ROW,
        1,
        numRows,
        STOCKCHECK_COLS.DoCreateLabels,
      )
      .getValues();
    for (var r = 0; r < current.length; r++) {
      var cur = current[r];
      var key = buildQueueKey_(
        cur[STOCKCHECK_COLS.PORTAL_ORDER_ID - 1],
        cur[STOCKCHECK_COLS.INTERNAL_ORDERNO - 1],
        cur[STOCKCHECK_COLS.SKU - 1],
      );
      if (key === "||||") continue;
      existingKeys[key] = true;
    }
  }

  var rowsToAppend = [];
  for (var j = 0; j < eligibleRows.length; j++) {
    if (!existingKeys[eligibleRows[j].key])
      rowsToAppend.push(eligibleRows[j].values);
  }

  if (rowsToAppend.length > 0) {
    var appendStart = Math.max(
      stockcheckSheet.getLastRow() + 1,
      WORKFLOW_DATA_START_ROW,
    );
    stockcheckSheet
      .getRange(
        appendStart,
        1,
        rowsToAppend.length,
        STOCKCHECK_COLS.DoCreateLabels,
      )
      .setValues(rowsToAppend);
    try {
      if (typeof updateStockcheckImageFormulas_ === "function") {
        updateStockcheckImageFormulas_(
          stockcheckSheet,
          appendStart,
          appendStart + rowsToAppend.length - 1,
        );
      }
    } catch (e) {
      Logger.log("updateStockcheckImageFormulas_ error: " + e);
    }
  }

  if (allOrderRowsToMarkReady.length > 0) {
    var readyBg = getWorkflowStatusBackground_("Ready to Ship");
    for (var m = 0; m < allOrderRowsToMarkReady.length; m++) {
      var absRow = allOrderRowsToMarkReady[m];
      allBundle.sheet
        .getRange(absRow, orderStatusCol)
        .setValue("Ready to Ship");
      if (readyBg) {
        allBundle.sheet
          .getRange(absRow, 1, 1, allBundle.headers.length)
          .setBackground(readyBg);
      }
    }
  }

  refreshPickListSheetIfAvailable_(true);

  if (!suppressUi && rowsToAppend.length === 0 && removed === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "stockcheckRTS is already up to date.",
      "Queue Sync (Reserve)",
      5,
    );
  }

  return {
    added: rowsToAppend.length,
    removed: removed,
    eligible: eligibleRows.length,
  };
}

function syncPendingOrdersToStockcheckRTSUsingCurrentReserve_(
  suppressUi,
  requireFullQty,
) {
  var bundle = getAllOrdersBundle_();
  if (!bundle) return { added: 0, removed: 0, eligible: 0 };

  var pendingCol = getHeaderColFromMap_(bundle.headerMap, [
    "PENDING_RESERVE",
    "Pending Reserve",
    "Pending_Reserve",
  ]);
  if (!pendingCol) {
    if (!suppressUi) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        "Run Update Reserve before sending rows to stockcheckRTS.",
        "Send to StockCheckRTS",
        6,
      );
    }
    return { added: 0, removed: 0, eligible: 0 };
  }

  return syncPendingOrdersToStockcheckRTSUsingReserve_(
    suppressUi,
    typeof requireFullQty === "undefined" ? true : !!requireFullQty,
  );
}

/**
 * Button wrapper: sync to stockcheckRTS using the last computed reserve values.
 */
function sendToStockcheckRTSButton(requireFullQty) {
  try {
    var res = syncPendingOrdersToStockcheckRTSUsingReserve_(
      false,
      typeof requireFullQty === "undefined" ? true : !!requireFullQty,
    );
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Sent to stockcheckRTS – added: " +
        res.added +
        ", removed: " +
        res.removed,
      "Send to StockCheckRTS",
      6,
    );
    return res;
  } catch (err) {
    Logger.log("sendToStockcheckRTSButton error: " + err);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Send to stockcheckRTS failed: " + String(err),
      "Send to StockCheckRTS",
      6,
    );
    throw err;
  }
}

function normalizeWorkflowChoice_(value, choices) {
  var raw = String(value == null ? "" : value).trim();
  if (!raw) return "";
  for (var i = 0; i < choices.length; i++) {
    if (raw.toLowerCase() === choices[i].toLowerCase()) return choices[i];
  }
  return raw;
}

/**
 * Deletes an array of row indices efficiently by grouping contiguous rows.
 */
function batchDeleteRows_(sheet, rowsToDelete) {
  if (!sheet || !rowsToDelete || rowsToDelete.length === 0) return;
  var uniqueRows = [];
  var seen = {};
  for (var i = 0; i < rowsToDelete.length; i++) {
    var r = Number(rowsToDelete[i]);
    if (r > 0 && !seen[r]) {
      uniqueRows.push(r);
      seen[r] = true;
    }
  }
  uniqueRows.sort(function (a, b) {
    return b - a;
  });
  if (uniqueRows.length === 0) return;

  var start = uniqueRows[0];
  var count = 1;
  for (var j = 1; j < uniqueRows.length; j++) {
    if (uniqueRows[j] === start - count) {
      count++;
    } else {
      sheet.deleteRows(start - count + 1, count);
      start = uniqueRows[j];
      count = 1;
    }
  }
  sheet.deleteRows(start - count + 1, count);
}

function ensureSheetMinColumns_(sheet, minCols) {
  if (!sheet) return;
  var maxCols = sheet.getMaxColumns();
  if (maxCols >= minCols) return;
  sheet.insertColumnsAfter(maxCols, minCols - maxCols);
}

function ensureWorkflowSheetStructures_() {
  var ss = getWorkflowSpreadsheet_();

  var createSheet = ss.getSheetByName(SHEET_NAMES.CREATE_LABEL);
  if (createSheet) {
    ensureSheetMinColumns_(createSheet, CREATE_LABEL_COLS.TRACKING_CODE);
    createSheet
      .getRange(1, 1, 1, CREATE_LABEL_COLS.TRACKING_CODE)
      .setValues([
        [
          "INTERNAL_ORDERNO",
          "SKU",
          "IMAGE",
          "IMAGE_URL",
          "FULLNAME",
          "ADDRESSLINE1",
          "ADDRESSLINE2",
          "CITY",
          "STATE",
          "PINCODE",
          "COUNTRY",
          "PHONE",
          "LabelCreated",
          "COURIER_NAME",
          "DELIVERY_PICKUP_DATE",
          "SALES_CHANNEL",
          "TRACKING_CODE",
        ],
      ]);
  }

  var dispatchedSheet = ss.getSheetByName(SHEET_NAMES.DISPATCHED);
  if (dispatchedSheet) {
    ensureSheetMinColumns_(dispatchedSheet, DISPATCHED_COLS.PACKET_COUNT);
    dispatchedSheet
      .getRange(1, 1, 1, DISPATCHED_COLS.PACKET_COUNT)
      .setValues([
        [
          "INTERNAL_ORDERNO",
          "SKU",
          "IMAGE",
          "IMAGE_URL",
          "FULLNAME",
          "ADDRESSLINE1",
          "ADDRESSLINE2",
          "CITY",
          "STATE",
          "PINCODE",
          "COUNTRY",
          "PHONE",
          "DISPACHED_STATUS",
          "PICKUP_DATE",
          "COURIER_NAME",
          "TRACKING_CODE",
          "PACKET_COUNT",
        ],
      ]);
  }

  getOrCreateDispatchHistorySheet_();
}

function appendHaltReason_(currentValue, reason) {
  var trimmed = String(currentValue == null ? "" : currentValue).trim();
  var stamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || "Asia/Kolkata",
    "dd-MMM-yyyy HH:mm",
  );
  var entry = "Halted: " + reason + " (" + stamp + ")";
  if (!trimmed) return entry;
  if (trimmed.toLowerCase().indexOf(("halted: " + reason).toLowerCase()) !== -1)
    return trimmed;
  return trimmed + " | " + entry;
}

function updateAllOrdersForHalt_(internalOrderNo, sku, reason, optionalBundle) {
  var bundle = optionalBundle || getAllOrdersBundle_();
  if (!bundle) return 0;

  var hm = bundle.headerMap;
  var internalCol = getHeaderColFromMap_(hm, [
    "INTERNAL_ORDERNO",
    "Internal OrderNo",
    "InternalOrderNo",
  ]);
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  var orderStatusCol = getHeaderColFromMap_(hm, [
    "ORDER_STATUS",
    "Order Status",
  ]);
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
    notesCell
      .setBackground("#fff3cd")
      .setFontColor("#b71c1c")
      .setFontWeight("bold");
    touched++;
  }

  return touched;
}

function removeRowsFromQueueByInternalSku_(
  sheet,
  internalCol,
  skuCol,
  internalOrderNo,
  sku,
) {
  if (!sheet) return 0;
  var targetKey = buildInternalSkuKey_(internalOrderNo, sku);
  var lastRow = sheet.getLastRow();
  if (lastRow < WORKFLOW_DATA_START_ROW) return 0;

  var rows = sheet
    .getRange(
      WORKFLOW_DATA_START_ROW,
      1,
      lastRow - WORKFLOW_DATA_START_ROW + 1,
      Math.max(internalCol, skuCol),
    )
    .getValues();
  var rowsToDelete = [];
  for (var i = 0; i < rows.length; i++) {
    var key = buildInternalSkuKey_(
      rows[i][internalCol - 1],
      rows[i][skuCol - 1],
    );
    if (key === targetKey) rowsToDelete.push(WORKFLOW_DATA_START_ROW + i);
  }
  batchDeleteRows_(sheet, rowsToDelete);
  return rowsToDelete.length;
}

function findAllOrderAddressRow_(internalOrderNo, sku, optionalBundle) {
  var bundle = optionalBundle || getAllOrdersBundle_();
  if (!bundle) return null;
  var hm = bundle.headerMap;
  var internalCol = getHeaderColFromMap_(hm, [
    "INTERNAL_ORDERNO",
    "Internal OrderNo",
    "InternalOrderNo",
  ]);
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  if (!internalCol || !skuCol) return null;

  var targetInternal = normalizeWorkflowKey_(internalOrderNo);
  var targetSku = normalizeWorkflowKey_(sku);

  for (var i = 0; i < bundle.data.length; i++) {
    var r = bundle.data[i];
    if (normalizeWorkflowKey_(r[internalCol - 1]) !== targetInternal) continue;
    if (normalizeWorkflowKey_(r[skuCol - 1]) !== targetSku) continue;
    return {
      row: r,
      headerMap: hm,
    };
  }

  for (var j = 0; j < bundle.data.length; j++) {
    var fallback = bundle.data[j];
    if (normalizeWorkflowKey_(fallback[internalCol - 1]) === targetInternal) {
      return {
        row: fallback,
        headerMap: hm,
      };
    }
  }

  return null;
}

/**
 * Read a specific field value from All Orders by matching Internal OrderNo (and SKU where possible).
 * Returns empty string when not found. This reads the cell directly to ensure latest value.
 */
function getAllOrdersFieldValueByInternal_(internalOrderNo, sku, aliases) {
  var bundle = getAllOrdersBundle_();
  if (!bundle || !bundle.sheet) return "";
  var hm = bundle.headerMap;
  var col = getHeaderColFromMap_(hm, aliases);
  if (!col) return "";

  var internalCol = getHeaderColFromMap_(hm, [
    "INTERNAL_ORDERNO",
    "Internal OrderNo",
    "InternalOrderNo",
  ]);
  var skuCol = getHeaderColFromMap_(hm, ["SKU", "PortalSKU"]);
  if (!internalCol) return "";

  var targetInternal = normalizeWorkflowKey_(internalOrderNo);
  var targetSku = normalizeWorkflowKey_(sku || "");

  // Prefer exact match on internal+sku
  for (var i = 0; i < bundle.data.length; i++) {
    var r = bundle.data[i];
    if (normalizeWorkflowKey_(r[internalCol - 1]) !== targetInternal) continue;
    if (
      targetSku &&
      skuCol &&
      normalizeWorkflowKey_(r[skuCol - 1]) !== targetSku
    )
      continue;
    var absRow = bundle.dataStartRow + i;
    try {
      return bundle.sheet.getRange(absRow, col).getValue();
    } catch (e) {
      return r[col - 1] || "";
    }
  }

  // Fallback: match on internal only
  for (var j = 0; j < bundle.data.length; j++) {
    var rr = bundle.data[j];
    if (normalizeWorkflowKey_(rr[internalCol - 1]) !== targetInternal) continue;
    var abs = bundle.dataStartRow + j;
    try {
      return bundle.sheet.getRange(abs, col).getValue();
    } catch (e2) {
      return rr[col - 1] || "";
    }
  }

  return "";
}

function getWorkflowMapValueByNormalizedKey_(map, key) {
  if (!map || !key) return "";
  if (Object.prototype.hasOwnProperty.call(map, key)) return map[key];

  var target = normalizeWorkflowKey_(key);
  for (var k in map) {
    if (!Object.prototype.hasOwnProperty.call(map, k)) continue;
    if (normalizeWorkflowKey_(k) === target) return map[k];
  }
  return "";
}

function buildWorkflowInventoryLogRow_(allOrderRow, beforeStock, afterStock) {
  var source = allOrderRow || [];
  var out = new Array(47).fill("");
  var copyCount = Math.min(source.length, 42);
  for (var i = 0; i < copyCount; i++) {
    var val = source[i];
    if (
      val !== null &&
      typeof val === "object" &&
      Object.prototype.toString.call(val) !== "[object Date]"
    ) {
      val = "";
    }
    out[i] = val;
  }

  if (source.length >= 29) out[42] = source[28];
  if (source.length >= 41) out[43] = source[40];

  out[44] = beforeStock;
  out[45] = afterStock;
  out[46] = new Date();
  return out;
}

function appendWorkflowInventoryLogRows_(rowsToLog) {
  if (!rowsToLog || rowsToLog.length === 0) return;

  var logSs = getCatalogSpreadsheet_();
  var logSheet =
    logSs.getSheetByName(INVENTORY_HISTORY_LOG_SHEET_NAME) ||
    logSs.insertSheet(INVENTORY_HISTORY_LOG_SHEET_NAME);
  var startRow = Math.max(logSheet.getLastRow() + 1, 2);
  logSheet
    .getRange(startRow, 1, rowsToLog.length, rowsToLog[0].length)
    .setValues(rowsToLog);
}

function applyCreateLabelInventoryChanges_(items) {
  if (!items || items.length === 0)
    return { updated: 0, logged: 0, skipped: 0 };

  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var invSheet = getInventorySheet_();
  if (!invSheet) return { updated: 0, logged: 0, skipped: items.length };

  var workingStock = {};
  var stockUpdates = {};
  var logRows = [];
  var skipped = 0;

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var sku = String(item.sku || "").trim();
    var qty = Number(item.qty) || 0;
    if (!sku || qty <= 0) {
      skipped++;
      continue;
    }

    var vendorSku = getWorkflowMapValueByNormalizedKey_(portalMap, sku);
    var inv = vendorSku
      ? getWorkflowMapValueByNormalizedKey_(invMap, vendorSku)
      : null;
    if (!vendorSku || !inv) {
      logRows.push(buildWorkflowInventoryLogRow_(item.allOrderRow, "", ""));
      skipped++;
      continue;
    }

    var invKey = normalizeWorkflowKey_(vendorSku);
    var beforeStock = Object.prototype.hasOwnProperty.call(workingStock, invKey)
      ? workingStock[invKey]
      : Number(inv.Stock) || 0;
    var afterStock = Math.max(0, beforeStock - qty);
    workingStock[invKey] = afterStock;
    stockUpdates[invKey] = {
      row: inv.invRow,
      after: afterStock,
    };
    logRows.push(
      buildWorkflowInventoryLogRow_(item.allOrderRow, beforeStock, afterStock),
    );
  }

  var updated = 0;
  for (var key in stockUpdates) {
    if (!Object.prototype.hasOwnProperty.call(stockUpdates, key)) continue;
    var update = stockUpdates[key];
    invSheet.getRange(update.row, INVENTORY_COLS.Stock).setValue(update.after);
    updated++;
  }

  appendWorkflowInventoryLogRows_(logRows);
  try {
    refreshAllOrdersStockAvailability_();
  } catch (refreshErr) {
    Logger.log("refreshAllOrdersStockAvailability_ error: " + refreshErr);
  }
  return { updated: updated, logged: logRows.length, skipped: skipped };
}

function moveStockcheckRowToCreateLabel_(
  stockcheckSheet,
  rowNumber,
  rowValues,
) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var createSheet = ss.getSheetByName(SHEET_NAMES.CREATE_LABEL);
  if (!createSheet) throw new Error("CreateLabels sheet not found.");
  ensureSheetMinColumns_(createSheet, CREATE_LABEL_COLS.TRACKING_CODE);

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

  var key = buildInternalSkuKey_(internalOrderNo, sku);
  var countryVal =
    rowValues[STOCKCHECK_COLS.COUNTRY - 1] ||
    fromAllOrders(["COUNTRY", "Country"]);
  var salesChannelVal =
    rowValues[STOCKCHECK_COLS.SALES_CHANNEL - 1] ||
    getAllOrdersFieldValueByInternal_(internalOrderNo, sku, [
      "SALES_CHANNEL",
      "Sales Channel",
    ]);
  var deliveryPickupDateVal = getAllOrdersFieldValueByInternal_(
    internalOrderNo,
    sku,
    ["DELIVERY_PICKUP_DATE", "DeliveryPickupDate", "Delivery Pickup Date"],
  );
  var courierName = resolveCourierNameForOrder_(countryVal, salesChannelVal);

  var payload = [
    [
      internalOrderNo,
      sku,
      // Do not send the image formula — only send the URL; sheet will convert it
      "",
      String(
        fromAllOrders(["IMAGE_URL", "Image URL", "ImageUrl"]) || "",
      ).trim(),
      fromAllOrders(["FULLNAME", "Fullname", "Full Name"]),
      fromAllOrders(["ADDRESSLINE1", "AddressLine1"]),
      fromAllOrders(["ADDRESSLINE2", "AddressLine2"]),
      fromAllOrders(["CITY", "City"]),
      fromAllOrders(["STATE", "State"]),
      fromAllOrders(["PINCODE", "Pincode"]),
      countryVal,
      fromAllOrders(["PHONE", "Phone"]),
      "No",
      courierName,
      deliveryPickupDateVal || "",
      salesChannelVal || "",
      "",
    ],
  ];

  var existing = getActiveInternalSkuKeysFromSheet_(
    createSheet,
    CREATE_LABEL_COLS.INTERNAL_ORDERNO,
    CREATE_LABEL_COLS.SKU,
    WORKFLOW_DATA_START_ROW,
  );

  if (!existing[key]) {
    var appendRow = Math.max(
      createSheet.getLastRow() + 1,
      WORKFLOW_DATA_START_ROW,
    );
    createSheet
      .getRange(appendRow, 1, 1, CREATE_LABEL_COLS.TRACKING_CODE)
      .setValues(payload);
    // Ensure IMAGE formulas are generated from IMAGE_URL for the newly appended row
    try {
      updateCreateLabelsImageFormulas_(createSheet, appendRow, appendRow);
    } catch (imgErr) {
      Logger.log("updateCreateLabelsImageFormulas_ error: " + imgErr);
    }
    applyCreateLabelInventoryChanges_([
      {
        sku: sku,
        qty: Number(rowValues[STOCKCHECK_COLS.QTY - 1]) || 0,
        allOrderRow: sourceRow,
      },
    ]);
  }

  updateAllOrdersDispatchDetailsByInternal_(internalOrderNo, sku, {
    orderStatus: "Create Label",
  });

  stockcheckSheet.deleteRow(rowNumber);
}

function haltStockcheckRow_(stockcheckSheet, rowNumber, rowValues, reason) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var internalOrderNo = rowValues[STOCKCHECK_COLS.INTERNAL_ORDERNO - 1];
  var sku = rowValues[STOCKCHECK_COLS.SKU - 1];

  updateAllOrdersForHalt_(internalOrderNo, sku, reason);
  stockcheckSheet.deleteRow(rowNumber);

  removeRowsFromQueueByInternalSku_(
    ss.getSheetByName(SHEET_NAMES.CREATE_LABEL),
    CREATE_LABEL_COLS.INTERNAL_ORDERNO,
    CREATE_LABEL_COLS.SKU,
    internalOrderNo,
    sku,
  );
  removeRowsFromQueueByInternalSku_(
    ss.getSheetByName(SHEET_NAMES.PICKLIST),
    2,
    7,
    internalOrderNo,
    sku,
  );
  removeRowsFromQueueByInternalSku_(
    ss.getSheetByName(SHEET_NAMES.DISPATCHED),
    DISPATCHED_COLS.INTERNAL_ORDERNO,
    DISPATCHED_COLS.SKU,
    internalOrderNo,
    sku,
  );
}

function handleStockcheckWorkflowEdit_(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.STOCKCHECK_RTS) return;

  var startRow = Math.max(e.range.getRow(), WORKFLOW_DATA_START_ROW);
  var endRow = e.range.getLastRow();
  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();

  if (endRow < WORKFLOW_DATA_START_ROW) return
