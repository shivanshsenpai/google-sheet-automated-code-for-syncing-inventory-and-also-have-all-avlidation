/* =============================================================================
 * DesignSheet Tools
 *
 * Works on the "DesignSheet" sheet with structure:
 * A: SKU | B: ITEM_NAME | C: IMAGE | D: Design_Category | E: DesignCode | F: IMAGE_URL
 *
 * Features:
 * 1. Extract design codes from SKU  (pattern -D / _D, e.g. GBB-D3-CZ-2.2 → D3)
 * 2. Detect Design_Category from ITEM_NAME using Configuration sheet keywords
 * 3. Load images from IMAGE_URL into IMAGE column via =IMAGE() formulas
 * 4. Fix broken URLs (Amazon, Flipkart known patterns)
 * 5. Size-aware duplicate detection (same code+category but diff size = skip)
 * 6. Sort sheet by Design_Category (alpha) then DesignCode within each category
 * 7. Remove rows marked as "skip"
 * ============================================================================= */

var DESIGN_SHEET_NAME = 'DesignSheet';
var CONFIG_SHEET_NAME = 'Configuration';

var DS_COLS = {
  SKU:             1,  // A
  ITEM_NAME:       2,  // B
  IMAGE:           3,  // C
  Design_Category: 4,  // D
  DesignCode:      5,  // E
  IMAGE_URL:       6   // F
};

// Error column is added dynamically as the 7th column (G)
var DS_ERROR_COL = 7;
var DS_HEADER_ROW = 1;
var DS_DATA_START = 2;

/* ================================ URL FIXING =============================== */

/**
 * Fixes known broken URL patterns.
 * Add more replacements below as new broken patterns are discovered.
 */
function fixUrl_(url) {
  if (!url) return '';
  var fixed = String(url).trim();
  if (!fixed) return '';

  // ── Amazon domain fix ──
  fixed = fixed.replace(/mediamazon\.com/gi, 'media-amazon.com');

  // ── Flipkart known category-slug fixes ──
  fixed = fixed.replace(/braceletrmlet/gi, 'bracelet-armlet');

  return fixed;
}

/**
 * Standalone button: Fix all URLs in the IMAGE_URL column of DesignSheet.
 * Also refreshes IMAGE formulas after fixing.
 */
function fixUrl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DESIGN_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "' + DESIGN_SHEET_NAME + '" not found.');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < DS_DATA_START) {
    SpreadsheetApp.getUi().alert('No data rows found.');
    return;
  }

  var numRows = lastRow - DS_DATA_START + 1;
  var urls = sheet.getRange(DS_DATA_START, DS_COLS.IMAGE_URL, numRows, 1).getValues();
  var fixedUrls = [];
  var fixCount = 0;

  for (var i = 0; i < numRows; i++) {
    var original = String(urls[i][0]).trim();
    var fixed = fixUrl_(original);
    if (fixed !== original && original) fixCount++;
    fixedUrls.push([fixed]);
  }

  // Write fixed URLs back
  sheet.getRange(DS_DATA_START, DS_COLS.IMAGE_URL, numRows, 1).setValues(fixedUrls);
  SpreadsheetApp.flush();

  // Refresh images
  ss.toast('Refreshing images...', 'Design Tools', -1);
  updateDesignSheetImages_(sheet, DS_DATA_START, lastRow);
  ss.toast('Images refreshed!', 'Design Tools', 3);

  SpreadsheetApp.getUi().alert(
    'Done!\n\nFixed ' + fixCount + ' URL(s) and refreshed images.'
  );
}

/* ============================== CONFIGURATION ============================== */

/**
 * Reads the Configuration sheet and builds a design-category lookup.
 */
function loadDesignCategoryConfig_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    throw new Error(
      'Configuration sheet not found. Expected a sheet named "' +
      CONFIG_SHEET_NAME + '".'
    );
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('Configuration sheet has no data rows.');
  }

  var headers = data[0].map(function(h) { return h.toString().trim(); });
  var colCat = headers.indexOf('Design_Category');
  var colKw  = headers.indexOf('Design_Keyword');

  if (colCat === -1 || colKw === -1) {
    throw new Error(
      'Configuration sheet must have "Design_Category" and "Design_Keyword" columns. ' +
      'Found headers: ' + headers.join(', ')
    );
  }

  var configs = [];

  for (var i = 1; i < data.length; i++) {
    var cat = data[i][colCat];
    var kwRaw = data[i][colKw];
    if (!cat || !String(cat).trim() || !kwRaw || !String(kwRaw).trim()) continue;

    cat = String(cat).trim();
    var keywords = String(kwRaw).split('|').map(function(k) {
      return k.trim();
    }).filter(Boolean);

    // Dynamically add the category name itself as a fallback keyword
    var cleanCat = cat.replace(/\([^)]*\)/g, '').trim();
    if (cleanCat) {
      if (keywords.indexOf(cleanCat) === -1) {
        keywords.push(cleanCat);
      }
      // Add plural form
      if (cleanCat.slice(-1).toLowerCase() !== 's') {
        var pluralCat = cleanCat + 's';
        if (keywords.indexOf(pluralCat) === -1) {
          keywords.push(pluralCat);
        }
      }
    }

    if (keywords.length === 0) continue;

    // Sort longest first for greedy matching
    keywords.sort(function(a, b) { return b.length - a.length; });

    // Build regex: each keyword gets word boundaries and flexible whitespace
    var parts = keywords.map(function(kw) {
      var escaped = kw.replace(/[/\-\\^$*+?.()|[\]{}]/g, '\\$&');
      escaped = escaped.replace(/[\s\-_]+/g, '[\\s\\-_]*');
      var startB = /^\w/.test(kw) ? '\\b' : '';
      var endB   = /\w$/.test(kw) ? '\\b' : '';
      return startB + escaped + endB;
    });

    var regex = new RegExp(parts.join('|'), 'i');
    configs.push({ category: cat, regex: regex });
  }

  // Sort configs so priority items are matched first
  var priorityOrder = [
    'Ring', 'Earring', 'Pendant', 'Necklace', 'Bracelet', 
    'Bangle', 'Anklet', 'Kada (Copper)', 'Rudraksha', 'Shaligram'
  ];

  configs.sort(function(a, b) {
    if (a.category === 'Gemstone' && b.category !== 'Gemstone') return 1;
    if (b.category === 'Gemstone' && a.category !== 'Gemstone') return -1;

    var idxA = priorityOrder.indexOf(a.category);
    var idxB = priorityOrder.indexOf(b.category);

    if (idxA !== -1 && idxB !== -1) {
      return idxA - idxB;
    }
    if (idxA !== -1) return -1;
    if (idxB !== -1) return 1;

    return a.category.localeCompare(b.category);
  });

  return configs;
}

/* ========================== DESIGN CODE HELPERS ============================ */

function extractDesignCodeFromSKU_(sku) {
  if (!sku) return '';
  var match = String(sku).match(/[-_](D\d+)/i);
  if (match && match[1]) {
    return match[1].toUpperCase();
  }
  return '';
}

function detectDesignCategory_(itemName, categoryConfigs) {
  if (!itemName) return '';
  var name = String(itemName);

  for (var i = 0; i < categoryConfigs.length; i++) {
    if (categoryConfigs[i].regex.test(name)) {
      return categoryConfigs[i].category;
    }
  }
  return '';
}

function extractSizeFromDesignSKU_(sku, designCode) {
  if (!sku || !designCode) return '';
  var upper = String(sku).toUpperCase();
  var dcUpper = designCode.toUpperCase();

  var idx = -1;
  var separators = ['-', '_'];
  for (var s = 0; s < separators.length; s++) {
    var search = separators[s] + dcUpper;
    var pos = upper.indexOf(search);
    if (pos !== -1) {
      idx = pos + search.length;
      break;
    }
  }

  if (idx === -1) return '';
  var suffix = String(sku).substring(idx).replace(/^[-_]+/, '');
  return suffix;
}

/* ============================ IMAGE HELPERS ================================ */

function updateDesignSheetImages_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var urlCol   = DS_COLS.IMAGE_URL;
  var imageCol = DS_COLS.IMAGE;

  var urls = sheet.getRange(startRow, urlCol, numRows, 1).getValues();
  var formulas = [];

  for (var i = 0; i < numRows; i++) {
    var url = String(urls[i][0]).trim();
    if (!url) {
      formulas.push(['']);
    } else {
      formulas.push(['=IMAGE("' + url.replace(/"/g, '""') + '")']);
    }
  }

  sheet.getRange(startRow, imageCol, numRows, 1).setFormulas(formulas);
}

/* ============================ MAIN FUNCTIONS =============================== */

function extractDesignCodes() {
  var startTime = new Date().getTime();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DESIGN_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Sheet "' + DESIGN_SHEET_NAME + '" not found in this spreadsheet.'
    );
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < DS_DATA_START) {
    SpreadsheetApp.getUi().alert('No data rows found in ' + DESIGN_SHEET_NAME + '.');
    return;
  }

  var headers = sheet.getRange(DS_HEADER_ROW, 1, 1, Math.max(lastCol, DS_ERROR_COL)).getValues()[0];
  var errorHeaderVal = (headers.length >= DS_ERROR_COL) ? String(headers[DS_ERROR_COL - 1]).trim() : '';
  if (errorHeaderVal.toLowerCase() !== 'error') {
    sheet.getRange(DS_HEADER_ROW, DS_ERROR_COL).setValue('Error');
  }

  var categoryConfigs;
  try {
    categoryConfigs = loadDesignCategoryConfig_();
  } catch (err) {
    SpreadsheetApp.getUi().alert('Config error: ' + err.message);
    return;
  }

  var numRows = lastRow - DS_DATA_START + 1;
  var dataRange = sheet.getRange(DS_DATA_START, 1, numRows, DS_COLS.IMAGE_URL);
  var data = dataRange.getValues();

  var designCodes  = [];
  var categories   = [];
  var errors       = [];
  var sizeSuffixes = [];
  var fixedUrls    = [];

  for (var i = 0; i < numRows; i++) {
    var sku      = String(data[i][DS_COLS.SKU - 1]).trim();
    var itemName = String(data[i][DS_COLS.ITEM_NAME - 1]).trim();
    var rawUrl   = String(data[i][DS_COLS.IMAGE_URL - 1]).trim();

    var dCode = extractDesignCodeFromSKU_(sku);
    designCodes.push(dCode);

    var cat = detectDesignCategory_(itemName, categoryConfigs);
    categories.push(cat);

    var sizeSuffix = extractSizeFromDesignSKU_(sku, dCode);
    sizeSuffixes.push(sizeSuffix);

    fixedUrls.push(fixUrl_(rawUrl));

    if (!dCode) {
      errors.push('skip');
    } else {
      errors.push('');
    }
  }

  var groupMap = {};
  for (var i = 0; i < numRows; i++) {
    var dCode = designCodes[i];
    var cat   = categories[i];

    if (!dCode) continue;

    var key = cat + '|' + dCode;
    if (!groupMap[key]) {
      groupMap[key] = [];
    }
    groupMap[key].push(i);
  }

  for (var key in groupMap) {
    var indices = groupMap[key];
    if (indices.length <= 1) continue;

    var firstIdx = indices[0];
    var firstSuffix = sizeSuffixes[firstIdx];

    for (var j = 1; j < indices.length; j++) {
      var rowIdx = indices[j];
      var thisSuffix = sizeSuffixes[rowIdx];

      if (thisSuffix !== firstSuffix) {
        errors[rowIdx] = 'skip';
      } else {
        errors[rowIdx] = 'duplicate';
      }
    }
  }

  var dcOut = designCodes.map(function(dc) { return [dc]; });
  sheet.getRange(DS_DATA_START, DS_COLS.DesignCode, numRows, 1).setValues(dcOut);

  var catOut = categories.map(function(c) { return [c]; });
  sheet.getRange(DS_DATA_START, DS_COLS.Design_Category, numRows, 1).setValues(catOut);

  var urlOut = fixedUrls.map(function(u) { return [u]; });
  sheet.getRange(DS_DATA_START, DS_COLS.IMAGE_URL, numRows, 1).setValues(urlOut);

  var errOut = errors.map(function(e) { return [e]; });
  sheet.getRange(DS_DATA_START, DS_ERROR_COL, numRows, 1).setValues(errOut);

  SpreadsheetApp.flush();

  ss.toast('Loading images from URLs... This may take a moment.', 'Design Tools', -1);
  updateDesignSheetImages_(sheet, DS_DATA_START, lastRow);
  ss.toast('Images loaded successfully!', 'Design Tools', 3);

  var elapsed = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  SpreadsheetApp.getUi().alert(
    'Done!\n\n' +
    'Processed ' + numRows + ' rows in ' + elapsed + ' seconds.\n' +
    'Design codes extracted, categories detected, URLs fixed, and images loaded.'
  );
}

/* =============================== SORTING =================================== */

function sortDesignSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DESIGN_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Sheet "' + DESIGN_SHEET_NAME + '" not found in this spreadsheet.'
    );
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < DS_DATA_START) {
    SpreadsheetApp.getUi().alert('No data rows to sort.');
    return;
  }

  var dataRange = sheet.getRange(DS_DATA_START, 1, lastRow - DS_DATA_START + 1, lastCol);

  dataRange.sort([
    { column: DS_COLS.Design_Category, ascending: true },
    { column: DS_COLS.DesignCode,      ascending: true }
  ]);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Sorted by Design Category → Design Code.',
    'Design Tools',
    5
  );
}

/* ============================= REMOVE SKIPPED ============================== */

function removeSkipped() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DESIGN_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Sheet "' + DESIGN_SHEET_NAME + '" not found in this spreadsheet.'
    );
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < DS_DATA_START) {
    SpreadsheetApp.getUi().alert('No data rows found.');
    return;
  }

  var numRows = lastRow - DS_DATA_START + 1;
  var range = sheet.getRange(DS_DATA_START, 1, numRows, lastCol);
  var values = range.getValues();
  var formulas = range.getFormulas();

  var writeGrid = [];
  var deleteCount = 0;

  for (var i = 0; i < numRows; i++) {
    var errorVal = String(values[i][DS_ERROR_COL - 1]).trim().toLowerCase();
    if (errorVal === 'skip') {
      deleteCount++;
    } else {
      var row = [];
      for (var c = 0; c < lastCol; c++) {
        var formula = formulas[i][c];
        if (formula) {
          row.push(formula);
        } else {
          row.push(values[i][c]);
        }
      }
      writeGrid.push(row);
    }
  }

  if (deleteCount === 0) {
    SpreadsheetApp.getUi().alert('No rows with error "skip" found.');
    return;
  }

  range.clearContent();

  if (writeGrid.length > 0) {
    sheet.getRange(DS_DATA_START, 1, writeGrid.length, lastCol).setValues(writeGrid);
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Done!\n\nRemoved ' + deleteCount + ' skipped row(s).'
  );
}

/* ============================ REMOVE DUPLICATES ============================ */

function removeDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DESIGN_SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Sheet "' + DESIGN_SHEET_NAME + '" not found in this spreadsheet.'
    );
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < DS_DATA_START) {
    SpreadsheetApp.getUi().alert('No data rows found.');
    return;
  }

  var numRows = lastRow - DS_DATA_START + 1;
  var range = sheet.getRange(DS_DATA_START, 1, numRows, lastCol);
  var values = range.getValues();
  var formulas = range.getFormulas();

  var writeGrid = [];
  var deleteCount = 0;

  for (var i = 0; i < numRows; i++) {
    var errorVal = String(values[i][DS_ERROR_COL - 1]).trim().toLowerCase();
    if (errorVal === 'duplicate') {
      deleteCount++;
    } else {
      var row = [];
      for (var c = 0; c < lastCol; c++) {
        var formula = formulas[i][c];
        if (formula) {
          row.push(formula);
        } else {
          row.push(values[i][c]);
        }
      }
      writeGrid.push(row);
    }
  }

  if (deleteCount === 0) {
    SpreadsheetApp.getUi().alert('No rows with error "duplicate" found.');
    return;
  }

  range.clearContent();

  if (writeGrid.length > 0) {
    sheet.getRange(DS_DATA_START, 1, writeGrid.length, lastCol).setValues(writeGrid);
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Done!\n\nRemoved ' + deleteCount + ' duplicate row(s).'
  );
}
