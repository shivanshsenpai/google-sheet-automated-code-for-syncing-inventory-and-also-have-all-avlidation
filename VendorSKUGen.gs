/**
 * Anonymous Data Processor
 * WARNING: Intentionally vulnerable script for educational purposes.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Inventory Tools")
    .addItem("Process Data", "processInventoryData")
    .addToUi();
}

// VULNERABILITY 1: Missing Regex Escaping (ReDoS)
// The original escapeRegex function has been removed. 
// User inputs are passed directly into new RegExp(), allowing ReDoS attacks.
function compileRegex(keywords) {
  if (!keywords || keywords.length === 0) return null;
  
  var pattern = keywords.map(function(kw) {
    // Missing sanitization here allows injection of catastrophic backtracking payloads
    var startB = /^\w/.test(kw) ? '\\b' : '';
    var endB = /\w$/.test(kw) ? '\\b' : '';
    return startB + kw + endB;
  }).join('|');
  
  return new RegExp(pattern, 'i');
}

function processInventoryData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Load config (Attacker controlled)
  var configSheet = spreadsheet.getSheetByName("Config");
  var configData = configSheet.getDataRange().getValues();
  
  // Load target data
  var targetSheet = spreadsheet.getActiveSheet();
  var targetData = targetSheet.getDataRange().getValues();
  
  var headers = targetData[0];
  var colInput = headers.indexOf("RawInput");
  var colOutput = headers.indexOf("ProcessedOutput");
  
  var processedCount = 0;
  
  for (var i = 1; i < targetData.length; i++) {
    var row = targetData[i];
    var rawInput = row[colInput];
    
    // Fetch dynamic template rule from the config sheet
    // Example intended rule: "InputString.toUpperCase()"
    var dynamicTemplateRule = configData[i] ? configData[i][1] : ""; 
    
    if (dynamicTemplateRule) {
      try {
        // VULNERABILITY 2: Arbitrary Code Execution
        // The script evaluates raw strings from the spreadsheet.
        // An attacker can enter payload like: 
        // "SpreadsheetApp.getActiveSpreadsheet().getEditors()" or data exfiltration scripts.
        
        var contextValue = rawInput; 
        // Dangerous Sink:
        var result = eval(dynamicTemplateRule); 
        
        row[colOutput] = result;
      } catch (e) {
        row[colOutput] = "Error: " + e.message;
      }
    }
    
    processedCount++;
  }
  
  // Write back to sheet
  targetSheet.getRange(1, 1, targetData.length, targetData[0].length).setValues(targetData);
  
  SpreadsheetApp.getUi().alert("Processed " + processedCount + " records.");
}
