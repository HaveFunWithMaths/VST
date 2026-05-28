/**
 * Google Apps Script Web App for Bank Statement Spreadsheet Updater
 * Deployed directly inside the target Google Sheet.
 * 
 * Instructions:
 * 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1a4UOK-OLhu70Rop3OQx3QGjPdLWhnvuPYPKaFt8J0kk
 * 2. Go to Extensions -> Apps Script.
 * 3. Delete any default code and paste this script.
 * 4. Click the Save icon (floppy disk).
 * 5. Click Deploy -> New deployment.
 * 6. Under "Select type", click the gear icon and select "Web app".
 * 7. Set:
 *    - Description: "Bank Statement Updater API"
 *    - Execute as: "Me (your email)"
 *    - Who has access: "Anyone" (Required so the local web app can contact this API without complex OAuth setup).
 * 8. Click Deploy. Authorize access if prompted.
 * 9. Copy the "Web app URL" and paste it into the Web App's URL field.
 */

function doGet(e) {
  var output = "";
  try {
    var info = getLastUpdateDateInfo();
    output = JSON.stringify({
      status: "success",
      lastUpdateDate: info.lastUpdateDate,
      initialBalance: info.initialBalance
    });
  } catch (err) {
    output = JSON.stringify({
      status: "error",
      message: err.toString()
    });
  }
  
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var output = "";
  try {
    var postData = JSON.parse(e.postData.contents);
    var transactions = postData.transactions; // List of parsed transactions
    var lastUpdateDateStr = postData.lastUpdateDate; // "YYYY-MM-DD" or null
    var startBalance = postData.startBalance; // Optional manual start balance
    
    updateLedger(transactions, lastUpdateDateStr, startBalance);
    
    output = JSON.stringify({
      status: "success",
      message: "Google Sheet successfully updated!"
    });
  } catch (err) {
    output = JSON.stringify({
      status: "error",
      message: err.toString()
    });
  }
  
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

// Global list of months in the financial year
var MONTHS_MAP = {
  "04": "Apr 2026",
  "05": "May 2026",
  "06": "Jun 2026",
  "07": "Jul 2026",
  "08": "Aug 2026",
  "09": "Sep 2026",
  "10": "Oct 2026",
  "11": "Nov 2026",
  "12": "Dec 2026",
  "01": "Jan 2027",
  "02": "Feb 2027",
  "03": "Mar 2027"
};

var MONTHS_LIST = [
  "Apr 2026", "May 2026", "Jun 2026", "Jul 2026", "Aug 2026", "Sep 2026",
  "Oct 2026", "Nov 2026", "Dec 2026", "Jan 2027", "Feb 2027", "Mar 2027"
];

function getLastUpdateDateInfo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lastUpdateDate = null;
  var initialBalance = 0.0;
  
  // Try to find first sheet's opening balance from April 2026 row 4
  var aprSheet = ss.getSheetByName("Apr 2026");
  if (aprSheet) {
    initialBalance = aprSheet.getRange("B4").getValue();
  }
  
  for (var i = 0; i < MONTHS_LIST.length; i++) {
    var sheetName = MONTHS_LIST[i];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;
    
    var values = sheet.getRange(1, 1, lastRow, 7).getValues();
    var tallyRowIdx = -1;
    for (var r = 3; r < values.length; r++) { // Row 4 starts at index 3
      if (values[r][0] === "Tally") {
        tallyRowIdx = r;
        break;
      }
    }
    
    if (tallyRowIdx === -1) continue;
    
    for (var r = 3; r < tallyRowIdx; r++) {
      var row = values[r];
      var dateVal = row[0];
      var inflow = row[2];
      var inflowDesc = row[3];
      var outflow = row[4];
      var outflowDesc = row[5];
      
      var isReal = false;
      if (inflow && parseFloat(inflow) !== 0) isReal = true;
      if (outflow && parseFloat(outflow) !== 0) isReal = true;
      if (inflowDesc && inflowDesc !== "-" && inflowDesc !== "") isReal = true;
      if (outflowDesc && outflowDesc !== "-" && outflowDesc !== "") isReal = true;
      
      if (isReal && dateVal instanceof Date) {
        lastUpdateDate = formatDate(dateVal);
      }
    }
  }
  
  return {
    lastUpdateDate: lastUpdateDate,
    initialBalance: initialBalance
  };
}

function updateLedger(newTxs, lastUpdateDateStr, startBalance) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // If startBalance is provided and Apr 2026 opening balance is 0 or empty, update it
  var aprSheet = ss.getSheetByName("Apr 2026");
  if (aprSheet && startBalance !== undefined && startBalance !== null) {
    var currentAprStart = aprSheet.getRange("B4").getValue();
    if (!currentAprStart || parseFloat(currentAprStart) === 0) {
      aprSheet.getRange("B4").setValue(startBalance);
    }
  }
  
  // 1. Group statement transactions by month
  var txsByMonth = {};
  for (var i = 0; i < MONTHS_LIST.length; i++) {
    txsByMonth[MONTHS_LIST[i]] = [];
  }
  
  for (var i = 0; i < newTxs.length; i++) {
    var tx = newTxs[i];
    var parts = tx.date.split("-"); // YYYY-MM-DD
    var monthKey = parts[1];
    var sheetName = MONTHS_MAP[monthKey];
    if (sheetName && txsByMonth[sheetName]) {
      txsByMonth[sheetName].push(tx);
    }
  }
  
  // Sort transactions inside each month chronologically
  for (var sheetName in txsByMonth) {
    txsByMonth[sheetName].sort(function(a, b) {
      return a.date.localeCompare(b.date) || a.row_index - b.row_index;
    });
  }
  
  // Find which months need to be modified
  var firstModifiedMonthIdx = 12;
  var cutoffDate = lastUpdateDateStr ? new Date(lastUpdateDateStr) : null;
  
  if (cutoffDate) {
    // If we have a cutoff date, find the first month that contains transactions >= cutoffDate
    for (var i = 0; i < MONTHS_LIST.length; i++) {
      var sName = MONTHS_LIST[i];
      var monthTxs = txsByMonth[sName] || [];
      var hasUpdates = false;
      for (var k = 0; k < monthTxs.length; k++) {
        if (new Date(monthTxs[k].date) >= cutoffDate) {
          hasUpdates = true;
          break;
        }
      }
      if (hasUpdates) {
        firstModifiedMonthIdx = i;
        break;
      }
    }
    // Also, if the first modified month index is still 12, check which month the cutoffDate falls in
    if (firstModifiedMonthIdx === 12) {
      var cutoffMonth = cutoffDate.getMonth() + 1; // 1-12
      var cutoffYear = cutoffDate.getFullYear();
      var monthStr = (cutoffMonth < 10 ? "0" : "") + cutoffMonth;
      var targetSheet = MONTHS_MAP[monthStr];
      if (targetSheet) {
        firstModifiedMonthIdx = MONTHS_LIST.indexOf(targetSheet);
      }
    }
  } else {
    firstModifiedMonthIdx = 0; // Update all months from April
  }
  
  if (firstModifiedMonthIdx === -1 || firstModifiedMonthIdx === 12) {
    firstModifiedMonthIdx = 0; // Fallback
  }
  
  var tallyRowsMap = {};
  
  // 2. Scan preceding months to capture their current Tally rows
  for (var i = 0; i < firstModifiedMonthIdx; i++) {
    var sheet = ss.getSheetByName(MONTHS_LIST[i]);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      var values = sheet.getRange(1, 1, lastRow, 1).getValues();
      for (var r = 0; r < values.length; r++) {
        if (values[r][0] === "Tally") {
          tallyRowsMap[MONTHS_LIST[i]] = r + 1; // 1-indexed
          break;
        }
      }
    }
  }
  
  // 3. Process sheets from firstModifiedMonthIdx onwards
  for (var i = firstModifiedMonthIdx; i < MONTHS_LIST.length; i++) {
    var sheetName = MONTHS_LIST[i];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    
    // Parse sheet month & year
    var parts = sheetName.split(" ");
    var monStr = parts[0];
    var year = parseInt(parts[1]);
    var monthMap = {"Apr":4, "May":5, "Jun":6, "Jul":7, "Aug":8, "Sep":9, "Oct":10, "Nov":11, "Dec":12, "Jan":1, "Feb":2, "Mar":3};
    var month = monthMap[monStr];
    
    // Cache Styles from Row 4 and the Tally Row
    var cachedStyles = cacheStyles(sheet);
    
    // Get existing transactions in this sheet if it's the active cutoff month
    var preservedTxs = [];
    if (cutoffDate && i === firstModifiedMonthIdx) {
      preservedTxs = getExistingTransactionsBeforeDate(sheet, cutoffDate);
    }
    
    // Get statement transactions for this month
    var statementTxs = txsByMonth[sheetName] || [];
    
    // Filter statement transactions to only keep dates >= cutoffDate
    var filteredStatementTxs = [];
    for (var k = 0; k < statementTxs.length; k++) {
      if (!cutoffDate || new Date(statementTxs[k].date) >= cutoffDate) {
        filteredStatementTxs.push(statementTxs[k]);
      }
    }
    
    // Combine preserved transactions and new transactions
    var mergedTxs = preservedTxs.concat(filteredStatementTxs);
    
    // Group transactions by day of the month
    var txByDay = {};
    for (var k = 0; k < mergedTxs.length; k++) {
      var txDate = new Date(mergedTxs[k].date);
      var day = txDate.getDate();
      if (!txByDay[day]) txByDay[day] = [];
      txByDay[day].push(mergedTxs[k]);
    }
    
    var numDays = daysInMonth(month, year);
    
    // Prepare the data to write
    var rowData = [];
    var cellFormulas = [];
    
    var currentWritingRow = 4;
    
    for (var day = 1; day <= numDays; day++) {
      var dayTxs = txByDay[day] || [];
      if (dayTxs.length === 0) {
        // Placeholder row
        var rowDate = new Date(year, month - 1, day);
        var opFormula = "";
        if (sheetName === "Apr 2026" && currentWritingRow === 4) {
          opFormula = startBalance !== undefined ? startBalance : 0.0;
        } else if (currentWritingRow === 4) {
          var prevSheet = MONTHS_LIST[i - 1];
          var prevTally = tallyRowsMap[prevSheet];
          opFormula = "='" + prevSheet + "'!G" + prevTally;
        } else {
          opFormula = "=G" + (currentWritingRow - 1);
        }
        
        rowData.push([
          rowDate,
          opFormula,
          0.0,
          "-",
          0.0,
          "-"
        ]);
        cellFormulas.push([
          "",
          opFormula.toString().indexOf("=") === 0 ? opFormula : "",
          "",
          "",
          "",
          "",
          "=B" + currentWritingRow + "+C" + currentWritingRow + "-E" + currentWritingRow
        ]);
        currentWritingRow++;
      } else {
        // Write each transaction
        for (var txIdx = 0; txIdx < dayTxs.length; txIdx++) {
          var tx = dayTxs[txIdx];
          var rowDate = new Date(tx.date);
          
          var opFormula = "";
          if (sheetName === "Apr 2026" && currentWritingRow === 4) {
            opFormula = startBalance !== undefined ? startBalance : 0.0;
          } else if (currentWritingRow === 4) {
            var prevSheet = MONTHS_LIST[i - 1];
            var prevTally = tallyRowsMap[prevSheet];
            opFormula = "='" + prevSheet + "'!G" + prevTally;
          } else {
            opFormula = "=G" + (currentWritingRow - 1);
          }
          
          var infVal = 0.0;
          var infPart = "-";
          var outVal = 0.0;
          var outPart = "-";
          
          var particulars = tx.description;
          if (tx.ref_no && tx.ref_no !== "-" && particulars.indexOf(tx.ref_no) === -1) {
            particulars = particulars + " / " + tx.ref_no;
          }
          
          if (tx.dr_cr === "CR") {
            infVal = tx.amount;
            infPart = particulars;
          } else {
            outVal = tx.amount;
            outPart = particulars;
          }
          
          rowData.push([
            rowDate,
            opFormula,
            infVal,
            infPart,
            outVal,
            outPart
          ]);
          
          cellFormulas.push([
            "",
            opFormula.toString().indexOf("=") === 0 ? opFormula : "",
            "",
            "",
            "",
            "",
            "=B" + currentWritingRow + "+C" + currentWritingRow + "-E" + currentWritingRow
          ]);
          currentWritingRow++;
        }
      }
    }
    
    // Tally row details
    var tallyRow = currentWritingRow;
    tallyRowsMap[sheetName] = tallyRow;
    
    var tallyRowData = [
      "Tally",
      "=B4",
      "=SUBTOTAL(109,C4:C" + (tallyRow - 1) + ")",
      null,
      "=SUBTOTAL(109,E4:E" + (tallyRow - 1) + ")",
      "-",
      "='" + sheetName + "'!$B$" + tallyRow + "+'" + sheetName + "'!$C$" + tallyRow + "-'" + sheetName + "'!$E$" + tallyRow
    ];
    
    // Clear everything from row 4 down to prevent trailing junk rows
    var maxRow = sheet.getMaxRows();
    if (maxRow >= 4) {
      // Clear values and formatting
      sheet.getRange(4, 1, maxRow - 3, 7).clear({contentsOnly: true, formatOnly: true});
    }
    
    // Write Data rows
    var dataRange = sheet.getRange(4, 1, rowData.length, 6);
    dataRange.setValues(rowData);
    
    // Write Formulas
    var formulaRange = sheet.getRange(4, 1, cellFormulas.length, 7);
    formulaRange.setFormulas(cellFormulas);
    
    // Write Tally row values & formulas
    var tallyRange = sheet.getRange(tallyRow, 1, 1, 7);
    var tallyValues = [];
    var tallyFormulas = [];
    for (var col = 0; col < 7; col++) {
      var val = tallyRowData[col];
      if (val && val.toString().indexOf("=") === 0) {
        tallyValues.push("");
        tallyFormulas.push(val);
      } else {
        tallyValues.push(val);
        tallyFormulas.push("");
      }
    }
    tallyRange.setValues([tallyValues]);
    tallyRange.setFormulas([tallyFormulas]);
    
    // Reapply Styles to Data rows
    var fullDataRange = sheet.getRange(4, 1, rowData.length, 7);
    applyStylesToRange(sheet, fullDataRange, cachedStyles.dataStyles);
    
    // Reapply Styles to Tally row
    applyStylesToRange(sheet, tallyRange, cachedStyles.tallyStyles);
    
    // Set heights
    for (var r = 4; r < tallyRow; r++) {
      sheet.setRowHeight(r, cachedStyles.dataRowHeight);
    }
    sheet.setRowHeight(tallyRow, cachedStyles.tallyRowHeight);
    
    // Delete any excess rows at the bottom
    var currentMaxRow = sheet.getMaxRows();
    if (currentMaxRow > tallyRow) {
      sheet.deleteRows(tallyRow + 1, currentMaxRow - tallyRow);
    }
    
    // Resize Spreadsheet Tables
    var tables = sheet.getTables();
    if (tables && tables.length > 0) {
      for (var t = 0; t < tables.length; t++) {
        // Set table size from header (row 3) to tallyRow
        tables[t].resize(sheet.getRange(3, 1, tallyRow - 2, 7));
      }
    }
  }
}

function getExistingTransactionsBeforeDate(sheet, cutoffDate) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 4) return [];
  
  var values = sheet.getRange(1, 1, lastRow, 7).getValues();
  var tallyRowIdx = -1;
  for (var r = 3; r < values.length; r++) {
    if (values[r][0] === "Tally") {
      tallyRowIdx = r;
      break;
    }
  }
  
  if (tallyRowIdx === -1) tallyRowIdx = lastRow;
  
  var txs = [];
  for (var r = 3; r < tallyRowIdx; r++) {
    var row = values[r];
    var dateVal = row[0];
    var inflow = parseFloat(row[2]) || 0;
    var inflowDesc = row[3];
    var outflow = parseFloat(row[4]) || 0;
    var outflowDesc = row[5];
    var balance = parseFloat(row[6]) || 0;
    
    var isReal = false;
    if (inflow !== 0 || outflow !== 0) isReal = true;
    if (inflowDesc && inflowDesc !== "-" && inflowDesc !== "") isReal = true;
    if (outflowDesc && outflowDesc !== "-" && outflowDesc !== "") isReal = true;
    
    if (isReal && dateVal instanceof Date) {
      var dateStr = formatDate(dateVal);
      var checkDate = new Date(dateStr);
      if (checkDate <= cutoffDate) {
        txs.push({
          date: dateStr,
          description: inflow !== 0 ? inflowDesc : outflowDesc,
          ref_no: "",
          amount: inflow !== 0 ? inflow : outflow,
          dr_cr: inflow !== 0 ? "CR" : "DR",
          balance: balance,
          row_index: r
        });
      }
    }
  }
  
  return txs;
}

function cacheStyles(sheet) {
  var lastRow = sheet.getLastRow();
  
  // Find tally row index
  var tallyRowIdx = 35; // Default fallback
  if (lastRow >= 4) {
    var vals = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (var r = 3; r < vals.length; r++) {
      if (vals[r][0] === "Tally") {
        tallyRowIdx = r + 1;
        break;
      }
    }
  }
  
  var dataRowHeight = sheet.getRowHeight(4) || 20;
  var tallyRowHeight = sheet.getRowHeight(tallyRowIdx) || 22;
  
  // Cache fonts, fills, alignments, borders, number formats for col 1 to 7
  var dataStyles = [];
  var tallyStyles = [];
  
  for (var col = 1; col <= 7; col++) {
    var dCell = sheet.getRange(4, col);
    var tCell = sheet.getRange(tallyRowIdx, col);
    
    dataStyles.push({
      fontColor: dCell.getFontColor(),
      fontFamily: dCell.getFontFamily(),
      fontSize: dCell.getFontSize(),
      fontStyle: dCell.getFontStyle(),
      fontWeight: dCell.getFontWeight(),
      background: dCell.getBackground(),
      horizontalAlignment: dCell.getHorizontalAlignment(),
      verticalAlignment: dCell.getVerticalAlignment(),
      numberFormat: dCell.getNumberFormat(),
      border: getCellBorder(dCell)
    });
    
    tallyStyles.push({
      fontColor: tCell.getFontColor(),
      fontFamily: tCell.getFontFamily(),
      fontSize: tCell.getFontSize(),
      fontStyle: tCell.getFontStyle(),
      fontWeight: tCell.getFontWeight(),
      background: tCell.getBackground(),
      horizontalAlignment: tCell.getHorizontalAlignment(),
      verticalAlignment: tCell.getVerticalAlignment(),
      numberFormat: tCell.getNumberFormat(),
      border: getCellBorder(tCell)
    });
  }
  
  return {
    dataStyles: dataStyles,
    tallyStyles: tallyStyles,
    dataRowHeight: dataRowHeight,
    tallyRowHeight: tallyRowHeight
  };
}

function getCellBorder(range) {
  // Returns border configuration
  // Since Apps Script doesn't have an easy getBorder() function, we will return null
  // and set default thin borders in our range style application, but we will preserve
  // backgrounds, fonts, alignments, and number formats which are the most important elements.
  return null;
}

function applyStylesToRange(sheet, range, styles) {
  var numRows = range.getNumRows();
  var startRow = range.getRow();
  var startCol = range.getColumn();
  
  for (var colIdx = 0; colIdx < styles.length; colIdx++) {
    var style = styles[colIdx];
    var colRange = sheet.getRange(startRow, startCol + colIdx, numRows, 1);
    
    colRange.setFontColor(style.fontColor);
    colRange.setFontFamily(style.fontFamily);
    colRange.setFontSize(style.fontSize);
    colRange.setFontStyle(style.fontStyle);
    colRange.setFontWeight(style.fontWeight);
    colRange.setBackground(style.background);
    colRange.setHorizontalAlignment(style.horizontalAlignment);
    colRange.setVerticalAlignment(style.verticalAlignment);
    
    if (style.numberFormat) {
      colRange.setNumberFormat(style.numberFormat);
    }
    
    // Draw gridlines
    colRange.setBorder(true, true, true, true, false, false, "#E2E8F0", SpreadsheetApp.BorderStyle.SOLID);
  }
}

function daysInMonth(month, year) {
  return new Date(year, month, 0).getDate();
}

function formatDate(date) {
  var d = new Date(date);
  var month = "" + (d.getMonth() + 1);
  var day = "" + d.getDate();
  var year = d.getFullYear();

  if (month.length < 2) month = "0" + month;
  if (day.length < 2) day = "0" + day;

  return [year, month, day].join("-");
}
