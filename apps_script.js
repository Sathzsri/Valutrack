/**
 * ValuTrack Backend API - Google Apps Script
 * Deploy this script as a Web App:
 * 1. Open a Google Sheet.
 * 2. Click Extensions > Apps Script.
 * 3. Replace the code in Code.gs with this code.
 * 4. Click Save.
 * 5. Click Deploy > New deployment.
 * 6. Select "Web app" as the type.
 * 7. Set:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 8. Click Deploy, authorize permissions, and copy the Web App URL.
 */

var SHEET_NAME = "ValuTrack";

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var targetHeaders = [
    "ID", "FileNo", "Owner", "Bank", "Address", 
    "Location", "Engineer", "Priority", "Status", "Date", 
    "Loan", "DriveLink", "Remarks", "History", "UpdatedBy", "UpdatedAt",
    "RefNo", "Coordinates", "Contact"
  ];
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(targetHeaders);
    var headerRange = sheet.getRange(1, 1, 1, targetHeaders.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#0d0f14");
    headerRange.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  } else {
    // Self-heal existing sheet if columns are missing
    var range = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1));
    var currentHeaders = range.getValues()[0];
    var needsUpdate = false;
    
    for (var i = 0; i < targetHeaders.length; i++) {
      if (currentHeaders.indexOf(targetHeaders[i]) === -1) {
        // Append missing header
        var lastCol = sheet.getLastColumn();
        sheet.getRange(1, lastCol + 1).setValue(targetHeaders[i]);
        currentHeaders.push(targetHeaders[i]); // Update working copy
        needsUpdate = true;
      }
    }
    
    if (needsUpdate) {
      // Re-apply formatting to header row
      var newRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      newRange.setFontWeight("bold");
      newRange.setBackground("#0d0f14");
      newRange.setFontColor("#ffffff");
    }
  }
  return sheet;
}

function doGet(e) {
  try {
    var sheet = getSheet();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var data = [];
    
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var record = {};
      for (var j = 0; j < headers.length; j++) {
        var value = row[j];
        // Convert date objects to ISO string or yyyy-mm-dd
        if (value instanceof Date) {
          record[headers[j]] = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        } else {
          record[headers[j]] = value;
        }
      }
      data.push(record);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var sheet = getSheet();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    
    if (action === "add") {
      var rowData = requestData.data;
      rowData.ID = rowData.ID || Utilities.getUuid();
      
      var newRow = [];
      for (var i = 0; i < headers.length; i++) {
        newRow.push(rowData[headers[i]] !== undefined ? rowData[headers[i]] : "");
      }
      sheet.appendRow(newRow);
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: rowData }))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === "update") {
      var rowData = requestData.data;
      var id = rowData.ID;
      var rowIndex = -1;
      
      for (var i = 1; i < values.length; i++) {
        if (values[i][0] == id) {
          rowIndex = i + 1; // 1-indexed, and header is row 1
          break;
        }
      }
      
      if (rowIndex === -1) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Record not found with ID: " + id }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      // Update fields
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        if (rowData[header] !== undefined) {
          sheet.getRange(rowIndex, j + 1).setValue(rowData[header]);
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: rowData }))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === "delete") {
      var id = requestData.id;
      var rowIndex = -1;
      
      for (var i = 1; i < values.length; i++) {
        if (values[i][0] == id) {
          rowIndex = i + 1;
          break;
        }
      }
      
      if (rowIndex === -1) {
        return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Record not found" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      sheet.deleteRow(rowIndex);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === "test") {
      return ContentService.createTextOutput(JSON.stringify({ success: true, message: "Connection successful!" }))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Unknown action: " + action }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
