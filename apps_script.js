/**
 * ValuTrack Backend API - Google Apps Script
 * Deploy this script as a Web App:
 * 1. Open a Google Sheet.
 * 2. Click Extensions > Apps Script.
 * 3. Replace the code in Code.gs with this code.
 * 4. Click Save.
 * 5. Run the function 'setupDailyReminderTrigger' to enable morning notifications.
 * 6. Click Deploy > New deployment.
 * 7. Select "Web app" as the type.
 * 8. Set:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 9. Click Deploy, authorize permissions, and copy the Web App URL.
 */

var SHEET_NAME = "ValuTrack";
var ENGINEERS_SHEET_NAME = "Engineers";

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

function getEngineersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ENGINEERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ENGINEERS_SHEET_NAME);
    sheet.appendRow(["Name", "Phone", "Apikey"]);
    var headerRange = sheet.getRange(1, 1, 1, 3);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#0d0f14");
    headerRange.setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  try {
    // 1. Get Cases List
    var sheet = getSheet();
    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var data = [];
    
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var record = {};
      for (var j = 0; j < headers.length; j++) {
        var value = row[j];
        if (value instanceof Date) {
          record[headers[j]] = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        } else {
          record[headers[j]] = value;
        }
      }
      data.push(record);
    }
    
    // 2. Get Engineers Contact Details List
    var engSheet = getEngineersSheet();
    var engValues = engSheet.getDataRange().getValues();
    var engHeaders = engValues[0];
    var engineersData = [];
    
    for (var i = 1; i < engValues.length; i++) {
      var row = engValues[i];
      var record = {};
      for (var j = 0; j < engHeaders.length; j++) {
        record[engHeaders[j]] = row[j];
      }
      engineersData.push(record);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ 
      success: true, 
      data: data,
      engineers: engineersData
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    
    if (action === "add") {
      var sheet = getSheet();
      var values = sheet.getDataRange().getValues();
      var headers = values[0];
      var rowData = requestData.data;
      rowData.ID = rowData.ID || Utilities.getUuid();
      
      var newRow = [];
      for (var i = 0; i < headers.length; i++) {
        newRow.push(rowData[headers[i]] !== undefined ? rowData[headers[i]] : "");
      }
      sheet.appendRow(newRow);
      
      // Trigger background WhatsApp message to assigned engineer
      triggerCaseAssignmentNotification(rowData);
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: rowData }))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === "update") {
      var sheet = getSheet();
      var values = sheet.getDataRange().getValues();
      var headers = values[0];
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
      
      var oldEngineer = values[rowIndex - 1][headers.indexOf("Engineer")] || "";
      var newEngineer = rowData.Engineer || "";
      
      // Update fields
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        if (rowData[header] !== undefined) {
          sheet.getRange(rowIndex, j + 1).setValue(rowData[header]);
        }
      }
      
      // Trigger background WhatsApp message if assigned engineer changed or newly assigned
      if (newEngineer && newEngineer !== oldEngineer) {
        triggerCaseAssignmentNotification(rowData);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: rowData }))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === "delete") {
      var sheet = getSheet();
      var values = sheet.getDataRange().getValues();
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
        
    } else if (action === "saveEngineers") {
      var engineers = requestData.engineers || [];
      var engSheet = getEngineersSheet();
      
      // Clear old rows under headers
      if (engSheet.getLastRow() > 1) {
        engSheet.deleteRows(2, engSheet.getLastRow() - 1);
      }
      
      // Append new engineer contacts list
      for (var i = 0; i < engineers.length; i++) {
        var eng = engineers[i];
        engSheet.appendRow([eng.Name || "", eng.Phone || "", eng.Apikey || ""]);
      }
      
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

// Helper to push WhatsApp messages via CallMeBot gateway
function sendCallMeBotMessage(phone, apikey, text) {
  if (!phone || !apikey || !text) return false;
  try {
    // Clean phone number (CallMeBot requires international format without + sign, e.g. 919876543210)
    var cleanedPhone = phone.toString().replace(/[^0-9]/g, "");
    var url = "https://api.callmebot.com/whatsapp.php?phone=" + cleanedPhone + 
              "&text=" + encodeURIComponent(text) + 
              "&apikey=" + apikey;
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return response.getResponseCode() === 200;
  } catch (e) {
    Logger.log("CallMeBot error: " + e.toString());
    return false;
  }
}

// Composes valuation case details text and sends a WhatsApp message on assignment
function triggerCaseAssignmentNotification(file) {
  if (!file.Engineer) return;
  
  var engSheet = getEngineersSheet();
  var values = engSheet.getDataRange().getValues();
  var phone = "";
  var apikey = "";
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] == file.Engineer) {
      phone = values[i][1];
      apikey = values[i][2];
      break;
    }
  }
  
  if (phone && apikey) {
    var mapsQuery = file.Coordinates ? file.Coordinates : file.Location;
    var mapsLink = "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(mapsQuery);
    
    var msg = "🔔 *New Case Assigned!*\n\n" +
              "*Case ID:* " + file.FileNo + "\n" +
              "*Ref No:* " + (file.RefNo || "-") + "\n" +
              "*Client:* " + file.Owner + "\n" +
              "*Location:* " + file.Location + "\n" +
              "*Contact:* " + (file.Contact || "-") + "\n\n" +
              "📍 *Google Maps:* " + mapsLink + "\n\n" +
              "Please check complete details in your ValuTrack App.";
              
    sendCallMeBotMessage(phone, apikey, msg);
  }
}

// daily time-driven reminder checking active jobs
function sendMorningReminders() {
  var engSheet = getEngineersSheet();
  var engValues = engSheet.getDataRange().getValues();
  
  var engineerMap = {};
  for (var i = 1; i < engValues.length; i++) {
    var name = engValues[i][0];
    var phone = engValues[i][1];
    var apikey = engValues[i][2];
    if (name && phone && apikey) {
      engineerMap[name] = { phone: phone, apikey: apikey };
    }
  }
  
  var sheet = getSheet();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var activeCasesByEngineer = {};
  
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var record = {};
    for (var j = 0; j < headers.length; j++) {
      record[headers[j]] = row[j];
    }
    
    var status = record.Status;
    // Active statuses: Site Visit Pending, Site Visit Completed, Report Progress
    if (record.Engineer && (status === "Site Visit Pending" || status === "Site Visit Completed" || status === "Report Progress")) {
      var eng = record.Engineer;
      if (!activeCasesByEngineer[eng]) {
        activeCasesByEngineer[eng] = [];
      }
      activeCasesByEngineer[eng].push(record);
    }
  }
  
  for (var engName in activeCasesByEngineer) {
    var contact = engineerMap[engName];
    if (!contact) continue;
    
    var cases = activeCasesByEngineer[engName];
    if (cases.length === 0) continue;
    
    var msg = "☀️ *Good Morning, " + engName + "!*\n\n" +
              "Overview of active cases assigned to you (" + cases.length + " total):\n\n";
              
    for (var k = 0; k < cases.length; k++) {
      var c = cases[k];
      msg += "▪️ *" + c.FileNo + "* [" + c.Status + "]\n" +
             "  *Client:* " + c.Owner + "\n" +
             "  *Loc:* " + c.Location + "\n\n";
    }
    
    msg += "Link to dashboard: https://graceful-starship-658d6e.netlify.app";
    sendCallMeBotMessage(contact.phone, contact.apikey, msg);
  }
}

// Run this once inside the Apps Script Editor to set up the daily 8 AM timer trigger
function setupDailyReminderTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "sendMorningReminders") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger("sendMorningReminders")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

// Dynamic UUID Generator helper
function generateUUID() {
  return Utilities.getUuid();
}

// Helper to get Admin credentials for WhatsApp alerts
function getAdminContact() {
  try {
    var engSheet = getEngineersSheet();
    if (!engSheet) return null;
    var data = engSheet.getDataRange().getValues();
    var headers = data[0];
    var nameCol = headers.indexOf("Name");
    var phoneCol = headers.indexOf("Phone");
    var keyCol = headers.indexOf("Apikey");
    
    if (nameCol === -1 || phoneCol === -1 || keyCol === -1) return null;
    
    // Find Sathish or Udhaya first, otherwise fallback to first engineer
    for (var i = 1; i < data.length; i++) {
      var name = data[i][nameCol];
      if (name && (name.indexOf("Sathish") !== -1 || name.indexOf("Udhaya") !== -1)) {
        return {
          phone: data[i][phoneCol],
          apikey: data[i][keyCol]
        };
      }
    }
    if (data.length > 1) {
      return {
        phone: data[1][phoneCol],
        apikey: data[1][keyCol]
      };
    }
  } catch (e) {
    Logger.log("Error loading admin contact: " + e.message);
  }
  return null;
}

// NATIVE GMAIL AUTO-CAPTURE SYSTEM
// Automatically processes unread emails matching specific keywords
function importCasesFromGmail() {
  // Define search query (customizable - looking for unread valuation emails)
  var query = 'is:unread (subject:"valuation" OR subject:"property valuation" OR subject:"request" OR subject:"initiate")';
  var threads = GmailApp.search(query, 0, 10); // Process top 10 threads at a time
  
  if (threads.length === 0) {
    Logger.log("No new unread valuation emails found.");
    return;
  }
  
  var sheet = getSheet();
  var dataRange = sheet.getDataRange().getValues();
  var headers = dataRange[0];
  
  // Get index positions for columns
  var idCol = headers.indexOf("ID");
  var fileNoCol = headers.indexOf("FileNo");
  var ownerCol = headers.indexOf("Owner");
  var bankCol = headers.indexOf("Bank");
  var addressCol = headers.indexOf("Address");
  var priorityCol = headers.indexOf("Priority");
  var statusCol = headers.indexOf("Status");
  var dateCol = headers.indexOf("Date");
  var contactCol = headers.indexOf("Contact");
  var refNoCol = headers.indexOf("RefNo");
  
  var now = new Date();
  var yyyy = now.getFullYear();
  var mm = String(now.getMonth() + 1).padStart(2, '0');
  var dd = String(now.getDate()).padStart(2, '0');
  var todayStr = yyyy + "-" + mm + "-" + dd;
  
  var nextSerial = dataRange.length; // Approximate total cases for serial numbering
  
  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var msg = messages[j];
      
      // Skip if already read
      if (!msg.isUnread()) continue;
      
      var body = msg.getPlainBody();
      var subject = msg.getSubject();
      var sender = msg.getFrom();
      
      // 1. EXTRACT DATA VIA REGEX
      // A. Extract Owner/Customer Name
      var owner = "";
      var ownerMatch = body.match(/(?:owner|customer|borrower|applicant)?\s*name\s*:\s*([^\n\r]+)/i) || 
                       body.match(/valuation\s*for\s*:\s*([^\n\r]+)/i);
      if (ownerMatch) {
        owner = ownerMatch[1].trim();
      }
      
      // B. Extract Contact Number (looks for 10-12 digit numbers)
      var contact = "";
      var contactMatch = body.match(/(?:contact|phone|mobile|tel)?\s*(?:no|number)?\s*:\s*([+0-9\s-]{10,15})/i) ||
                         body.match(/\b\d{10}\b/);
      if (contactMatch) {
        contact = contactMatch[1].replace(/[^0-9+]/g, "").trim();
      }
      
      // C. Extract Reference Number
      var refNo = "";
      var refMatch = body.match(/(?:ref|reference|case)?\s*(?:no|number)?\s*:\s*([a-zA-Z0-9_-]+)/i);
      if (refMatch) {
        refNo = refMatch[1].trim();
      }
      
      // D. Determine Bank
      var bank = "Unknown Bank";
      var senderLower = sender.toLowerCase();
      var subjectLower = subject.toLowerCase();
      
      var bodyLower = body.toLowerCase();
      if (senderLower.indexOf("truhome") !== -1 || subjectLower.indexOf("truhome") !== -1 || bodyLower.indexOf("truhome") !== -1) bank = "Truhome";
      else if (senderLower.indexOf("hdfc") !== -1 || subjectLower.indexOf("hdfc") !== -1 || bodyLower.indexOf("hdfc") !== -1) bank = "HDFC";
      else if (senderLower.indexOf("equitas") !== -1 || subjectLower.indexOf("equitas") !== -1 || bodyLower.indexOf("equitas") !== -1) bank = "Equitas";
      else if (senderLower.indexOf("niwas") !== -1 || subjectLower.indexOf("niwas") !== -1 || bodyLower.indexOf("niwas") !== -1) bank = "Niwas";
      
      // 2. CHECK FOR DUPLICATES
      // If we found a name, let's verify if we already added it recently
      if (owner) {
        var isDuplicate = false;
        for (var r = 1; r < dataRange.length; r++) {
          var existingOwner = dataRange[r][ownerCol];
          var existingRef = refNoCol !== -1 ? dataRange[r][refNoCol] : "";
          if (existingOwner && existingOwner.toLowerCase() === owner.toLowerCase()) {
            isDuplicate = true;
            break;
          }
          if (refNo && existingRef && existingRef.toString() === refNo.toString()) {
            isDuplicate = true;
            break;
          }
        }
        
        if (isDuplicate) {
          Logger.log("Skipping duplicate case: " + owner);
          msg.markRead();
          continue;
        }
        
        // 3. GENERATE AUTOMATIC ROW DATA
        var fileId = generateUUID();
        var fileNo = "VT-" + yyyy + "-" + mm + "-" + dd + "-" + nextSerial;
        nextSerial++;
        
        // Create full row array matching headers
        var newRow = [];
        for (var c = 0; c < headers.length; c++) {
          var header = headers[c];
          if (header === "ID") newRow.push(fileId);
          else if (header === "FileNo") newRow.push(fileNo);
          else if (header === "Owner") newRow.push(owner);
          else if (header === "Bank") newRow.push(bank);
          else if (header === "Address") newRow.push("");
          else if (header === "Location") newRow.push("");
          else if (header === "Engineer") newRow.push(""); // Unassigned initially
          else if (header === "Priority") newRow.push("High");
          else if (header === "Status") newRow.push("Site Visit Pending");
          else if (header === "Date") newRow.push(todayStr);
          else if (header === "Loan") newRow.push("");
          else if (header === "DriveLink") newRow.push("");
          else if (header === "Remarks") newRow.push("Automatically captured from Gmail");
          else if (header === "History") newRow.push("[]");
          else if (header === "UpdatedBy") newRow.push("Gmail Auto-Capture");
          else if (header === "UpdatedAt") newRow.push(todayStr + " 00:00:00");
          else if (header === "RefNo") newRow.push(refNo);
          else if (header === "Coordinates") newRow.push("");
          else if (header === "Contact") newRow.push(contact);
          else newRow.push("");
        }
        
        // Save to spreadsheet
        sheet.appendRow(newRow);
        Logger.log("Successfully imported case from Gmail: " + owner);
        
        // 4. SEND WHATSAPP NOTIFICATION TO ADMIN
        var admin = getAdminContact();
        if (admin && admin.phone && admin.apikey) {
          var text = "🔔 *ValuTrack Auto-Import*\n\n" +
                     "A new case has been captured from your Gmail inbox:\n\n" +
                     "▪️ *Case No:* " + fileNo + "\n" +
                     "▪️ *Owner:* " + owner + "\n" +
                     "▪️ *Bank:* " + bank + "\n" +
                     "▪️ *Contact:* " + (contact || "Not specified") + "\n\n" +
                     "Open dashboard: https://graceful-starship-658d6e.netlify.app";
          sendCallMeBotMessage(admin.phone, admin.apikey, text);
        }
      }
      
      // Mark as processed
      msg.markRead();
    }
  }
}

// Run this once inside the Apps Script Editor to check Gmail every 10 minutes automatically
function setupGmailImportTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "importCasesFromGmail") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger("importCasesFromGmail")
    .timeBased()
    .everyMinutes(10)
    .create();
}
