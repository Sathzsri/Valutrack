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
    "RefNo", "Coordinates", "Contact", "Contact2", "Contact3"
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
    deduplicateSheetRows(sheet);
    var values = sheet.getDataRange().getDisplayValues();
    var headers = values[0];
    var data = [];
    
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var record = {};
      for (var j = 0; j < headers.length; j++) {
        var value = row[j];
        record[headers[j]] = value;
      }
      data.push(record);
    }
    
    // 2. Get Engineers Contact Details List
    var engSheet = getEngineersSheet();
    var engValues = engSheet.getDataRange().getDisplayValues();
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
      var values = sheet.getDataRange().getDisplayValues();
      var headers = values[0];
      var rowData = requestData.data;
      var id = rowData.ID || Utilities.getUuid();
      rowData.ID = id;
      
      var rowIndex = -1;
      var idColIndex = headers.indexOf("ID");
      if (idColIndex === -1) idColIndex = 0;
      
      for (var i = 1; i < values.length; i++) {
        if (values[i][idColIndex] == id) {
          rowIndex = i + 1;
          break;
        }
      }
      
      if (rowIndex !== -1) {
        // Record already exists, update instead of append to prevent duplicates
        for (var j = 0; j < headers.length; j++) {
          var header = headers[j];
          if (rowData[header] !== undefined) {
            sheet.getRange(rowIndex, j + 1).setValue(rowData[header]);
          }
        }
      } else {
        var newRow = [];
        for (var i = 0; i < headers.length; i++) {
          newRow.push(rowData[headers[i]] !== undefined ? rowData[headers[i]] : "");
        }
        sheet.appendRow(newRow);
      }
      
      // Trigger background WhatsApp message to assigned engineer
      triggerCaseAssignmentNotification(rowData);
      
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: rowData }))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === "update") {
      var sheet = getSheet();
      var values = sheet.getDataRange().getDisplayValues();
      var headers = values[0];
      var rowData = requestData.data;
      var id = rowData.ID;
      var rowIndex = -1;
      var idColIndex = headers.indexOf("ID");
      if (idColIndex === -1) idColIndex = 0;
      
      for (var i = 1; i < values.length; i++) {
        if (values[i][idColIndex] == id) {
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
      var values = sheet.getDataRange().getDisplayValues();
      var id = requestData.id;
      var rowIndex = -1;
      var idColIndex = -1;
      if (values.length > 0) {
        idColIndex = values[0].indexOf("ID");
      }
      if (idColIndex === -1) idColIndex = 0;
      
      for (var i = 1; i < values.length; i++) {
        if (values[i][idColIndex] == id) {
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
  var values = engSheet.getDataRange().getDisplayValues();
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
  var engValues = engSheet.getDataRange().getDisplayValues();
  
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
  var values = sheet.getDataRange().getDisplayValues();
  var headers = values[0];
  var activeCasesByEngineer = {};
  
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var record = {};
    for (var j = 0; j < headers.length; j++) {
      record[headers[j]] = row[j];
    }
    
    var status = record.Status;
    // Active statuses: Site Visit Pending
    if (record.Engineer && status === "Site Visit Pending") {
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
    var data = engSheet.getDataRange().getDisplayValues();
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
  var query = 'is:unread (subject:valuation OR subject:initiate OR subject:request OR subject:tech OR subject:technical)';
  var threads = GmailApp.search(query, 0, 15); // Process top 15 threads at a time
  
  if (threads.length === 0) {
    Logger.log("No new unread valuation emails found.");
    return;
  }
  
  var sheet = getSheet();
  var dataRange = sheet.getDataRange().getDisplayValues();
  var headers = dataRange[0];
  
  // Get index positions for columns
  var idCol = headers.indexOf("ID");
  var fileNoCol = headers.indexOf("FileNo");
  var ownerCol = headers.indexOf("Owner");
  var bankCol = headers.indexOf("Bank");
  var addressCol = headers.indexOf("Address");
  var locationCol = headers.indexOf("Location");
  var priorityCol = headers.indexOf("Priority");
  var statusCol = headers.indexOf("Status");
  var dateCol = headers.indexOf("Date");
  var contactCol = headers.indexOf("Contact");
  var contact2Col = headers.indexOf("Contact2");
  var contact3Col = headers.indexOf("Contact3");
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
      var html = msg.getBody();
      var subject = msg.getSubject();
      var sender = msg.getFrom();
      
      var senderLower = sender.toLowerCase();
      var subjectLower = subject.toLowerCase();
      var bodyLower = body.toLowerCase();
      
      // Check for horizontal table data (HTML or plain text fallback)
      var horizTable = parseHtmlTable(html) || parseHorizontalTable(body);
      
      // 1. DETERMINE BANK
      var bank = "Others";
      function checkKeywords(keywords) {
        for (var k = 0; k < keywords.length; k++) {
          var kw = keywords[k].toLowerCase();
          if (senderLower.indexOf(kw) !== -1 || subjectLower.indexOf(kw) !== -1 || bodyLower.indexOf(kw) !== -1) {
            return true;
          }
        }
        return false;
      }
      
      if (checkKeywords(["centrum", "peoplehome"])) bank = "Centrum";
      else if (checkKeywords(["crotis", "nivara", "crotisindia"])) bank = "Crotisindia";
      else if (checkKeywords(["mahindra"])) bank = "Mahindra";
      else if (checkKeywords(["growxcd"])) bank = "GrowXCD";
      else if (checkKeywords(["equitas"])) bank = "Equitas";
      else if (checkKeywords(["niwas", "niwas home"])) bank = "Niwas";
      else if (checkKeywords(["truhome"])) bank = "Truhome";
      else if (checkKeywords(["hdfc"])) bank = "HDFC";
      else if (checkKeywords(["icici"])) bank = "ICICI";
      else if (checkKeywords(["axis"])) bank = "Axis";
      else if (checkKeywords(["sbi", "state bank"])) bank = "SBI";
      
      // 2. PARSE REFERENCE NUMBER (Case ID / Lead ID / Application Number)
      var refNo = "";
      if (horizTable) {
        refNo = horizTable["LEAD ID"] || horizTable["LEAD ID:"] || horizTable["LEADID"] ||
                horizTable["LOS NUMBER"] || horizTable["LOS NUMBER:"] ||
                horizTable["APPLICATION NUMBER"] || horizTable["APPLICATION NUMBER:"] ||
                horizTable["CASE ID"] || horizTable["CASE ID:"] ||
                horizTable["TECH ID"] || horizTable["TECH ID:"] || "";
      }
      
      if (!refNo) {
        if (bank === "Centrum") {
          var match = body.match(/Application\s*Number\s*(?::|-|\s|\t)\s*([A-Za-z0-9\/-]+)/i);
          if (match) refNo = match[1].trim();
        } else if (bank === "Equitas") {
          var match = body.match(/Lead\s*Number\s*(?::|-|\s|\t)\s*([A-Za-z0-9\/-]+)/i);
          if (match) refNo = match[1].trim();
          else {
            var match2 = subject.match(/(LOS-\d+)/i) || body.match(/(LOS-\d+)/i);
            if (match2) refNo = match2[1].trim();
          }
        } else if (bank === "Mahindra") {
          var match = subject.match(/(A00000\d+)/i) || body.match(/(A00000\d+)/i);
          if (match) refNo = match[1].trim();
        } else if (bank === "Niwas") {
          var match = subject.match(/(DLVLP[A-Za-z0-9_-]+)/i) || body.match(/(DLVLP[A-Za-z0-9_-]+)/i) ||
                      subject.match(/(DLPON[A-Za-z0-9_-]+)/i) || body.match(/(DLPON[A-Za-z0-9_-]+)/i);
          if (match) refNo = match[1].trim();
        } else if (bank === "Truhome") {
          var match = body.match(/Lead\s*ID\s*(?::|-|\s|\t)\s*([A-Za-z0-9_-]+)/i) ||
                      body.match(/LEAD\s*ID\s*(?::|-|\s|\t)\s*([A-Za-z0-9_-]+)/i);
          if (match) refNo = match[1].trim();
        }
        
        // Subject fallback for Truhome subject containing "LEAD ID"
        if (!refNo && bank === "Truhome") {
          var matchSub = subject.match(/LEAD\s*ID\s*(\d+)/i);
          if (matchSub) refNo = matchSub[1].trim();
        }
        
        // Fallback for general Case ID patterns
        if (!refNo) {
          var match = body.match(/(?:Lead|LOS|Case|Ref|Reference)?\s*(?:No|Number|ID)?\s*(?::|-)\s*([A-Za-z0-9\/-]+)/i);
          if (match) refNo = match[1].trim();
        }
      }
      
      // 3. PARSE CUSTOMER NAME (Owner)
      var owner = "";
      if (horizTable) {
        owner = horizTable["CUSTOMER NAME"] || horizTable["CUSTOMER NAME:"] ||
                horizTable["APPLICANT NAME"] || horizTable["APPLICANT NAME:"] ||
                horizTable["PROPERTY OWNER NAME"] || horizTable["PROPERTY OWNER NAME:"] ||
                horizTable["APPLICANT & CO APP. NAME"] || horizTable["APPLICANT & CO APP. NAME:"] || "";
      }
      
      if (!owner) {
        if (bank === "Centrum") {
          var match = body.match(/Applicant\s*Name\s*(?::|-|\s|\t)\s*([^\n\r]+)/i);
          if (match) owner = match[1].trim();
        } else if (bank === "Equitas") {
          var match = body.match(/Applicant\s*(?:\/|\s*&\s*Co\s*App\.)\s*Name\s*(?::|-|\s|\t)\s*([^\n\r]+)/i) ||
                      body.match(/Applicant\s*\/?\s*Co\s*App\.\s*Name\s*(?::|-|\s|\t)\s*([^\n\r]+)/i);
          if (match) {
            owner = match[1].trim();
          } else {
            // Parse subject slash format: NEED VALUATION REPORT-.../LOCATION/PONDY/Mr.NAME/NAME-LOS-861961
            var cleanSub = subject.replace(/^[Ff][Ww]:\s*/, "").replace(/^[Rr][Ee]:\s*/, "");
            var parts = cleanSub.split("/");
            if (parts.length >= 4) {
              var name1 = parts[3].replace(/Mr\.|Mrs\.|Ms\./ig, "").trim();
              var name2 = "";
              if (parts[4]) {
                var dashIdx = parts[4].indexOf("-");
                name2 = (dashIdx !== -1 ? parts[4].substring(0, dashIdx) : parts[4]).replace(/Mr\.|Mrs\.|Ms\./ig, "").trim();
              }
              owner = name1 + (name2 ? " & " + name2 : "");
            }
          }
        } else if (bank === "Mahindra") {
          var match = body.match(/Customer\s*Name\s*-\s*([^\n\r]+)/i) ||
                      body.match(/Property\s*inspection\s*for\s*([^-]+)-/i) ||
                      subject.match(/Final\s*Stage\s*-\s*([^-]+)\s*-/i);
          if (match) owner = match[1].trim();
        } else if (bank === "Crotisindia") {
          var match = body.match(/Name\s*of\s*the\s*Applicant\s*(?::|-|\s|\t)\s*([^\n\r]+)/i) ||
                      body.match(/Owner\s*of\s*the\s*Property\s*(?::|-|\s|\t)\s*([^\n\r]+)/i) ||
                      body.match(/1\.\s*Customer\s*name\s*(?::|-|\s|\t)\s*([^\n\r]+)/i);
          if (match) owner = match[1].trim();
        } else if (bank === "Niwas") {
          var match = body.match(/Property\s*Owner\s*Name\s*(?:\([^\)]+\))?\s*(?::|-|\s|\t)\s*([^\n\r]+)/i) ||
                      body.match(/Name\s*(?::|-|\s|\t)\s*([^\n\r]+)/i);
          if (match) owner = match[1].trim();
        } else if (bank === "Truhome") {
          var match = body.match(/Customer\s*Name\s*(?::|-|\s|\t)\s*([^\n\r]+)/i) ||
                      body.match(/CUSTOMER\s*NAME\s*(?::|-|\s|\t)\s*([^\n\r\t]+)/i);
          if (match) owner = match[1].trim();
        }
        
        // Subject fallback for Truhome subject containing "CUSTOMER NAME" or "CUSTOEMR NAME"
        if (!owner && bank === "Truhome") {
          var matchSub = subject.match(/(?:CUSTOMER|CUSTOEMR)\s*NAME\s*([A-Za-z\s]+?)\s*(?:LEAD|ID|$)/i);
          if (matchSub) owner = matchSub[1].trim();
        }
        
        // Universal double slash subject fallback (Niwas / Centrum): Customer Name // Mr. X // RefNo // Branch
        if (!owner && subject.indexOf("//") !== -1) {
          var parts = subject.split("//");
          if (parts.length >= 3) {
            for (var p = 0; p < parts.length; p++) {
              var part = parts[p].trim();
              if (part.match(/(?:Mr\.|Mrs\.|Ms\.)/i) || (!part.match(/^(?:PDY|LOS|DLVLP|DLPON|A000)\d+/i) && !part.match(/^[A-Za-z0-9_-]+-\d+$/) && part.indexOf("Branch") === -1 && part.indexOf("Stage") === -1 && part.indexOf("Reg") === -1 && part.length > 3)) {
                owner = part;
                break;
              }
            }
          }
        }
        
        // Global fallback for Owner/Applicant Name
        if (!owner) {
          var match = body.match(/(?:Customer|Applicant|Borrower|Owner)?\s*Name\s*(?::|-|\s|\t)\s*([^\n\r\t]+)/i);
          if (match) owner = match[1].trim();
        }
      }
      
      // Clean up Name salutations
      if (owner) {
        owner = owner.replace(/Mr\.\s*|Mrs\.\s*|Ms\.\s*|Mr\s+|Mrs\s+|Ms\s+/ig, "").replace(/,\s*$/g, "").trim();
      }
      
      // 4. PARSE PROPERTY ADDRESS
      var address = "";
      if (horizTable) {
        address = horizTable["PROPERTY ADDRESS"] || horizTable["PROPERTY ADDRESS:"] ||
                  horizTable["ADDRESS"] || horizTable["ADDRESS:"] || "";
      }
      
      if (!address) {
        var addrMatch = body.match(/Property\s*Address\s*(?:\([^\)]+\))?\s*(?::|-|\s|\t)\s*([\s\S]+?)(?=\r?\n\s*(?:[A-Za-z0-9#\-\.\s_\[\]]+(?::|-|\t)|\r?\n|$))/i) ||
                        body.match(/9\.\s*Property\s*Address\s*(?::|-|\s|\t)\s*([\s\S]+?)(?=\r?\n\s*(?:[A-Za-z0-9#\-\.\s_\[\]]+(?::|-|\t)|\r?\n|$))/i) ||
                        body.match(/Property\s*Address\s*-\s*([^\n\r]+)/i) ||
                        body.match(/PROPERTY\s*ADDRESS\s*(?::|-|\s|\t)\s*([\s\S]+?)(?=\r?\n\s*(?:[A-Za-z0-9#\-\.\s_\[\]]+(?::|-|\t)|\r?\n|$))/i) ||
                        body.match(/Property\s*Address\s*(?::|-|\s|\t)\s*([^\n\r]+)/i);
                        
        if (addrMatch) {
          address = addrMatch[1].trim();
        }
      }
      
      // 5. PARSE AREA / LOCATION (Branch Name)
      var location = "";
      if (horizTable) {
        location = horizTable["BRANCH"] || horizTable["BRANCH:"] ||
                   horizTable["BRANCH NAME"] || horizTable["BRANCH NAME:"] ||
                   horizTable["LOCATION"] || horizTable["LOCATION:"] || "";
      }
      
      if (!location) {
        var locMatch = body.match(/(?:Sourcing\s*|Branch\s*)?Branch\s*(?:Name)?\s*(?::|-|\s|\t)\s*([\w-]+)/i) ||
                       body.match(/Branch\s*Credit\s*Manager,\s*([\w-]+)\s*Branch/i) ||
                       body.match(/([\w-]+)\s*Branch\b/i) ||
                       subject.match(/([\w-]+)\s*Branch\b/i) ||
                       body.match(/Location\s*(?::|-|\s|\t)\s*([\w-]+)/i);
                       
        if (locMatch) {
          location = locMatch[1].replace(/Branch/ig, "").trim();
        }
        
        // Double slash subject Location check (Equitas / Niwas)
        if (!location) {
          var cleanSub = subject.replace(/^[Ff][Ww]:\s*/, "").replace(/^[Rr][Ee]:\s*/, "");
          if (cleanSub.indexOf("//") !== -1) {
            var parts = cleanSub.split("//");
            for (var p = 0; p < parts.length; p++) {
              var part = parts[p].trim();
              if (part.indexOf("Branch") !== -1) {
                location = part.replace(/Branch/ig, "").trim();
                break;
              }
            }
          } else if (bank === "Equitas") {
            var parts = cleanSub.split("/");
            if (parts.length >= 3) {
              location = parts[1].trim();
            }
          }
        }
      }
      
      // Clean location formatting (e.g. "VILLUPURAM-1" -> "Villupuram")
      if (location) {
        location = location.split(" ")[0]; // Take first word
        location = location.replace(/[^a-zA-Z]/g, ""); // Keep only letters
        if (location) {
          location = location.charAt(0).toUpperCase() + location.slice(1).toLowerCase();
        }
      }
      
      // 6. PARSE CONTACT NUMBERS (Up to 3 unique numbers)
      var uniquePhones = [];
      if (horizTable) {
        var rawContact = horizTable["CONTACT NUMBER"] || horizTable["CONTACT NUMBER:"] ||
                         horizTable["CUSTOMER CONTACT NO"] || horizTable["CUSTOMER CONTACT NO:"] ||
                         horizTable["CONTACT PERSON PHONE NO"] || horizTable["CONTACT PERSON PHONE NO:"] ||
                         horizTable["PHONE NUMBER"] || horizTable["PHONE NUMBER:"] || "";
        if (rawContact) {
          var cleaned = rawContact.replace(/(\d)\s+(\d)/g, "$1$2");
          var matches = cleaned.match(/\b\d{10,12}\b/g) || [];
          for (var p = 0; p < matches.length; p++) {
            var num = matches[p].replace(/^0+/, "");
            if (num.length === 12 && num.startsWith("91")) num = num.substring(2);
            if (num.length === 10 && uniquePhones.indexOf(num) === -1) uniquePhones.push(num);
          }
        }
      }
      
      if (uniquePhones.length === 0) {
        var contactLine = "";
        var numMatch = body.match(/Contact\s*Person\s*(?:Name\s*and\s*Number|&\s*Numbers)\s*(?::|-|\s|\t)\s*([^\n\r\t]+)/i) ||
                       body.match(/Contact\s*(?:Person\s*)?Number\s*(?::|-|\s|\t)\s*([^\n\r\t]+)/i) ||
                       body.match(/Customer\s*Contact\s*No\s*(?::|-|\s|\t)\s*([^\n\r\t]+)/i) ||
                       body.match(/4\.\s*Mobile\s*No\s*(?::|-|\s|\t)\s*([^\n\r\t]+)/i) ||
                       body.match(/Phone\s*number\s*-\s*([^\n\r]+)/i) ||
                       body.match(/Mobile\s*(?::|-|\s|\t)\s*([^\n\r\t]+)/i) ||
                       body.match(/CONTACT\s*NUMBER\s*(?::|-|\s|\t)\s*([^\n\r\t]+)/i) ||
                       body.match(/(?:contact|phone|mobile|tel)?\s*(?:no|number)?\s*(?::|-)\s*([^\n\r]+)/i);
        
        if (numMatch) {
          contactLine = numMatch[1].trim();
        }
        
        var phoneNumbers = [];
        if (contactLine) {
          var cleanedLine = contactLine.replace(/(\d)\s+(\d)/g, "$1$2");
          var digitsMatches = cleanedLine.match(/\b\d{10,12}\b/g) || [];
          for (var p = 0; p < digitsMatches.length; p++) {
            var num = digitsMatches[p].replace(/^0+/, "");
            if (num.length === 12 && num.startsWith("91")) {
              num = num.substring(2);
            }
            if (num.length === 10) {
              phoneNumbers.push(num);
            }
          }
        }
        
        // Fallback: search body text (excluding signature) for any 10-digit number
        if (phoneNumbers.length === 0) {
          var bodyForPhones = body;
          var regardsIndex = body.search(/(?:Regards|Thanks|Credit Admin|Credit Manager|Branch Credit|Credit Dept)/i);
          if (regardsIndex !== -1) {
            bodyForPhones = body.substring(0, regardsIndex);
          }
          var rawMatches = bodyForPhones.match(/\b\d{10}\b/g) || [];
          phoneNumbers = rawMatches;
        }
        
        for (var u = 0; u < phoneNumbers.length; u++) {
          if (uniquePhones.indexOf(phoneNumbers[u]) === -1) {
            uniquePhones.push(phoneNumbers[u]);
          }
        }
      }
      
      var contact1 = uniquePhones[0] || "";
      var contact2 = uniquePhones[1] || "";
      var contact3 = uniquePhones[2] || "";
      
      // 7. CHECK FOR DUPLICATES (By Owner or by RefNo)
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
        
        // 8. GENERATE AUTO CASE ROW
        var fileId = generateUUID();
        var fileNo = "VT-" + yyyy + "-" + mm + "-" + dd + "-" + nextSerial;
        nextSerial++;
        
        // Create full row array matching sheet headers
        var newRow = [];
        for (var c = 0; c < headers.length; c++) {
          var header = headers[c];
          if (header === "ID") newRow.push(fileId);
          else if (header === "FileNo") newRow.push(fileNo);
          else if (header === "Owner") newRow.push(owner);
          else if (header === "Bank") newRow.push(bank);
          else if (header === "Address") newRow.push(address);
          else if (header === "Location") newRow.push(location);
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
          else if (header === "Contact") newRow.push(contact1);
          else if (header === "Contact2") newRow.push(contact2);
          else if (header === "Contact3") newRow.push(contact3);
          else newRow.push("");
        }
        
        // Save to spreadsheet
        sheet.appendRow(newRow);
        Logger.log("Successfully imported case from Gmail: " + owner);
        
        // 9. SEND WHATSAPP NOTIFICATION
        var admin = getAdminContact();
        if (admin && admin.phone && admin.apikey) {
          var text = "🔔 *ValuTrack Auto-Import*\n\n" +
                     "A new case has been captured from your Gmail inbox:\n\n" +
                     "▪️ *Case No:* " + fileNo + "\n" +
                     "▪️ *Owner:* " + owner + "\n" +
                     "▪️ *Bank:* " + bank + "\n" +
                     "▪️ *Area:* " + (location || "Not specified") + "\n" +
                     "▪️ *Contact:* " + (contact1 || "Not specified") + "\n\n" +
                     "Open dashboard: https://sathzsri.github.io/Valutrack/";
          sendCallMeBotMessage(admin.phone, admin.apikey, text);
        }
      }
      
      // Mark as processed
      msg.markRead();
    }
  }
}

// Helper to parse HTML tables directly from raw Gmail HTML
function parseHtmlTable(htmlText) {
  if (!htmlText) return null;
  // Look for all table rows <tr>...</tr>
  var rowRegex = /<tr[^>]*>([\s\S]+?)<\/tr>/gi;
  var cellRegex = /<t[dh][^>]*>([\s\S]+?)<\/t[dh]>/gi;
  var rows = [];
  var match;
  
  while ((match = rowRegex.exec(htmlText)) !== null) {
    var rowHtml = match[1];
    var cells = [];
    var cellMatch;
    while ((cellMatch = cellRegex.exec(rowHtml)) !== null) {
      var cellText = cellMatch[1].replace(/<[^>]+>/g, "")
                                  .replace(/&nbsp;/gi, " ")
                                  .replace(/&amp;/gi, "&")
                                  .replace(/&lt;/gi, "<")
                                  .replace(/&gt;/gi, ">")
                                  .trim();
      cells.push(cellText);
    }
    if (cells.length > 0) {
      rows.push(cells);
    }
  }
  
  // Find the row containing "CUSTOMER NAME" or "APPLICANT NAME" or "LEAD ID"
  for (var r = 0; r < rows.length; r++) {
    var row = rows[r];
    var upperRow = row.map(function(c) { return c.toUpperCase(); });
    var customerIdx = -1;
    for (var c = 0; c < upperRow.length; c++) {
      if (upperRow[c].indexOf("CUSTOMER NAME") !== -1 || upperRow[c].indexOf("APPLICANT NAME") !== -1 || upperRow[c].indexOf("LEAD ID") !== -1) {
        customerIdx = c;
        break;
      }
    }
    
    if (customerIdx !== -1) {
      // Find the first row after this one that has cells
      var valueRow = rows[r + 1];
      if (valueRow) {
        var data = {};
        for (var col = 0; col < row.length; col++) {
          var headerName = row[col].toUpperCase().trim();
          var val = valueRow[col] ? valueRow[col].trim() : "";
          data[headerName] = val;
        }
        return data;
      }
    }
  }
  return null;
}

// Helper to parse horizontal tables in Gmail body text
function parseHorizontalTable(bodyText) {
  var lines = bodyText.split("\n");
  var headersLine = -1;
  var headers = [];
  
  for (var l = 0; l < lines.length; l++) {
    var line = lines[l].toUpperCase();
    if (line.indexOf("CUSTOMER NAME") !== -1 && (line.indexOf("PROPERTY ADDRESS") !== -1 || line.indexOf("LEAD ID") !== -1)) {
      headersLine = l;
      var cleanedHeaders = lines[l].trim().replace(/\t+/g, "|").replace(/\s{2,}/g, "|");
      headers = cleanedHeaders.split("|");
      break;
    }
  }
  
  if (headersLine !== -1) {
    var valuesLine = -1;
    for (var l = headersLine + 1; l < lines.length; l++) {
      if (lines[l].trim().length > 0) {
        valuesLine = l;
        break;
      }
    }
    
    if (valuesLine !== -1) {
      var cleanedValues = lines[valuesLine].trim().replace(/\t+/g, "|").replace(/\s{2,}/g, "|");
      var values = cleanedValues.split("|");
      
      var data = {};
      for (var h = 0; h < headers.length; h++) {
        var header = headers[h].trim().toUpperCase();
        var value = values[h] ? values[h].trim() : "";
        data[header] = value;
      }
      return data;
    }
  }
  return null;
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

function deduplicateSheetRows(sheet) {
  var values = sheet.getDataRange().getDisplayValues();
  var headers = values[0];
  var idColIndex = headers.indexOf("ID");
  var historyColIndex = headers.indexOf("History");
  if (idColIndex === -1) return;
  
  var idMap = {};
  var rowsToDelete = [];
  
  for (var i = 1; i < values.length; i++) {
    var id = values[i][idColIndex];
    if (!id) continue;
    
    var historyStr = historyColIndex !== -1 ? values[i][historyColIndex] : "";
    var historyLen = 0;
    try {
      historyLen = historyStr ? JSON.parse(historyStr).length : 0;
    } catch (e) {
      historyLen = 0;
    }
    
    var rowIndex = i + 1;
    
    if (idMap[id]) {
      var existing = idMap[id];
      if (historyLen > existing.historyLen) {
        rowsToDelete.push(existing.rowIndex);
        idMap[id] = { rowIndex: rowIndex, historyLen: historyLen };
      } else {
        rowsToDelete.push(rowIndex);
      }
    } else {
      idMap[id] = { rowIndex: rowIndex, historyLen: historyLen };
    }
  }
  
  if (rowsToDelete.length > 0) {
    rowsToDelete.sort(function(a, b) { return b - a; });
    for (var k = 0; k < rowsToDelete.length; k++) {
      sheet.deleteRow(rowsToDelete[k]);
    }
  }
}
