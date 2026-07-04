// Google Apps Script - 雲端完整讀寫
// 所有資料存在 Google Sheets，前端可新增/編輯/刪除

var TIMEZONE = "Asia/Taipei";

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: getAllData()
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var reqData = JSON.parse(e.postData.contents);
    var action = reqData.action;
    
    if (action === "addRecords") {
      return addRecordsOnly(reqData.records);
    } else if (action === "deleteRecord") {
      return deleteRecordById(reqData.recordId);
    } else if (action === "updateRecord") {
      return updateRecordById(reqData.recordId, reqData.record);
    } else if (action === "updateAccounts") {
      return updateAccountsOnly(reqData.accounts);
    } else if (action === "updateBudgets") {
      return updateBudgetsAndGroups(reqData.budgets, reqData.customGroupBudgets);
    } else if (action === "clearAllRecords") {
      return clearAllRecordsAction();
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Unknown action" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    Logger.log("Error: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// 新增記錄（追加，保留舊資料）
// ============================================
function addRecordsOnly(newRecords) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRec = ss.getSheetByName("records") || ss.insertSheet("records");
  
  // 如果工作表為空，先加標題列
  if (sheetRec.getLastRow() === 0) {
    sheetRec.getRange(1, 1, 1, 10).setValues([["id", "date", "type", "amount", "note", "account", "mainCategory", "subCategory", "fromAccount", "toAccount"]]);
  }
  
  var rows = [];
  newRecords.forEach(function(r) {
    var dateStr = r.date || "";
    if (dateStr.includes("T")) dateStr = dateStr.slice(0, 10);
    
    rows.push([
      r.id || "",
      dateStr,
      r.type || "",
      r.amount || 0,
      r.note || "",
      r.account || "",
      r.mainCategory || "",
      r.subCategory || "",
      r.fromAccount || "",
      r.toAccount || ""
    ]);
  });
  
  if (rows.length > 0) {
    var startRow = sheetRec.getLastRow() + 1;
    sheetRec.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: getAllData() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// 刪除單筆記錄
// ============================================
function deleteRecordById(recordId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRec = ss.getSheetByName("records");
  
  if (!sheetRec || sheetRec.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "records sheet not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var vals = sheetRec.getDataRange().getValues();
  var rowToDelete = -1;
  
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] == recordId) {
      rowToDelete = i + 1;
      break;
    }
  }
  
  if (rowToDelete > 0) {
    sheetRec.deleteRow(rowToDelete);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: getAllData() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// 編輯單筆記錄
// ============================================
function updateRecordById(recordId, updatedRecord) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRec = ss.getSheetByName("records");
  
  if (!sheetRec || sheetRec.getLastRow() <= 1) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "records sheet not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var vals = sheetRec.getDataRange().getValues();
  var rowToUpdate = -1;
  
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] == recordId) {
      rowToUpdate = i + 1;
      break;
    }
  }
  
  if (rowToUpdate > 0) {
    var dateStr = updatedRecord.date || "";
    if (dateStr.includes("T")) dateStr = dateStr.slice(0, 10);
    
    var newRow = [
      recordId,
      dateStr,
      updatedRecord.type || "",
      updatedRecord.amount || 0,
      updatedRecord.note || "",
      updatedRecord.account || "",
      updatedRecord.mainCategory || "",
      updatedRecord.subCategory || "",
      updatedRecord.fromAccount || "",
      updatedRecord.toAccount || ""
    ];
    
    sheetRec.getRange(rowToUpdate, 1, 1, newRow.length).setValues([newRow]);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: getAllData() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// 更新帳戶（替換整個帳戶表）
// ============================================
function updateAccountsOnly(accountsData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetAcc = ss.getSheetByName("accounts") || ss.insertSheet("accounts");
  sheetAcc.clear();
  
  if (accountsData && accountsData.length > 0) {
    sheetAcc.getRange(1, 1, accountsData.length, accountsData[0].length).setValues(accountsData);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: getAllData() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// 更新預算（替換預算表）
// ============================================
function updateBudgetsAndGroups(budgetsData, customGroupBudgetsData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetBud = ss.getSheetByName("budgets") || ss.insertSheet("budgets");
  sheetBud.clear();
  if (budgetsData && budgetsData.length > 0) {
    sheetBud.getRange(1, 1, budgetsData.length, budgetsData[0].length).setValues(budgetsData);
  }
  
  var sheetCustBud = ss.getSheetByName("custom_budgets") || ss.insertSheet("custom_budgets");
  sheetCustBud.clear();
  if (customGroupBudgetsData) {
    sheetCustBud.getRange(1, 1).setValue(customGroupBudgetsData);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: getAllData() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// 清空所有記錄
// ============================================
function clearAllRecordsAction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetRec = ss.getSheetByName("records");
  
  if (sheetRec) {
    sheetRec.clear();
    sheetRec.getRange(1, 1, 1, 10).setValues([["id", "date", "type", "amount", "note", "account", "mainCategory", "subCategory", "fromAccount", "toAccount"]]);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", data: getAllData() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// 核心讀取函式
// ============================================
function getAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = { accounts: {}, budgets: {}, customGroupBudgets: "{}", records: [] };
  
  // 讀取帳戶
  var sheetAcc = ss.getSheetByName("accounts");
  if (sheetAcc && sheetAcc.getLastRow() > 0) {
    var vals = sheetAcc.getDataRange().getValues();
    vals.forEach(function(row) { 
      if(row[0]) res.accounts[row[0]] = Number(row[1]) || 0; 
    });
  }
  
  // 讀取單一預算
  var sheetBud = ss.getSheetByName("budgets");
  if (sheetBud && sheetBud.getLastRow() > 0) {
    var vals = sheetBud.getDataRange().getValues();
    vals.forEach(function(row) { 
      if(row[0]) res.budgets[row[0]] = Number(row[1]) || 0; 
    });
  }
  
  // 讀取自訂群組預算
  var sheetCustBud = ss.getSheetByName("custom_budgets");
  if (sheetCustBud && sheetCustBud.getLastRow() >= 1) {
    res.customGroupBudgets = sheetCustBud.getRange(1, 1).getValue().toString();
  }
  
  // 讀取歷史紀錄
  var sheetRec = ss.getSheetByName("records");
  if (sheetRec && sheetRec.getLastRow() > 1) {
    var vals = sheetRec.getDataRange().getValues();
    var headers = vals[0];
    
    for (var i = 1; i < vals.length; i++) {
      var row = vals[i];
      var r = {};
      
      headers.forEach(function(h, idx) { 
        var val = row[idx];
        
        if (h === "date") {
          if (typeof val === "string") {
            r[h] = val;
          } else if (val instanceof Date) {
            r[h] = Utilities.formatDate(val, TIMEZONE, "yyyy-MM-dd");
          } else {
            r[h] = "";
          }
        } else {
          r[h] = val;
        }
      });
      
      r.amount = Number(r.amount) || 0;
      
      if (r.date) {
        res.records.push(r);
      }
    }
  }
  
  return res;
}
