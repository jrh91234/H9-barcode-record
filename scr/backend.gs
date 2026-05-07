// ==========================================
// CONFIGURATION & CONSTANTS
// ==========================================
var DB_SHEET_NAME = "Data Base"; 
var LOG_SHEET_NAME = "Log";
var ADMIN_PASSWORD = "1234"; // รหัสผ่านสำหรับ Admin

var CAP_SPREADSHEET_ID = "1PYcAatoJ4QX28uQ_LF8dDC6oTiMWbfPs5TZDfGJVa4U";
var CAP_SHEET_NAME = "Cap";
var PLAN_SHEET_NAME = "Plan";

// ==========================================
// WEB APP SERVING
// ==========================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Scanner v17.2 Auto Check')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// CORE FUNCTIONS
// ==========================================

// 1. ดึงข้อมูล Database Offline
function getOfflineDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!dbSheet) return [];
  var lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return [];
  return dbSheet.getRange(2, 1, lastRow - 1, 2).getValues();
}

// 2. ดึงข้อมูล Capacity (Hourly/Daily Cap)
function getCapDatabase() {
  try {
    var ss = SpreadsheetApp.openById(CAP_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CAP_SHEET_NAME);
    if (!sheet) return {}; 
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 5) return {}; 
    
    var data = sheet.getRange(5, 1, lastRow - 4, 8).getValues();
    var capMap = {};
    data.forEach(function(row) {
      var modelName = String(row[0]).trim();
      var hourlyCap = row[6];
      var dailyCap = row[7];
      
      if (modelName) {
         capMap[modelName] = { 
           hourly: (hourlyCap && !isNaN(hourlyCap)) ? hourlyCap : 0,
           daily: (dailyCap && !isNaN(dailyCap)) ? dailyCap : 0
         };
      }
    });
    return capMap;
  } catch (e) {
    return {}; 
  }
}

// 3. ดึง Job Order ที่ยังไม่เสร็จจาก Plan
function getActiveJobOrders() {
  try {
    var ss = SpreadsheetApp.openById(CAP_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(PLAN_SHEET_NAME);
    
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Job Order = Col D (Index 3), Model = Col G (Index 6), Progress = Col K (Index 10)
    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    var activeJobs = [];
    
    data.forEach(function(row) {
      var jobOrder = String(row[3]).trim(); 
      var orderModel = String(row[6]).trim();
      var progress = row[10];              
      
      if (jobOrder !== "" && (progress === "" || parseFloat(progress) < 100)) {
        activeJobs.push({ job: jobOrder, model: orderModel });
      }
    });
    
    return activeJobs;
  } catch (e) {
    return [{ error: e.message }];
  }
}

// 4. บันทึกข้อมูลลง Log
function saveBatchData(jsonString) {
  try {
    var dataArray = JSON.parse(jsonString);
    if (!dataArray || dataArray.length === 0) return "Empty";
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) return "Error: Sheet '" + LOG_SHEET_NAME + "' Not Found";
    
    var lock = LockService.getScriptLock();
    if (lock.tryLock(10000)) {
       var lastRow = logSheet.getLastRow();
       logSheet.getRange(lastRow + 1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
       SpreadsheetApp.flush();
       lock.releaseLock();
       return "Saved " + dataArray.length;
    } else {
       return "Error: Server Busy (Try again)";
    }
  } catch (e) {
    return "Error: " + e.message;
  }
}

// ==========================================
// MODEL STATE MANAGEMENT (ส่วนที่แก้ไขเพิ่ม)
// ==========================================

function getServerModel() { 
  return PropertiesService.getScriptProperties().getProperty("CURRENT_MODEL") || ""; 
}

// ฟังก์ชันสำหรับเปลี่ยน Model แบบปกติ
function setServerModel(model) { 
  PropertiesService.getScriptProperties().setProperty("CURRENT_MODEL", model); 
  return "OK"; 
}

// *** ฟังก์ชันใหม่: บังคับเปลี่ยน Model โดยใช้รหัสผ่าน ***
function forceChangeModel(newModel, passwordInput) {
  if (passwordInput === ADMIN_PASSWORD) {
    PropertiesService.getScriptProperties().setProperty("CURRENT_MODEL", newModel);
    return { success: true, message: "Model changed to " + newModel };
  } else {
    return { success: false, message: "Incorrect Password!" };
  }
}

// ดึงข้อมูลการผลิตวันนี้
function getTodayProductionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Data"); // ตรวจสอบชื่อ Sheet ให้ตรงกับที่มีจริง
  
  if (!sheet) return JSON.stringify([]);
  
  var data = sheet.getDataRange().getValues();
  var todayStr = new Date().toLocaleDateString("th-TH"); 
  var todayData = [];

  for (var i = 1; i < data.length; i++) {
    var rowDate = "";
    if (data[i][0] instanceof Date) {
      rowDate = data[i][0].toLocaleDateString("th-TH");
    } else {
      rowDate = String(data[i][0]).split(" ")[0]; 
    }

    if (rowDate === todayStr) {
      todayData.push({
        model: data[i][2], 
        timestamp: data[i][0] 
      });
    }
  }
  return JSON.stringify(todayData);
}
