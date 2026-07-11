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

// 3. ดึง Job Order ที่ยังไม่มี Actual Complete Date จาก Plan
function getActiveJobOrders() {
  try {
    var ss = SpreadsheetApp.openById(CAP_SPREADSHEET_ID);
    var sheet = ss.getSheetByName(PLAN_SHEET_NAME);
    
    if (!sheet) return [];
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    // Job Order = Col D (Index 3), Model = Col G (Index 6), Incomplete = header "Actual complete date" is blank
    var lastCol = Math.max(sheet.getLastColumn(), 11);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var actualCompleteDateColIndex = -1; // -1 = not found

    headers.forEach(function(header, index) {
      if (String(header).trim().toLowerCase() === "actual complete date") {
        actualCompleteDateColIndex = index;
      }
    });

    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var activeJobs = [];

    data.forEach(function(row) {
      var jobOrder = String(row[3]).trim();
      var orderModel = String(row[6]).trim();
      var planQty = row[7] || 0;

      if (jobOrder === "") return;

      if (actualCompleteDateColIndex >= 0) {
        var actualCompleteDate = String(row[actualCompleteDateColIndex]).trim().toLowerCase();
        if (actualCompleteDate !== "" && actualCompleteDate !== "incomplete") return;
      }

      activeJobs.push({ job: jobOrder, model: orderModel, qty: planQty });
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

// 4b. แก้ไขยอดเกิน: Void รายการสแกนล่าสุดของ Job Order ที่ระบุ (ไม่ลบแถวจริง เพื่อให้ตรวจสอบย้อนหลังได้)
function voidLastJobScans(job, count, passwordInput) {
  if (passwordInput !== ADMIN_PASSWORD) {
    return { success: false, message: "Incorrect Password!" };
  }
  if (!job || !count || count <= 0) {
    return { success: false, message: "Invalid job or count" };
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) return { success: false, message: "Sheet '" + LOG_SHEET_NAME + "' Not Found" };

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return { success: false, message: "Server Busy (Try again)" };
    }

    var lastRow = logSheet.getLastRow();
    if (lastRow < 2) {
      lock.releaseLock();
      return { success: false, message: "No data" };
    }

    var data = logSheet.getRange(2, 1, lastRow - 1, 5).getValues(); // Timestamp, Job, Model, Barcode, Status
    var voided = 0;

    for (var i = data.length - 1; i >= 0 && voided < count; i--) {
      var rowJob = String(data[i][1]).trim();
      var rowStatus = String(data[i][4]).trim();
      if (rowJob === String(job).trim() && rowStatus !== "VOID") {
        var sheetRow = i + 2; // +2: header row + 1-indexed range
        logSheet.getRange(sheetRow, 5).setValue("VOID");
        voided++;
      }
    }

    SpreadsheetApp.flush();
    lock.releaseLock();
    return { success: true, message: "Voided " + voided + " record(s) for job " + job, voided: voided };
  } catch (e) {
    return { success: false, message: "Error: " + e.message };
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

// อ่านแถว Log ของวันนี้ (ไม่รวม VOID) — คืน [{job, model, hour, station}]
function readTodayLogRows_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Log เก่าอาจมีคอลัมน์น้อยกว่า 6 — ห้ามขอ Range เกินขนาดจริงของ Sheet ไม่งั้น throw
  var numCols = Math.min(sheet.getLastColumn(), 6);
  if (numCols < 5) return []; // ไม่มีคอลัมน์ Status = ข้อมูลไม่อยู่ในรูปแบบที่ใช้ได้
  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var now = new Date();
  var todayDay   = now.getDate();
  var todayMonth = now.getMonth() + 1;
  var todayYear  = now.getFullYear() + 543; // ปี พ.ศ.
  var rows = [];

  for (var i = 0; i < data.length; i++) {
    var rowDay, rowMonth, rowYear, rowHour;

    if (data[i][0] instanceof Date) {
      var d = data[i][0];
      rowDay   = d.getDate();
      rowMonth = d.getMonth() + 1;
      rowYear  = d.getFullYear() + 543;
      rowHour  = d.getHours().toString().padStart(2, '0');
    } else {
      var str      = String(data[i][0]);
      var datePart = str.split(" ")[0]; // "14/3/2569"
      var timePart = str.split(" ")[1]; // "8:19:39"
      var dp = datePart.split("/");
      rowDay   = parseInt(dp[0]);
      rowMonth = parseInt(dp[1]);
      rowYear  = parseInt(dp[2]);
      rowHour  = (timePart ? timePart.split(":")[0] : "0").padStart(2, '0');
    }

    if (rowDay === todayDay && rowMonth === todayMonth && rowYear === todayYear && String(data[i][4]).trim() !== "VOID") {
      rows.push({
        job: String(data[i][1]).trim(),
        model: data[i][2],
        hour: rowHour,
        station: numCols >= 6 ? String(data[i][5]).trim() : ""
      });
    }
  }
  return rows;
}

// ดึงข้อมูลการผลิตวันนี้ (ทุก Line — ใช้โดย Dashboard)
function getTodayProductionData() {
  var todayData = readTodayLogRows_().map(function(r) {
    return { model: r.model, hour: r.hour };
  });
  return JSON.stringify(todayData);
}

// ดึงข้อมูลการผลิตวันนี้เฉพาะ Line ที่ระบุ — ใช้กู้ยอดหน้าจอหลังรีเฟรช
// (localStorage ของ Apps Script web app หายได้เมื่อ deploy ใหม่/origin เปลี่ยน จึงต้องสร้างยอดใหม่จาก Sheet)
function getTodayStationData(station) {
  var target = String(station || "").trim();
  if (!target) return JSON.stringify([]);

  var todayData = [];
  readTodayLogRows_().forEach(function(r) {
    if (r.station === target) {
      todayData.push({ model: r.model, hour: r.hour, job: r.job });
    }
  });
  return JSON.stringify(todayData);
}
