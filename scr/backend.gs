// ==========================================
// CONFIGURATION & CONSTANTS
// ==========================================
var DB_SHEET_NAME = "Data Base"; 
var LOG_SHEET_NAME = "Log";
var ADMIN_PASSWORD = "1234"; // รหัสผ่านสำหรับ Admin

var CAP_SPREADSHEET_ID = "1PYcAatoJ4QX28uQ_LF8dDC6oTiMWbfPs5TZDfGJVa4U";
var CAP_SHEET_NAME = "Cap";
var PLAN_SHEET_NAME = "Plan";

// Snapshot ยอดสแกนราย Job Order: เก็บผลนับไว้พร้อมเลขแถวสุดท้ายที่นับไปแล้ว
// รอบถัดไปจึงอ่านเฉพาะแถวใหม่ (ไม่ต้องอ่าน Log ทั้งชีททุกครั้งที่เครื่องสแกน poll เข้ามา)
// แต่ยอดที่ได้ยังเป็นข้อมูลล่าสุดเสมอ — สำคัญ เพราะฝั่งหน้าจอใช้ยืนยันว่ายอดที่เพิ่งส่งขึ้น Sheet
// ถูกนับรวมแล้วหรือยัง
var JOB_COUNT_SNAPSHOT_KEY = "JOB_SCAN_SNAPSHOT_V1";
var JOB_COUNT_SNAPSHOT_TTL = 21600; // 6 ชม.
var JOB_COUNT_FULL_RECOUNT_MS = 10 * 60 * 1000; // นับใหม่ทั้งชีททุก 10 นาที เผื่อมีคนแก้ Log ย้อนหลัง

// ยอด "Actual Scan" จากชีท Plan (แหล่งของหลังบ้าน) อ่านใหม่ทุก 1 นาทีก็พอ เพราะเป็นยอดสรุป
var PLAN_SCAN_CACHE_KEY = "PLAN_ACTUAL_SCAN_V1";
var PLAN_SCAN_CACHE_SEC = 60;

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

// 4a. ยอดสแกนสะสมของแต่ละ Job Order จาก Log (รวมทุกเครื่อง/ทุกวัน, ไม่นับแถวที่ VOID)
// ใช้เป็นค่าหลักในการแสดง "Scanned / คงเหลือ" ให้พนักงานเห็นตรงกันทุก Line
function getJobScanCounts() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sheet) return JSON.stringify({ ok: false, message: "Sheet '" + LOG_SHEET_NAME + "' Not Found" });

    var lastRow = sheet.getLastRow();
    var cache = null;
    var snap = null;
    try {
      cache = CacheService.getScriptCache();
      var raw = cache.get(JOB_COUNT_SNAPSHOT_KEY);
      if (raw) snap = JSON.parse(raw);
    } catch (e) {
      cache = null;
      snap = null;
    }

    var now = new Date().getTime();
    var snapUsable = snap && snap.counts && typeof snap.lastRow === "number" &&
                     snap.lastRow >= 1 && snap.lastRow <= lastRow &&
                     snap.fullAt && (now - snap.fullAt) <= JOB_COUNT_FULL_RECOUNT_MS;

    // นับใหม่ทั้งชีทเมื่อ: ยังไม่มี snapshot / snapshot หมดอายุ / มีการลบแถวใน Log
    if (!snapUsable) snap = { counts: {}, lastRow: 1, fullAt: now };

    if (lastRow > snap.lastRow) {
      // อ่านเฉพาะแถวใหม่ คอลัมน์ B..E (Job, Model, Barcode, Status)
      var data = sheet.getRange(snap.lastRow + 1, 2, lastRow - snap.lastRow, 4).getValues();
      for (var i = 0; i < data.length; i++) {
        var job = String(data[i][0]).trim();
        if (job === "") continue;
        if (String(data[i][3]).trim().toUpperCase() === "VOID") continue;
        snap.counts[job] = (snap.counts[job] || 0) + 1;
      }
      snap.lastRow = lastRow;
    }

    if (cache) {
      try { cache.put(JOB_COUNT_SNAPSHOT_KEY, JSON.stringify(snap), JOB_COUNT_SNAPSHOT_TTL); } catch (e) {}
    }
    return JSON.stringify({ ok: true, counts: snap.counts, planCounts: getPlanActualScanCounts_() });
  } catch (e) {
    // ห้ามส่ง counts ว่างเมื่อ error เพราะฝั่งหน้าจอจะเข้าใจผิดว่ายอดสแกนเป็น 0
    return JSON.stringify({ ok: false, message: e.message });
  }
}

// ยอด "Actual Scan" ที่หลังบ้านคุมไว้ในชีท Plan (คนละแหล่งกับการนับแถวใน Log)
// หาคอลัมน์จากชื่อหัวตาราง ไม่ยึดตำแหน่งคอลัมน์ เผื่อมีการแทรก/ย้ายคอลัมน์ในอนาคต
function getPlanActualScanCounts_() {
  var cache = null;
  try {
    cache = CacheService.getScriptCache();
    var raw = cache.get(PLAN_SCAN_CACHE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {
    cache = null;
  }

  var counts = {};
  try {
    var sheet = SpreadsheetApp.openById(CAP_SPREADSHEET_ID).getSheetByName(PLAN_SHEET_NAME);
    if (!sheet) return counts;

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return counts;

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var scanColIndex = -1;
    headers.forEach(function(header, index) {
      if (String(header).trim().toLowerCase() === "actual scan") scanColIndex = index;
    });
    if (scanColIndex < 0) return counts; // ไม่มีคอลัมน์นี้ ก็ใช้ยอดจาก Log อย่างเดียว

    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    for (var i = 0; i < data.length; i++) {
      var job = String(data[i][3]).trim(); // Job Order = Col D เหมือน getActiveJobOrders
      if (job === "") continue;
      var n = toNonNegativeInt_(data[i][scanColIndex]);
      if (n !== null) counts[job] = n; // Job ซ้ำ ใช้แถวล่างสุดเป็นค่าล่าสุด
    }
  } catch (e) {
    return {}; // อ่านชีท Plan ไม่ได้ ไม่ใช่เหตุให้ยอดจาก Log ใช้ไม่ได้
  }

  if (cache) {
    try { cache.put(PLAN_SCAN_CACHE_KEY, JSON.stringify(counts), PLAN_SCAN_CACHE_SEC); } catch (e) {}
  }
  return counts;
}

// รับค่าที่อาจเป็นตัวเลข, ข้อความ "1,186" หรือช่องว่าง/สูตร error -> คืน null เมื่อใช้ไม่ได้
function toNonNegativeInt_(value) {
  if (value === "" || value === null || value === undefined) return null;
  var n = Number(String(value).replace(/,/g, "").trim());
  if (isNaN(n) || n < 0) return null;
  return Math.round(n);
}

// Void เป็นการแก้แถวเก่า ซึ่งการนับแบบเพิ่มทีละแถวใหม่จะมองไม่เห็น จึงต้องล้าง snapshot
function clearJobScanCountCache_() {
  try { CacheService.getScriptCache().remove(JOB_COUNT_SNAPSHOT_KEY); } catch (e) {}
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
    clearJobScanCountCache_(); // ให้ทุกเครื่องเห็นยอดใหม่ทันที ไม่ต้องรอ cache หมดอายุ
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

// ดึงข้อมูลการผลิตวันนี้
function getTodayProductionData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!sheet) return JSON.stringify([]);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  var data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var now = new Date();
  var todayDay   = now.getDate();
  var todayMonth = now.getMonth() + 1;
  var todayYear  = now.getFullYear() + 543; // ปี พ.ศ.
  var todayData  = [];

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
      todayData.push({ model: data[i][2], hour: rowHour });
    }
  }
  return JSON.stringify(todayData);
}
