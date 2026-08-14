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
      .setTitle('Scanner v17.19 Shift Name A/B')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// CORE FUNCTIONS
// ==========================================

// 1. ดึงข้อมูล Database Offline (Model, Barcode, Multiply)
function getOfflineDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!dbSheet) return [];
  var lastRow = dbSheet.getLastRow();
  if (lastRow < 2) return [];
  return dbSheet.getRange(2, 1, lastRow - 1, 3).getValues();
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

// ค่าที่อนุญาตสำหรับคอลัมน์ Shift เพื่อป้องกันข้อมูลสะกดไม่ตรงกันใน Log
function normalizeShiftName_(value) {
  var match = String(value || "").trim().toUpperCase().match(/^(?:SHIFT\s*)?([AB])$/);
  return match ? "Shift " + match[1] : "";
}

// ค้นหาคอลัมน์ Shift จากหัวตาราง (คืนค่าเป็น index แบบเริ่มที่ 0)
function getLogShiftColumn_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    var header = String(headers[i] || "").trim().toLowerCase().replace(/\s+/g, " ");
    if (header === "shift" || header === "shift name") return i;
  }
  return -1;
}

// เพิ่มหัวตาราง Shift ให้กับ Log เดิมโดยไม่กระทบคอลัมน์เก่า
function ensureLogShiftColumn_(sheet) {
  var existingColumn = getLogShiftColumn_(sheet);
  if (existingColumn >= 0) return existingColumn + 1; // 1-based column number

  var shiftColumn = Math.max(sheet.getLastColumn() + 1, 7);
  sheet.getRange(1, shiftColumn).setValue("Shift");
  return shiftColumn;
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
       var shiftColumn = ensureLogShiftColumn_(logSheet); // 1-based column number
       var writeWidth = Math.max(shiftColumn, 6);
       var rowsToWrite = dataArray.map(function(row) {
         var output = [];
         for (var col = 0; col < writeWidth; col++) output.push("");

         // คอลัมน์เดิม A:F ยังคงรูปแบบเดิม: Timestamp, Job, Model, Barcode, Status, Station
         for (var baseCol = 0; baseCol < Math.min(row.length, 6); baseCol++) {
           output[baseCol] = row[baseCol];
         }

         // Client รุ่นใหม่ส่ง Shift มาในสมาชิกตัวที่ 7 (index 6)
         output[shiftColumn - 1] = normalizeShiftName_(row.length > 6 ? row[6] : "");
         return output;
       });

       var lastRow = logSheet.getLastRow();
       logSheet.getRange(lastRow + 1, 1, rowsToWrite.length, writeWidth).setValues(rowsToWrite);
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

// แปลงปี ค.ศ. หรือ พ.ศ. ให้เป็นปี พ.ศ. เดียวกัน
// Google Sheets อาจส่งค่าวันที่เป็น Date object (ค.ศ.) หรือข้อความ (พ.ศ.)
function normalizeThaiYear_(year) {
  var n = parseInt(year, 10);
  if (isNaN(n)) return NaN;
  return n >= 2400 ? n : n + 543;
}

// แปลงปีที่รับเข้ามาให้เป็น ค.ศ. สำหรับสร้างคีย์วันที่แบบไม่ขึ้นกับ timezone
function toGregorianYear_(year) {
  var n = parseInt(year, 10);
  if (isNaN(n)) return NaN;
  return n >= 2400 ? n - 543 : n;
}

// แยก Timestamp จาก Google Sheets ทั้งแบบ Date object และข้อความ
function parseLogTimestamp_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return {
      day: value.getDate(),
      month: value.getMonth() + 1,
      year: value.getFullYear(),
      hour: value.getHours(),
      minute: value.getMinutes()
    };
  }

  var str = String(value || "").trim();
  if (!str) return null;

  var parts = str.split(/\s+/);
  var dateParts = parts[0].split(/[\/-]/);
  var timeParts = (parts[1] || "0:0:0").split(":");
  if (dateParts.length !== 3) return null;

  var first = parseInt(dateParts[0], 10);
  var second = parseInt(dateParts[1], 10);
  var third = parseInt(dateParts[2], 10);
  if ([first, second, third].some(isNaN)) return null;

  var isYearFirst = String(dateParts[0]).length === 4;
  return {
    day: isYearFirst ? third : first,
    month: second,
    year: isYearFirst ? first : third,
    hour: parseInt(timeParts[0], 10) || 0,
    minute: parseInt(timeParts[1], 10) || 0
  };
}

// คืนคีย์รอบเวลาสำหรับสรุปข้อมูล: 08:00-20:00 หรือ 20:00-08:00
// เวลา 00:00-07:59 จะใช้วันที่เริ่มกะของวันก่อนหน้า
function getShiftKey_(year, month, day, hour) {
  var gregorianYear = toGregorianYear_(year);
  var h = parseInt(hour, 10);
  if (isNaN(gregorianYear) || isNaN(h)) return "";

  var shiftStartHour = (h >= 8 && h < 20) ? 8 : 20;
  var shiftDate = new Date(Date.UTC(gregorianYear, month - 1, day));
  if (h < 8) shiftDate.setUTCDate(shiftDate.getUTCDate() - 1);

  var yyyy = shiftDate.getUTCFullYear();
  var mm = String(shiftDate.getUTCMonth() + 1).padStart(2, '0');
  var dd = String(shiftDate.getUTCDate()).padStart(2, '0');
  return yyyy + '-' + mm + '-' + dd + '@' + String(shiftStartHour).padStart(2, '0');
}

// อ่านแถว Log ของกะปัจจุบันที่ยังไม่ VOID
// คืน [{job, model, hour, station, shift}] เพื่อใช้ทั้ง Dashboard และกู้หน้าจอหลัง refresh
function readTodayLogRows_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var lastCol = sheet.getLastColumn();
  var shiftColumnIndex = getLogShiftColumn_(sheet);
  var numCols = Math.max(6, shiftColumnIndex + 1);
  numCols = Math.min(lastCol, numCols);
  if (numCols < 5) return []; // ต้องมีอย่างน้อย Timestamp..Status

  var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var now = new Date();
  var currentShiftKey = getShiftKey_(now.getFullYear(), now.getMonth() + 1, now.getDate(), now.getHours());
  var shiftRows = [];

  for (var i = 0; i < data.length; i++) {
    var timestamp = parseLogTimestamp_(data[i][0]);
    if (!timestamp) continue;

    var model = String(data[i][2] || "").trim();
    var status = String(data[i][4] || "").trim().toUpperCase();
    var rowShiftKey = getShiftKey_(timestamp.year, timestamp.month, timestamp.day, timestamp.hour);
    if (rowShiftKey === currentShiftKey && status !== "VOID" && model !== "") {
      shiftRows.push({
        job: String(data[i][1] || "").trim(),
        model: model,
        hour: String(timestamp.hour).padStart(2, '0'),
        station: numCols >= 6 ? String(data[i][5] || "").trim() : "",
        shift: shiftColumnIndex >= 0 ? normalizeShiftName_(data[i][shiftColumnIndex]) : ""
      });
    }
  }
  return shiftRows;
}

// ดึงข้อมูลการผลิตของกะปัจจุบันทุก Line (ใช้โดย Dashboard mode)
function getTodayProductionData() {
  var shiftData = readTodayLogRows_().map(function(row) {
    return { model: row.model, hour: row.hour, shift: row.shift };
  });
  return JSON.stringify(shiftData);
}

// ดึงข้อมูลการผลิตของกะปัจจุบันเฉพาะ Line เพื่อกู้ยอดหน้าจอหลัง refresh/deploy
function getTodayStationData(station) {
  var targetStation = String(station || "").trim();
  if (!targetStation) return JSON.stringify([]);

  var stationData = readTodayLogRows_().filter(function(row) {
    return row.station === targetStation;
  }).map(function(row) {
    return { model: row.model, hour: row.hour, job: row.job, shift: row.shift };
  });

  return JSON.stringify(stationData);
}
