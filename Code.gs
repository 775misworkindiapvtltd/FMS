// ============================================
// FMS - FLOW MANAGEMENT SYSTEM
// Google Apps Script - Server Side Code
// ============================================
// 
// FILE STRUCTURE:
// Code.gs          - Server-side logic (this file)
// Login.html       - Login page
// Dashboard.html   - Main dashboard after login
// Styles.html      - Shared CSS styles (included via include())
// Scripts.html     - Shared JS scripts (included via include())
// ThemeEngine.html - Appearance & Dynamic Color Scheme system (FOUC-safe,
//                    included FIRST in <head> on both Login + Dashboard)
// TAT_Calculator.gs - TAT calculation functions
// WorkflowEngine.gs - Pending list / popup form / step submission logic
//
// ============================================

// ============================================
// FIXED CONFIGURATION (not entered by user - hardcoded for this deployment)
// User only enters Login ID + Password on the Login page.
// ============================================
var FMS_SHEET_URL = 'https://docs.google.com/spreadsheets/d/10Aqav9bM_XQ28TkfjoXOoj8jRHenMerIEcx8gYkwo48/edit?usp=drive_web&ouid=114024121405657848674';
var FMS_MASTER_SHEET = 'MASTER';
var FMS_STEPS_SHEET = 'STEPS';
var FMS_DROPDOWN_SHEET = 'DROPDPWN';

// Google Drive folder where FILE-type question uploads are saved.
// Leave as '' to auto-create/use a folder named "FMS Uploads" in the
// script owner's My Drive (default behavior). To use a SPECIFIC folder
// instead, paste its Folder ID here (the long string in its Drive URL:
// https://drive.google.com/drive/folders/<THIS_PART_IS_THE_ID>).
var FMS_DRIVE_FOLDER_ID = '';

function doGet() {
  // FMS_DOGET_MARKER - do not remove. fmsDiagnose() looks for this string
  // to confirm THIS doGet() is the one Apps Script is actually serving.
  return HtmlService.createTemplateFromFile('Login')
    .evaluate()
    .setTitle('FMS - Flow Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ============================================
 * SETUP CHECKER - run this from the Apps Script editor
 * ============================================
 * Select "fmsDiagnose" in the function dropdown and click Run. The
 * Execution log prints exactly what is wrong.
 *
 * Why this exists: if a leftover starter-kit / template file also defines
 * doGet(), Apps Script keeps only the LAST definition it loads, so the web
 * app can silently keep serving the starter page even after every FMS file
 * has been pasted in correctly. That failure is invisible from the editor,
 * so this check reports it directly instead of leaving you guessing.
 */
function fmsDiagnose() {
  var lines = [];
  lines.push('===== FMS SETUP CHECK =====');

  // 1) Is the doGet() being served actually FMS's one?
  try {
    var activeSource = doGet.toString();
    if (activeSource.indexOf('FMS_DOGET_MARKER') !== -1) {
      lines.push('OK       doGet() -> FMS Login page');
    } else {
      lines.push('PROBLEM  doGet() is NOT FMS\'s version.');
      lines.push('         Another file in this project also defines doGet()');
      lines.push('         (usually the starter kit\'s Code/Index/Main file) and it');
      lines.push('         is overriding FMS. Delete that file, or remove its');
      lines.push('         doGet() function, then run this check again.');
    }
  } catch (e) {
    lines.push('PROBLEM  doGet() not found at all: ' + e.message);
  }

  // 2) Do all required HTML files exist, under the exact expected names?
  var required = ['Login', 'Dashboard', 'Scripts', 'Styles', 'ThemeEngine'];
  for (var i = 0; i < required.length; i++) {
    try {
      HtmlService.createHtmlOutputFromFile(required[i]);
      lines.push('OK       ' + required[i] + '.html found');
    } catch (e) {
      lines.push('PROBLEM  ' + required[i] + '.html is MISSING or misnamed.');
      lines.push('         Create an HTML file named exactly "' + required[i] + '"');
      lines.push('         (type the name WITHOUT ".html" - Apps Script adds it).');
    }
  }

  // 3) Leftover starter-kit files that commonly hijack doGet()
  var leftovers = ['Index', 'index', 'Main', 'Server'];
  for (var j = 0; j < leftovers.length; j++) {
    try {
      HtmlService.createHtmlOutputFromFile(leftovers[j]);
      lines.push('PROBLEM  Leftover starter file "' + leftovers[j] + '.html" still exists.');
      lines.push('         Delete it - FMS never uses it.');
    } catch (e) { /* not present = good */ }
  }

  // 4) Are the required server-side functions present?
  var fns = {
    'validateLogin': typeof validateLogin,
    'getUserPermissions': typeof getUserPermissions,
    'getDashboardHtml': typeof getDashboardHtml,
    'include': typeof include,
    'wfGetHomeSummary': typeof wfGetHomeSummary,
    'wfGetStepTableData': typeof wfGetStepTableData,
    'wfGetMultiStepTableData': typeof wfGetMultiStepTableData,
    'wfSubmitStep': typeof wfSubmitStep
  };
  for (var name in fns) {
    if (fns[name] === 'function') {
      lines.push('OK       ' + name + '()');
    } else {
      lines.push('PROBLEM  ' + name + '() missing - paste the .gs file that defines it');
      lines.push('         (Code.gs / WorkflowEngine.gs / TAT_Calculator.gs).');
    }
  }

  // 5) Can the configured Spreadsheet actually be opened?
  try {
    var ss = SpreadsheetApp.openByUrl(FMS_SHEET_URL);
    lines.push('OK       Spreadsheet opened: ' + ss.getName());
    var needed = [FMS_MASTER_SHEET, FMS_STEPS_SHEET];
    for (var k = 0; k < needed.length; k++) {
      lines.push((ss.getSheetByName(needed[k]) ? 'OK       tab "' : 'PROBLEM  missing tab "') + needed[k] + '"');
    }
  } catch (e) {
    lines.push('PROBLEM  Cannot open the Spreadsheet: ' + e.message);
  }

  var report = lines.join('\n');
  Logger.log(report);
  return report;
}

/**
 * Include external HTML files (for CSS/JS separation)
 * Usage in HTML: <?!= include('Styles') ?>
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================
// LOGIN VALIDATION
// Called from Login.html with ONLY loginId + password.
// Sheet URL / Sheet names are fixed server-side (FMS_SHEET_URL etc. above).
// ============================================
function validateLogin(loginId, password) {
  try {
    var ss = SpreadsheetApp.openByUrl(FMS_SHEET_URL);
    var sheet = ss.getSheetByName(FMS_MASTER_SHEET);
    
    if (!sheet) {
      return { success: false, message: 'Sheet not found! Check Sheet Name.' };
    }
    
    var data = sheet.getDataRange().getValues();
    
    // Row 4 (index 3) has headers: LOGIN, PASSWORD, etc.
    // Data starts from row 5 (index 4)
    // Col A (index 0) = LOGIN
    // Col B (index 1) = PASSWORD
    // Col C (index 2) = Name / Permission reference
    
    for (var i = 4; i < data.length; i++) {
      var sheetLogin = String(data[i][0]).trim();
      var sheetPassword = String(data[i][1]).trim();
      var userName = sheetLogin; // Using login name as display name
      
      if (sheetLogin === '' || sheetLogin === 'undefined') continue;
      
      if (sheetLogin.toUpperCase() === loginId.toUpperCase() && sheetPassword === password) {
        return { 
          success: true, 
          message: 'Login Successful! Welcome ' + userName,
          userName: userName,
          loginId: sheetLogin,
          sheetUrl: FMS_SHEET_URL,
          masterSheet: FMS_MASTER_SHEET,
          stepsSheet: FMS_STEPS_SHEET,
          dropdownSheet: FMS_DROPDOWN_SHEET
        };
      }
    }
    
    return { success: false, message: 'Login Failed! ID or Password not matched.' };
    
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================
// GET PERMISSIONS FOR LOGGED IN USER
// ============================================
function getUserPermissions(userName) {
  try {
    var ss = SpreadsheetApp.openByUrl(FMS_SHEET_URL);
    var sheet = ss.getSheetByName(FMS_STEPS_SHEET);
    
    if (!sheet) {
      return { success: false, message: 'Steps sheet not found!', permissions: {} };
    }
    
    var data = sheet.getDataRange().getValues();
    var permissions = [];
    
    // STEPS sheet structure:
    // Col A = Header/Category (like "Utility FMS", "Copy of AMC FMS")
    // Col B = Sub-item/Page name (like "Check Bill Details Outstanding")
    // Col C = User Name who has permission
    
    for (var i = 0; i < data.length; i++) {
      var colC = String(data[i][2]).trim().toUpperCase();
      
      if (colC === userName.toUpperCase()) {
        var header = String(data[i][0]).trim();
        var subItem = String(data[i][1]).trim();
        
        if (header && subItem) {
          permissions.push({
            header: header,
            subItem: subItem
          });
        }
      }
    }
    
    // Group by header
    var grouped = {};
    for (var j = 0; j < permissions.length; j++) {
      var h = permissions[j].header;
      if (!grouped[h]) {
        grouped[h] = [];
      }
      grouped[h].push(permissions[j].subItem);
    }
    
    return { success: true, permissions: grouped };
    
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message, permissions: {} };
  }
}

// ============================================
// GET DASHBOARD HTML (called after login success)
// ============================================
function getDashboardHtml() {
  return HtmlService.createTemplateFromFile('Dashboard').evaluate().getContent();
}
