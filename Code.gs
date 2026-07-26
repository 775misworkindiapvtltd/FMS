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

function doGet() {
  return HtmlService.createTemplateFromFile('Login')
    .evaluate()
    .setTitle('FMS - Flow Management System')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
