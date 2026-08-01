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
// TAT_Calculator.gs - Optional standalone/manual TAT utilities. Production
//                     workflow TAT is self-contained in WorkflowEngine.gs.
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

// ============================================
// SERVER-SIDE CACHE (makes "initial load" instant for EVERY user, not
// just the same browser)
// ============================================
// Google Sheets reads are the slow part of every wf* call. Without this,
// each login/page-open recomputes Home summary / permissions / table data
// fresh from the Sheet every single time. CacheService stores the last
// computed result for FMS_CACHE_TTL_SECONDS (10 minutes) so a cache HIT
// returns instantly - no Sheet read at all - while a scheduled trigger
// (see fmsScheduledRefresh() below) keeps the Home/permissions cache
// pre-warmed in the background, so a user logging in normally sees data
// that is at most ~10 minutes old, INSTANTLY, even on their very first
// request of the day.
var FMS_CACHE_TTL_SECONDS = 600; // 10 minutes - matches the client's own auto-refresh interval

// Human-readable deployment fingerprint. It is shown on the login page and
// by fmsDiagnose(), making it immediately obvious when an old /exec version
// or a different Apps Script project is being opened.
var FMS_BUILD_ID = '2026-07-22-column-width-2';

function fmsCacheKey(parts) {
  return 'fms_v1_' + parts.map(function (p) { return String(p); }).join('|');
}

function fmsCacheGet(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null; // cache unavailable/corrupt - caller just computes fresh
  }
}

function fmsCacheSet(key, value, ttlSeconds) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), ttlSeconds || FMS_CACHE_TTL_SECONDS);
  } catch (e) {
    // Value too large (CacheService caps a single entry at 100KB) or the
    // cache service is temporarily unavailable - silently skip caching.
    // Everything still works, just without the instant-cache benefit for
    // this one oversized payload (e.g. a huge Table View for one header).
  }
}

function fmsCacheRemove(key) {
  try { CacheService.getScriptCache().remove(key); } catch (e) { /* ignore */ }
}

// ============================================
// AUTOMATIC BACKGROUND REFRESH (every 10 minutes, even with nobody active)
// ============================================
//
// ONE-TIME SETUP REQUIRED: select "installFmsAutoRefreshTrigger" in the
// function dropdown above and click Run once. Apps Script triggers can
// only be created by an authorized run from the editor - this is a
// platform requirement, not something any code can do fully automatically
// on its own. After that one click, fmsScheduledRefresh() runs every 10
// minutes on its own, forever (survives redeployments; only needs to be
// installed once per script project). fmsDiagnose() reports whether it is
// currently installed.
function fmsScheduledRefresh() {
  try {
    var ss = SpreadsheetApp.openByUrl(FMS_SHEET_URL);
    var masterSheet = ss.getSheetByName(FMS_MASTER_SHEET);
    if (!masterSheet) return;
    var data = masterSheet.getDataRange().getValues();

    var startTime = Date.now();
    var BUDGET_MS = 4 * 60 * 1000; // stay well under Apps Script's ~6 min trigger execution limit

    for (var i = 4; i < data.length; i++) {
      if (Date.now() - startTime > BUDGET_MS) break; // ran out of time this cycle - rest catch up next cycle
      var login = String(data[i][0]).trim();
      if (!login || login === 'undefined') continue;

      try {
        var homeResult = getHomeSummaryComputeFresh(FMS_SHEET_URL, FMS_MASTER_SHEET, FMS_STEPS_SHEET, login);
        if (homeResult && homeResult.success) {
          fmsCacheSet(fmsCacheKey(['home', login]), homeResult);
        }
        var permsResult = getUserPermissionsComputeFresh(login);
        if (permsResult && permsResult.success) {
          fmsCacheSet(fmsCacheKey(['perms', login]), permsResult);
        }
      } catch (innerErr) {
        // One user's malformed row should never abort refresh for everyone else.
        continue;
      }
    }
  } catch (e) {
    // Never let a trigger failure surface to a user - worst case, the next
    // login simply computes fresh (exactly like before this feature existed).
  }
}

function installFmsAutoRefreshTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'fmsScheduledRefresh') {
      return 'Already installed - fmsScheduledRefresh runs every 10 minutes.';
    }
  }
  ScriptApp.newTrigger('fmsScheduledRefresh').timeBased().everyMinutes(10).create();
  // Warm the cache immediately too, so the benefit starts right now
  // instead of waiting for the first scheduled run.
  fmsScheduledRefresh();
  return 'Installed! fmsScheduledRefresh will now run automatically every 10 minutes, and the cache has been warmed immediately.';
}

function isFmsAutoRefreshTriggerInstalled() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'fmsScheduledRefresh') return true;
    }
  } catch (e) { /* ignore */ }
  return false;
}

// ============================================
// DEPLOYMENT CONTRACT + WEB ENTRY POINT
// ============================================
// Apps Script deployments are immutable snapshots. Saving new files does
// NOT update an existing /exec URL until that deployment is edited and a
// "New version" is selected. These checks make an incomplete/mixed snapshot
// fail before login with a useful setup page instead of later throwing
// "validateLogin is not a function" or "formatResult is not defined".
function fmsGetRuntimeInfo() {
  var requiredFunctions = {
    'validateLogin': typeof validateLogin,
    'getUserPermissions': typeof getUserPermissions,
    'getDashboardHtml': typeof getDashboardHtml,
    'include': typeof include,
    'wfCalculateTAT': typeof wfCalculateTAT,
    'formatDateSafe': typeof formatDateSafe,
    'wfGetHomeSummary': typeof wfGetHomeSummary,
    'wfGetStepTableData': typeof wfGetStepTableData,
    'wfGetMultiStepTableData': typeof wfGetMultiStepTableData,
    'wfGetStepForm': typeof wfGetStepForm,
    'wfSubmitStep': typeof wfSubmitStep
  };
  var missing = [];
  for (var name in requiredFunctions) {
    if (requiredFunctions[name] !== 'function') missing.push(name + '()');
  }
  if (typeof WF_BUILD_ID === 'undefined' || WF_BUILD_ID !== FMS_BUILD_ID) {
    missing.push('WorkflowEngine.gs build mismatch (expected ' + FMS_BUILD_ID + ')');
  }

  // Raw existence checks catch missing include files; evaluating both page
  // templates additionally catches broken include names/scriptlets now,
  // before a user can pass login and hit a blank Dashboard.
  var requiredHtml = ['Login', 'Dashboard', 'Scripts', 'Styles', 'ThemeEngine'];
  for (var i = 0; i < requiredHtml.length; i++) {
    try {
      HtmlService.createHtmlOutputFromFile(requiredHtml[i]);
    } catch (e) {
      missing.push(requiredHtml[i] + '.html');
    }
  }
  var clientBuildFiles = ['Scripts', 'Styles'];
  for (var clientIndex = 0; clientIndex < clientBuildFiles.length; clientIndex++) {
    var clientFile = clientBuildFiles[clientIndex];
    try {
      var clientSource = HtmlService.createHtmlOutputFromFile(clientFile).getContent();
      if (clientSource.indexOf(FMS_BUILD_ID) === -1) {
        missing.push(clientFile + '.html build mismatch (expected ' + FMS_BUILD_ID + ')');
      }
    } catch (clientBuildError) {
      // The missing-file check above already provides the clearest message.
    }
  }
  var templates = ['Login', 'Dashboard'];
  for (var t = 0; t < templates.length; t++) {
    try {
      HtmlService.createTemplateFromFile(templates[t]).evaluate().getContent();
    } catch (templateError) {
      missing.push(templates[t] + '.html render (' + templateError.message + ')');
    }
  }

  return {
    success: missing.length === 0,
    buildId: FMS_BUILD_ID,
    missing: missing,
    message: missing.length ? 'Missing from this deployed version: ' + missing.join(', ') : 'FMS runtime is complete.'
  };
}

function fmsEscapeHtmlServer(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function fmsRuntimeStatusPage(info, isExplicitCheck) {
  var ok = info && info.success;
  var title = ok ? 'FMS deployment is ready' : 'FMS deployment is incomplete';
  var detail = ok
    ? 'Login, dashboard, table, form and workflow server functions are available.'
    : fmsEscapeHtmlServer((info && info.message) || 'Unknown deployment error.');
  var next = ok
    ? 'You can remove ?fms_check=1 from the URL and use the application.'
    : 'Open the Apps Script project attached to THIS deployment, replace the required files, save, then use Deploy > Manage deployments > Edit > New version. Do not create a second deployment URL unless you intend to switch URLs.';
  var html = '<!doctype html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + title + '</title><style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#0f172a;color:#e2e8f0;font:14px Arial,sans-serif}' +
    '.card{max-width:720px;margin:24px;padding:28px;border:1px solid #334155;border-radius:16px;background:#111827;box-shadow:0 20px 50px #0006}' +
    'h1{margin:0 0 12px;color:' + (ok ? '#4ade80' : '#fb7185') + '}p{line-height:1.6}.build{color:#94a3b8;font-family:monospace}.next{padding:14px;border-radius:10px;background:#1e293b}</style></head><body><main class="card">' +
    '<h1>' + title + '</h1><p>' + detail + '</p><p class="build">Build: ' + fmsEscapeHtmlServer((info && info.buildId) || FMS_BUILD_ID) + '</p>' +
    '<p class="next">' + fmsEscapeHtmlServer(next) + '</p>' +
    (isExplicitCheck ? '' : '<p>The application was stopped before login so it cannot fail halfway through loading a table or saving a form.</p>') +
    '</main></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doGet(e) {
  // FMS_DOGET_MARKER - do not remove. fmsDiagnose() looks for this string
  // to confirm THIS doGet() is the one Apps Script is actually serving.
  var runtime = fmsGetRuntimeInfo();
  var explicitCheck = !!(e && e.parameter && String(e.parameter.fms_check) === '1');
  if (explicitCheck || !runtime.success) return fmsRuntimeStatusPage(runtime, explicitCheck);

  try {
    return HtmlService.createTemplateFromFile('Login')
      .evaluate()
      .setTitle('FMS - Flow Management System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return fmsRuntimeStatusPage({
      success: false,
      buildId: FMS_BUILD_ID,
      missing: [],
      message: 'Login template could not be rendered: ' + err.message
    }, false);
  }
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
  lines.push('BUILD    ' + FMS_BUILD_ID);
  try {
    var deployedUrl = ScriptApp.getService().getUrl();
    lines.push(deployedUrl ? 'WEB APP ' + deployedUrl : 'PROBLEM  This script has no web-app deployment URL yet.');
    if (deployedUrl) lines.push('CHECK    ' + deployedUrl + '?fms_check=1');
  } catch (serviceErr) {
    lines.push('PROBLEM  Could not read the web-app deployment URL: ' + serviceErr.message);
  }

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
    'wfCalculateTAT': typeof wfCalculateTAT,
    'formatDateSafe': typeof formatDateSafe,
    'wfGetHomeSummary': typeof wfGetHomeSummary,
    'wfGetStepTableData': typeof wfGetStepTableData,
    'wfGetMultiStepTableData': typeof wfGetMultiStepTableData,
    'wfGetStepForm': typeof wfGetStepForm,
    'wfSubmitStep': typeof wfSubmitStep
  };
  for (var name in fns) {
    if (fns[name] === 'function') {
      lines.push('OK       ' + name + '()');
    } else {
      lines.push('PROBLEM  ' + name + '() missing - replace and save the latest server files');
      lines.push('         (Code.gs and WorkflowEngine.gs).');
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

  // 6) Is the 10-minute background refresh trigger installed? Without it,
  // caching still makes repeat loads instant, but the "10-minute-old data
  // even on a NEW user's very first login" guarantee needs this trigger.
  try {
    if (isFmsAutoRefreshTriggerInstalled()) {
      lines.push('OK       10-minute auto-refresh trigger is installed');
    } else {
      lines.push('PROBLEM  10-minute auto-refresh trigger is NOT installed.');
      lines.push('         Run installFmsAutoRefreshTrigger() once from this editor');
      lines.push('         (select it in the function dropdown above and click Run).');
      lines.push('         One-time setup only - it then runs forever on its own.');
    }
  } catch (e) {
    lines.push('PROBLEM  Could not check triggers: ' + e.message);
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
    var normalizedLoginId = String(loginId == null ? '' : loginId).trim();
    var normalizedPassword = String(password == null ? '' : password);
    if (!normalizedLoginId || !normalizedPassword) {
      return { success: false, message: 'Login ID and Password are required.', buildId: FMS_BUILD_ID };
    }

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
      
      if (sheetLogin.toUpperCase() === normalizedLoginId.toUpperCase() && sheetPassword === normalizedPassword) {
        return { 
          success: true, 
          message: 'Login Successful! Welcome ' + userName,
          userName: userName,
          loginId: sheetLogin,
          sheetUrl: FMS_SHEET_URL,
          masterSheet: FMS_MASTER_SHEET,
          stepsSheet: FMS_STEPS_SHEET,
          dropdownSheet: FMS_DROPDOWN_SHEET,
          buildId: FMS_BUILD_ID
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
// Cache-first: a HIT returns instantly with no Sheet read at all. The
// background trigger (fmsScheduledRefresh) keeps this warm every 10
// minutes; a cache MISS (e.g. right after a deployment, or the trigger
// hasn't run yet) computes fresh AND populates the cache for next time,
// so the very next call - by this user or any other - is instant too.
function getUserPermissions(userName) {
  var cacheKey = fmsCacheKey(['perms', userName]);
  var cached = fmsCacheGet(cacheKey);
  if (cached) return cached;

  var fresh = getUserPermissionsComputeFresh(userName);
  if (fresh && fresh.success) fmsCacheSet(cacheKey, fresh);
  return fresh;
}

function getUserPermissionsComputeFresh(userName) {
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
