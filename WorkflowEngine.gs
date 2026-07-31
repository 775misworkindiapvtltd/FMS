// ============================================
// FMS WORKFLOW ENGINE
// Pending List -> Popup Form -> Row Update (per Step)
// ============================================
//
// TARGET SHEET LAYOUT (fixed template - e.g. "Utility FMS", "AMC FMS" tabs):
// Row 1 : Title banner + "Step 1" / "Step 2" / "Step 3" labels (visual only)
// Row 2 : URL/link (base) + ACTUAL STEP NAME text per step block
//         (e.g. "PI FOLLOW UP FROM VENDOR") -- USED to locate step's columns
// Row 3-5 : "Yes/No/Draw" notification mini-table -- IGNORED
// Row 6 : Data Type per column (DROPDOWN, NUMBER, LONG TEXT, FILE, etc.)
// Row 7 : Field Name per column (shown to user as question label).
//         First 2 fields of every step block are always "Planned" and "Actual"
// Row 8+: Actual data rows (one row = one work item, flows left-to-right
//         through each Step's columns as it gets actioned)
//
// NOTE: Row numbers below are configurable constants in case the real sheet
// differs slightly from what was described - adjust and no other code needs
// to change.
// ============================================

var WF_STEP_NAME_ROW = 2;   // Row containing the exact step name text
var WF_TYPE_ROW = 6;        // Row containing data type per column
var WF_FIELD_NAME_ROW = 7;  // Row containing field/question name per column
var WF_DATA_START_ROW = 8;  // First row where actual data entries begin

// Default TAT/office config used when auto-computing next step's "Planned" date.
// TODO: Wire these to real per-step TAT hours / office config once defined.
var WF_DEFAULT_OFFICE_START = '10:00';
var WF_DEFAULT_OFFICE_END = '18:30';
var WF_DEFAULT_WEEKOFF = [0]; // Sunday off
var WF_DEFAULT_HOLIDAY_SHEET = 'HOLIDAYS';
var WF_DEFAULT_TAT_HOURS = 24;

// ============================================
// STEP COLUMN DETECTION
// ============================================

/**
 * Locate the column range (start/end) occupied by a given Step's block
 * inside the target sheet, by matching WF_STEP_NAME_ROW text.
 * Uses merged-cell span if available, otherwise falls back to
 * "until next step's start column" or "end of sheet".
 */
function findStepColumnRange(sheet, stepName) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return null;

  var rowValues = sheet.getRange(WF_STEP_NAME_ROW, 1, 1, lastCol).getValues()[0];
  var boundaries = [];
  for (var c = 0; c < lastCol; c++) {
    var val = String(rowValues[c]).trim();
    if (val !== '') boundaries.push({ col: c + 1, name: val });
  }

  var targetIndex = -1;
  for (var i = 0; i < boundaries.length; i++) {
    if (boundaries[i].name.toUpperCase() === String(stepName).trim().toUpperCase()) {
      targetIndex = i;
      break;
    }
  }
  if (targetIndex === -1) return null;

  var startCol = boundaries[targetIndex].col;
  var endCol;

  try {
    var cell = sheet.getRange(WF_STEP_NAME_ROW, startCol);
    var merged = cell.getMergedRanges();
    if (merged && merged.length > 0) {
      endCol = merged[0].getColumn() + merged[0].getNumColumns() - 1;
    } else if (targetIndex < boundaries.length - 1) {
      endCol = boundaries[targetIndex + 1].col - 1;
    } else {
      endCol = lastCol;
    }
  } catch (e) {
    // getMergedRanges not available (e.g. in test mocks) - use boundary fallback
    if (targetIndex < boundaries.length - 1) {
      endCol = boundaries[targetIndex + 1].col - 1;
    } else {
      endCol = lastCol;
    }
  }

  return { startCol: startCol, endCol: endCol, orderIndex: targetIndex, allBoundaries: boundaries };
}

/**
 * Get field metadata (col, type, name) for a step's column range,
 * reading WF_TYPE_ROW and WF_FIELD_NAME_ROW.
 */
function getStepFields(sheet, range) {
  var width = range.endCol - range.startCol + 1;
  var types = sheet.getRange(WF_TYPE_ROW, range.startCol, 1, width).getValues()[0];
  var names = sheet.getRange(WF_FIELD_NAME_ROW, range.startCol, 1, width).getValues()[0];

  var fields = [];
  for (var i = 0; i < width; i++) {
    fields.push({
      col: range.startCol + i,
      type: String(types[i]).trim(),
      name: String(names[i]).trim()
    });
  }
  return fields;
}

/**
 * Identify the Planned and Actual column numbers within a field list,
 * matched by name (not fixed position) for robustness.
 */
function getPlannedActualCols(fields) {
  var plannedCol = null, actualCol = null;
  fields.forEach(function (f) {
    var n = f.name.toLowerCase();
    if (n.indexOf('planned') !== -1) plannedCol = f.col;
    if (n.indexOf('actual') !== -1) actualCol = f.col;
  });
  return { plannedCol: plannedCol, actualCol: actualCol };
}

/**
 * A field is only shown (in the form OR the table) if it has a Data Type
 * configured in Row 6 - EXCEPT "Planned" and "Actual", which are always
 * shown regardless of whether Row 6 has a type for them:
 * - "Planned" is always shown, always non-editable (auto-computed via TAT).
 * - "Actual" is always kept internally (needed for pending/completed
 *   status detection) and shown as a column in the table.
 * Any other column left with an EMPTY Data Type cell is treated as "not
 * configured yet" and is hidden everywhere (form + table).
 */
function isFieldConfigured(field) {
  var n = field.name.toLowerCase();
  if (n.indexOf('planned') !== -1) return true;
  if (n.indexOf('actual') !== -1) return true;
  if (n.indexOf('modified by') !== -1) return true;
  if (n.indexOf('modified at') !== -1) return true;
  return !!field.type; // must have a non-empty Data Type (Row 6) to appear
}

/**
 * "Modified By" / "Modified At" are an optional audit trail: written
 * automatically on every submit (see submitStepData()), never shown as
 * an editable question in the popup form, and hidden by default in the
 * table (still toggleable via the Show/Hide Columns panel for anyone who
 * wants to check who filled a row).
 */
function isAuditTrailField(fieldName) {
  var n = (fieldName || '').toLowerCase();
  return n.indexOf('modified by') !== -1 || n.indexOf('modified at') !== -1;
}

// ============================================
// STEP SEQUENCE (from STEPS sheet, for a given Header)
// ============================================

/**
 * Get all Step names (Col B) belonging to a Header (Col A), in sheet-row
 * order, regardless of which user (Col C) they're assigned to.
 * This defines the pipeline sequence for that Header.
 */
function getAllStepsForHeader(ss, stepsSheetName, header) {
  var sheet = ss.getSheetByName(stepsSheetName);
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  var steps = [];
  for (var i = 0; i < data.length; i++) {
    var rowHeader = String(data[i][0]).trim();
    var rowStep = String(data[i][1]).trim();
    if (rowHeader.toUpperCase() === header.trim().toUpperCase() && rowStep) {
      steps.push(rowStep);
    }
  }
  return steps;
}

// ============================================
// TARGET SHEET LOOKUP (from MASTER sheet)
// ============================================

/**
 * MASTER sheet Row 1 = Sheet URLs (per header, from column C onward)
 * MASTER sheet Row 2 = Sheet Names (matching header text, e.g. "AMC FMS")
 * Given a header (from STEPS Col A, e.g. "Copy of AMC FMS"), find its
 * target Sheet URL + Sheet Name.
 *
 * Matching strategy (in order):
 * 1. Exact match (case-insensitive)
 * 2. Contains match (either string contains the other) - handles cases
 *    like STEPS="Copy of AMC FMS" vs MASTER="AMC FMS"
 */
function getTargetSheetInfo(masterSheetUrl, masterSheetName, header) {
  var ss = SpreadsheetApp.openByUrl(masterSheetUrl);
  var sheet = ss.getSheetByName(masterSheetName);
  if (!sheet) return null;

  var lastCol = sheet.getLastColumn();
  var urls = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var names = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  var headerUpper = header.trim().toUpperCase();

  // Pass 1: exact match
  for (var c = 0; c < names.length; c++) {
    var nameUpper = String(names[c]).trim().toUpperCase();
    if (nameUpper === headerUpper) {
      return { sheetUrl: String(urls[c]).trim(), sheetName: String(names[c]).trim() };
    }
  }

  // Pass 2: contains match (either direction)
  for (var c2 = 0; c2 < names.length; c2++) {
    var nameUpper2 = String(names[c2]).trim().toUpperCase();
    if (!nameUpper2) continue;
    if (headerUpper.indexOf(nameUpper2) !== -1 || nameUpper2.indexOf(headerUpper) !== -1) {
      return { sheetUrl: String(urls[c2]).trim(), sheetName: String(names[c2]).trim() };
    }
  }

  return null;
}

/**
 * Same logic as Code.gs's getUserPermissions(userName), but parametrized
 * (doesn't rely on the FMS_SHEET_URL/FMS_STEPS_SHEET globals) so it can be
 * reused here for the Home dashboard summary without duplicating logic
 * in two places that could drift apart.
 */
function getUserPermissionsInternal(masterSheetUrl, stepsSheetName, userName) {
  try {
    var ss = SpreadsheetApp.openByUrl(masterSheetUrl);
    var sheet = ss.getSheetByName(stepsSheetName);
    if (!sheet) return { success: false, message: 'Steps sheet not found!', permissions: {} };

    var data = sheet.getDataRange().getValues();
    var permissions = [];

    for (var i = 0; i < data.length; i++) {
      var colC = String(data[i][2]).trim().toUpperCase();
      if (colC === userName.toUpperCase()) {
        var header = String(data[i][0]).trim();
        var subItem = String(data[i][1]).trim();
        if (header && subItem) permissions.push({ header: header, subItem: subItem });
      }
    }

    var grouped = {};
    for (var j = 0; j < permissions.length; j++) {
      var h = permissions[j].header;
      if (!grouped[h]) grouped[h] = [];
      grouped[h].push(permissions[j].subItem);
    }

    return { success: true, permissions: grouped };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message, permissions: {} };
  }
}

// ============================================
// HOME DASHBOARD SUMMARY (Pending/Completed counts per assigned Step)
// ============================================
//
// Lightweight version of getStepTableData() that only counts row
// statuses (pending/completed) instead of building full column/cell
// data - much cheaper to run for EVERY header+step a user has, which is
// what the Home page needs to show its summary cards.
function getStepCounts(masterSheetUrl, masterSheetName, stepsSheetName, header, stepName) {
  var target = getTargetSheetInfo(masterSheetUrl, masterSheetName, header);
  if (!target) return { pending: 0, completed: 0, error: 'Target sheet not found for header: ' + header };

  var targetSs = SpreadsheetApp.openByUrl(target.sheetUrl);
  var sheet = targetSs.getSheetByName(target.sheetName);
  if (!sheet) return { pending: 0, completed: 0, error: 'Sheet "' + target.sheetName + '" not found' };

  var masterSs = SpreadsheetApp.openByUrl(masterSheetUrl);
  var allSteps = getAllStepsForHeader(masterSs, stepsSheetName, header);
  var stepIndex = -1;
  for (var i = 0; i < allSteps.length; i++) {
    if (allSteps[i].toUpperCase() === stepName.trim().toUpperCase()) { stepIndex = i; break; }
  }
  if (stepIndex === -1) return { pending: 0, completed: 0, error: 'Step not found in STEPS sheet' };

  var currentRange = findStepColumnRange(sheet, stepName);
  if (!currentRange) return { pending: 0, completed: 0, error: 'Step columns not found (check Row ' + WF_STEP_NAME_ROW + ')' };
  var currentFields = getStepFields(sheet, currentRange);
  var currentPA = getPlannedActualCols(currentFields);

  var prevActualCol = null;
  if (stepIndex > 0) {
    var prevRange = findStepColumnRange(sheet, allSteps[stepIndex - 1]);
    if (prevRange) {
      var prevFields = getStepFields(sheet, prevRange);
      prevActualCol = getPlannedActualCols(prevFields).actualCol;
    }
  }

  var lastRow = sheet.getLastRow();
  var pending = 0, completed = 0, overdue = 0;
  var rowCount = Math.max(0, lastRow - WF_DATA_START_ROW + 1);
  if (!rowCount || (stepIndex > 0 && !prevActualCol)) return { pending: 0, completed: 0, overdue: 0 };

  // Batch each required column once. Spreadsheet service calls dominate
  // Apps Script runtime; reading cell-by-cell here would multiply calls by
  // every row and every assigned step on the Home dashboard.
  var applicabilityValues = sheet.getRange(
    WF_DATA_START_ROW,
    stepIndex === 0 ? 1 : prevActualCol,
    rowCount,
    1
  ).getValues();
  var actualValues = currentPA.actualCol
    ? sheet.getRange(WF_DATA_START_ROW, currentPA.actualCol, rowCount, 1).getValues()
    : [];
  var plannedValues = currentPA.plannedCol
    ? sheet.getRange(WF_DATA_START_ROW, currentPA.plannedCol, rowCount, 1).getValues()
    : [];
  var nowMs = Date.now();

  for (var offset = 0; offset < rowCount; offset++) {
    if (!applicabilityValues[offset][0]) continue;

    var actualVal = currentPA.actualCol ? actualValues[offset][0] : '';
    if (actualVal !== '' && actualVal !== null) {
      completed++;
    } else {
      pending++;
      if (currentPA.plannedCol) {
        var plannedVal = plannedValues[offset][0];
        if (plannedVal instanceof Date && plannedVal.getTime() < nowMs) overdue++;
      }
    }
  }

  return { pending: pending, completed: completed, overdue: overdue };
}

/**
 * Builds the full Home dashboard summary: for every Header/Step the user
 * has permission for (same source as getUserPermissions), returns the
 * Pending/Completed counts, plus grand totals across everything.
 */
// Cache-first, same pattern as getUserPermissions() in Code.gs: a HIT
// returns instantly with zero Sheet reads. This is the single most
// expensive call in the whole app (loops getStepCounts() across every
// step the user owns), so it benefits the most from caching - this is
// the main fix for "initial data loading takes too long".
function getHomeSummary(masterSheetUrl, masterSheetName, stepsSheetName, userName) {
  var cacheKey = fmsCacheKey(['home', userName]);
  var cached = fmsCacheGet(cacheKey);
  if (cached) return cached;

  var fresh = getHomeSummaryComputeFresh(masterSheetUrl, masterSheetName, stepsSheetName, userName);
  if (fresh && fresh.success) fmsCacheSet(cacheKey, fresh);
  return fresh;
}

function getHomeSummaryComputeFresh(masterSheetUrl, masterSheetName, stepsSheetName, userName) {
  try {
    var permsRes = getUserPermissionsInternal(masterSheetUrl, stepsSheetName, userName);
    if (!permsRes.success) return { success: false, message: permsRes.message, groups: [] };

    var groups = [];
    var totalPending = 0, totalCompleted = 0, totalOverdue = 0;

    for (var header in permsRes.permissions) {
      var steps = permsRes.permissions[header];
      var stepEntries = [];

      steps.forEach(function (stepName) {
        var counts = getStepCounts(masterSheetUrl, masterSheetName, stepsSheetName, header, stepName);
        stepEntries.push({
          step: stepName,
          pending: counts.pending || 0,
          completed: counts.completed || 0,
          overdue: counts.overdue || 0,
          error: counts.error || null
        });
        totalPending += (counts.pending || 0);
        totalCompleted += (counts.completed || 0);
        totalOverdue += (counts.overdue || 0);
      });

      groups.push({ header: header, steps: stepEntries });
    }

    return { success: true, groups: groups, totalPending: totalPending, totalCompleted: totalCompleted, totalOverdue: totalOverdue };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message, groups: [] };
  }
}

// ============================================
// PENDING ROWS FOR A STEP
// ============================================

/**
 * Get rows pending for a given Step:
 * - Current step's Actual cell is empty
 * - AND (it's the first step in sequence AND base row has data)
 *   OR (previous step's Actual cell is filled)
 */
function getPendingRowsForStep(masterSheetUrl, masterSheetName, stepsSheetName, header, stepName) {
  try {
    var target = getTargetSheetInfo(masterSheetUrl, masterSheetName, header);
    if (!target) return { success: false, message: 'Target sheet info not found in MASTER for header: ' + header, rows: [] };

    var targetSs = SpreadsheetApp.openByUrl(target.sheetUrl);
    var sheet = targetSs.getSheetByName(target.sheetName);
    if (!sheet) return { success: false, message: 'Target sheet "' + target.sheetName + '" not found', rows: [] };

    var masterSs = SpreadsheetApp.openByUrl(masterSheetUrl);
    var allSteps = getAllStepsForHeader(masterSs, stepsSheetName, header);
    var stepIndex = -1;
    for (var i = 0; i < allSteps.length; i++) {
      if (allSteps[i].toUpperCase() === stepName.trim().toUpperCase()) { stepIndex = i; break; }
    }
    if (stepIndex === -1) return { success: false, message: 'Step not found in STEPS sheet for this header', rows: [] };

    var currentRange = findStepColumnRange(sheet, stepName);
    if (!currentRange) return { success: false, message: 'Step columns not found in target sheet (check Row ' + WF_STEP_NAME_ROW + ')', rows: [] };
    var currentFields = getStepFields(sheet, currentRange);
    var currentPA = getPlannedActualCols(currentFields);

    var prevActualCol = null;
    if (stepIndex > 0) {
      var prevRange = findStepColumnRange(sheet, allSteps[stepIndex - 1]);
      if (prevRange) {
        var prevFields = getStepFields(sheet, prevRange);
        prevActualCol = getPlannedActualCols(prevFields).actualCol;
      }
    }

    var lastRow = sheet.getLastRow();
    var pending = [];

    for (var r = WF_DATA_START_ROW; r <= lastRow; r++) {
      var actualVal = currentPA.actualCol ? sheet.getRange(r, currentPA.actualCol).getValue() : '';
      if (actualVal !== '' && actualVal !== null) continue; // already completed for this step

      if (stepIndex === 0) {
        var baseVal = sheet.getRange(r, 1).getValue();
        if (!baseVal) continue; // empty row, nothing to do
      } else if (prevActualCol) {
        var prevVal = sheet.getRange(r, prevActualCol).getValue();
        if (!prevVal) continue; // previous step not done yet - not pending for THIS step
      }

      pending.push({
        row: r,
        planned: currentPA.plannedCol ? formatDateSafe(sheet.getRange(r, currentPA.plannedCol).getValue()) : '',
        summary: getRowSummary(sheet, r, currentRange.startCol)
      });
    }

    return { success: true, rows: pending, sheetName: target.sheetName };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message, rows: [] };
  }
}

/**
 * Build a short identifying summary of a row using the "base" columns
 * (everything before the first step's start column).
 */
function getRowSummary(sheet, rowNumber, firstStepStartCol) {
  var baseColsCount = Math.max(0, firstStepStartCol - 1);
  if (baseColsCount === 0) return '';
  var maxCols = Math.min(baseColsCount, 6); // limit summary length
  var values = sheet.getRange(rowNumber, 1, 1, maxCols).getValues()[0];
  return values
    .map(function (v) { return String(v).trim(); })
    .filter(function (v) { return v !== ''; })
    .join(' | ');
}

function formatDateSafe(val) {
  if (val === null || val === undefined || val === '') return '';
  if (!(val instanceof Date)) return String(val);

  // Keep this formatter local and self-contained. Table/dashboard rendering
  // must not depend on TAT_Calculator.gs's test/display helper being present
  // in a particular deployment version. This intentionally mirrors the
  // fixed FMS display format without calling any function from another file.
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var pad2 = function (n) { return ('0' + n).slice(-2); };
  var datePart = pad2(val.getDate()) + ' ' + months[val.getMonth()] + ' ' + pad2(val.getFullYear() % 100);
  var isMidnight = val.getHours() === 0 && val.getMinutes() === 0 && val.getSeconds() === 0;

  if (isMidnight) return datePart;
  return datePart + ' ' + pad2(val.getHours()) + ':' + pad2(val.getMinutes()) + ':' + pad2(val.getSeconds());
}

// ============================================
// FULL TABLE DATA FOR A STEP (mirrors the real sheet)
// ============================================
//
// Returns ALL rows applicable to this step (previous step done, or first
// step with base data present) - both pending and completed - along with
// column headers taken from Row 7 (base columns + this step's own columns
// only, so the user only sees the step he has rights to work on).
//
function getStepTableData(masterSheetUrl, masterSheetName, stepsSheetName, header, stepName) {
  try {
    var target = getTargetSheetInfo(masterSheetUrl, masterSheetName, header);
    if (!target) return { success: false, message: 'Target sheet info not found in MASTER for header: ' + header };

    var targetSs = SpreadsheetApp.openByUrl(target.sheetUrl);
    var sheet = targetSs.getSheetByName(target.sheetName);
    if (!sheet) return { success: false, message: 'Target sheet "' + target.sheetName + '" not found' };

    var masterSs = SpreadsheetApp.openByUrl(masterSheetUrl);
    var allSteps = getAllStepsForHeader(masterSs, stepsSheetName, header);
    var stepIndex = -1;
    for (var i = 0; i < allSteps.length; i++) {
      if (allSteps[i].toUpperCase() === stepName.trim().toUpperCase()) { stepIndex = i; break; }
    }
    if (stepIndex === -1) return { success: false, message: 'Step not found in STEPS sheet for this header' };

    var currentRange = findStepColumnRange(sheet, stepName);
    if (!currentRange) return { success: false, message: 'Step columns not found in target sheet (check Row ' + WF_STEP_NAME_ROW + ')' };
    var currentFields = getStepFields(sheet, currentRange);
    var currentPA = getPlannedActualCols(currentFields);

    var prevActualCol = null;
    if (stepIndex > 0) {
      var prevRange = findStepColumnRange(sheet, allSteps[stepIndex - 1]);
      if (prevRange) {
        var prevFields = getStepFields(sheet, prevRange);
        prevActualCol = getPlannedActualCols(prevFields).actualCol;
      }
    }

    // Column headers: base IDENTIFYING columns are everything BEFORE the
    // FIRST step in the pipeline (not before the current step - otherwise
    // earlier steps' own columns would leak in) + this step's own columns only.
    var firstStepRange = findStepColumnRange(sheet, allSteps[0]);
    var baseColCount = firstStepRange ? Math.max(0, firstStepRange.startCol - 1) : Math.max(0, currentRange.startCol - 1);
    var headerRowVals = baseColCount > 0
      ? sheet.getRange(WF_FIELD_NAME_ROW, 1, 1, baseColCount).getValues()[0]
      : [];
    var baseTypeVals = baseColCount > 0
      ? sheet.getRange(WF_TYPE_ROW, 1, 1, baseColCount).getValues()[0]
      : [];

    // Skip this step's own columns that have no Data Type configured in
    // Row 6 (except Planned/Actual, which are always shown - see
    // isFieldConfigured()). Base identifying columns are always shown.
    var visibleFields = currentFields.filter(isFieldConfigured);

    // columnTypes runs parallel to columnNames - lets the client know
    // which columns hold dates/times, so it can render a calendar-picker
    // in that column's search box instead of a plain text search.
    // defaultHiddenCols marks which column INDEXES should start hidden in
    // the table by default (currently just the audit trail columns) -
    // the user can still reveal them anytime via Show/Hide Columns.
    var columnNames = [];
    var columnTypes = [];
    var defaultHiddenCols = [];
    for (var bc = 0; bc < baseColCount; bc++) {
      var baseName = String(headerRowVals[bc]).trim() || ('Col ' + (bc + 1));
      var baseType = String(baseTypeVals[bc]).trim().toUpperCase();
      var baseNameLower = baseName.toLowerCase();
      var baseIsDate = baseType.indexOf('DATE') !== -1 ||
        baseNameLower.indexOf('date') !== -1 || baseNameLower.indexOf('timestamp') !== -1;
      columnNames.push(baseName);
      columnTypes.push(baseIsDate ? 'DATE' : baseType);
      if (isAuditTrailField(baseName)) defaultHiddenCols.push(columnNames.length - 1);
    }
    visibleFields.forEach(function (f) {
      columnNames.push(f.name || ('Col ' + f.col));
      // "Planned"/"Actual" are always date-times even if Row 6 has no
      // type configured for them (see isFieldConfigured()).
      var nameLower = (f.name || '').toLowerCase();
      var isDateField = f.type.toUpperCase().indexOf('DATE') !== -1 ||
        nameLower.indexOf('planned') !== -1 || nameLower.indexOf('actual') !== -1 ||
        nameLower.indexOf('modified at') !== -1;
      columnTypes.push(isDateField ? 'DATE' : f.type.toUpperCase());
      if (isAuditTrailField(f.name)) defaultHiddenCols.push(columnNames.length - 1);
    });

    var lastRow = sheet.getLastRow();
    var rows = [];

    for (var r = WF_DATA_START_ROW; r <= lastRow; r++) {
      var isApplicable = false;
      if (stepIndex === 0) {
        var baseVal = sheet.getRange(r, 1).getValue();
        isApplicable = !!baseVal;
      } else if (prevActualCol) {
        var prevVal = sheet.getRange(r, prevActualCol).getValue();
        isApplicable = !!prevVal;
      }
      if (!isApplicable) continue;

      var actualVal = currentPA.actualCol ? sheet.getRange(r, currentPA.actualCol).getValue() : '';
      var status = (actualVal !== '' && actualVal !== null) ? 'completed' : 'pending';

      // Overdue: still pending (Actual empty) AND its Planned date/time has
      // already passed. Never applies to completed rows.
      var isOverdue = false;
      if (status === 'pending' && currentPA.plannedCol) {
        var plannedVal = sheet.getRange(r, currentPA.plannedCol).getValue();
        if (plannedVal instanceof Date && plannedVal.getTime() < Date.now()) {
          isOverdue = true;
        }
      }

      var rowValues = sheet.getRange(r, 1, 1, currentRange.endCol).getValues()[0];
      var cells = [];
      for (var bc2 = 0; bc2 < baseColCount; bc2++) {
        cells.push(formatDateSafe(rowValues[bc2]));
      }
      visibleFields.forEach(function (f) {
        cells.push(formatDateSafe(rowValues[f.col - 1]));
      });

      rows.push({ row: r, status: status, overdue: isOverdue, cells: cells });
    }

    return { success: true, columns: columnNames, columnTypes: columnTypes, defaultHiddenCols: defaultHiddenCols, rows: rows };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================
// STEP FORM FIELDS (for building the popup dynamically)
// ============================================

/**
 * Returns the fields to render in the popup form for a given step + row.
 * "Planned" -> read-only display value (already stored, auto-computed
 *              when the PREVIOUS step was submitted)
 * "Actual"  -> never shown (auto-filled on submit)
 * Others    -> editable inputs; DROPDOWN/CHECKBOX types get their options
 *              from the DROPDOWN sheet (header row split by comma)
 */
function getStepFormFields(masterSheetUrl, masterSheetName, dropdownSheetName, header, stepName, rowNumber) {
  try {
    var target = getTargetSheetInfo(masterSheetUrl, masterSheetName, header);
    if (!target) return { success: false, message: 'Target sheet info not found', fields: [] };

    var targetSs = SpreadsheetApp.openByUrl(target.sheetUrl);
    var sheet = targetSs.getSheetByName(target.sheetName);
    if (!sheet) return { success: false, message: 'Target sheet not found', fields: [] };

    var range = findStepColumnRange(sheet, stepName);
    if (!range) return { success: false, message: 'Step not found in target sheet', fields: [] };

    var fields = getStepFields(sheet, range);
    var formFields = [];

    fields.forEach(function (f) {
      var nameLower = f.name.toLowerCase();
      if (nameLower.indexOf('actual') !== -1) return; // never shown
      if (isAuditTrailField(f.name)) return; // "Modified By/At" are automatic - never a form question

      if (nameLower.indexOf('planned') !== -1) {
        // Always shown, always non-editable - regardless of whether Row 6
        // has a Data Type for it (Planned is auto-computed via TAT).
        formFields.push({
          name: f.name,
          type: 'PLANNED_DISPLAY',
          value: formatDateSafe(sheet.getRange(rowNumber, f.col).getValue()),
          options: []
        });
        return;
      }

      // Any other field with NO Data Type configured in Row 6 is treated
      // as "not set up yet" and must not appear in the form at all.
      if (!f.type) return;

      var field = { name: f.name, type: f.type.toUpperCase(), options: [] };
      if (field.type.indexOf('DROPDOWN') !== -1 || field.type.indexOf('CHECKBOX') !== -1) {
        field.options = getDropdownOptions(masterSheetUrl, dropdownSheetName, f.name);
      }
      formFields.push(field);
    });

    return { success: true, fields: formFields };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message, fields: [] };
  }
}

/**
 * Read DROPDOWN sheet: Row 1 = header(s) - can be comma-separated
 * (e.g. "Vendor Name A,Final Vendor name") meaning multiple field names
 * share the SAME options column. Row 2+ = option values for that column.
 */
function getDropdownOptions(masterSheetUrl, dropdownSheetName, fieldName) {
  try {
    var ss = SpreadsheetApp.openByUrl(masterSheetUrl);
    var sheet = ss.getSheetByName(dropdownSheetName);
    if (!sheet) return [];

    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();
    if (lastCol < 1 || lastRow < 2) return [];

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var targetCol = -1;

    for (var c = 0; c < headers.length; c++) {
      var parts = String(headers[c]).split(',').map(function (p) { return p.trim().toUpperCase(); });
      if (parts.indexOf(fieldName.trim().toUpperCase()) !== -1) { targetCol = c + 1; break; }
    }
    if (targetCol === -1) return [];

    var values = sheet.getRange(2, targetCol, lastRow - 1, 1).getValues();
    var options = [];
    values.forEach(function (row) {
      var v = String(row[0]).trim();
      if (v) options.push(v);
    });
    return options;

  } catch (e) {
    return [];
  }
}

// ============================================
// SUBMIT STEP DATA
// ============================================

/**
 * Save a Step's form data into the target row, mark Actual = now,
 * upload any files to Drive, and auto-compute + write the NEXT step's
 * "Planned" date (via TAT calculator) if a next step exists.
 *
 * fileUploads: { fieldName: [ {base64, fileName, mimeType}, ... ], ... }
 */
function submitStepData(masterSheetUrl, masterSheetName, stepsSheetName, header, stepName, rowNumber, formValues, fileUploads, submittedByUser) {
  try {
    var target = getTargetSheetInfo(masterSheetUrl, masterSheetName, header);
    if (!target) return { success: false, message: 'Target sheet info not found' };

    var targetSs = SpreadsheetApp.openByUrl(target.sheetUrl);
    var sheet = targetSs.getSheetByName(target.sheetName);
    if (!sheet) return { success: false, message: 'Target sheet not found' };

    var range = findStepColumnRange(sheet, stepName);
    if (!range) return { success: false, message: 'Step columns not found' };

    var fields = getStepFields(sheet, range);
    var pa = getPlannedActualCols(fields);

    // Audit trail columns ("Modified By" / "Modified At") - OPTIONAL. If
    // the sheet has columns with these exact names inside this Step's
    // block (Row 7), they get written on every submit. If the sheet
    // doesn't have them, nothing happens (fully backward compatible -
    // no sheet changes required for this to work elsewhere).
    var modifiedByCol = null, modifiedAtCol = null;
    fields.forEach(function (f) {
      var n = f.name.toLowerCase();
      if (n.indexOf('modified by') !== -1) modifiedByCol = f.col;
      if (n.indexOf('modified at') !== -1) modifiedAtCol = f.col;
    });

    fields.forEach(function (f) {
      var nameLower = f.name.toLowerCase();
      if (nameLower.indexOf('planned') !== -1 || nameLower.indexOf('actual') !== -1) return;
      if (nameLower.indexOf('modified by') !== -1 || nameLower.indexOf('modified at') !== -1) return; // written separately below

      if (f.type.toUpperCase().indexOf('FILE') !== -1) {
        var files = (fileUploads && fileUploads[f.name]) || [];
        if (files.length > 0) {
          var urls = files.map(function (file) {
            return uploadFileToDrive(file.base64, file.fileName, file.mimeType);
          });
          sheet.getRange(rowNumber, f.col).setValue(urls.join(', '));
        }
      } else if (formValues && formValues.hasOwnProperty(f.name)) {
        sheet.getRange(rowNumber, f.col).setValue(formValues[f.name]);
      }
    });

    var now = new Date();
    if (pa.actualCol) sheet.getRange(rowNumber, pa.actualCol).setValue(now);
    if (modifiedAtCol) sheet.getRange(rowNumber, modifiedAtCol).setValue(now);
    if (modifiedByCol && submittedByUser) sheet.getRange(rowNumber, modifiedByCol).setValue(submittedByUser);

    // Auto-compute next step's Planned date
    var masterSs = SpreadsheetApp.openByUrl(masterSheetUrl);
    var allSteps = getAllStepsForHeader(masterSs, stepsSheetName, header);
    var stepIndex = -1;
    for (var i = 0; i < allSteps.length; i++) {
      if (allSteps[i].toUpperCase() === stepName.trim().toUpperCase()) { stepIndex = i; break; }
    }

    if (stepIndex !== -1 && stepIndex < allSteps.length - 1) {
      var nextStepName = allSteps[stepIndex + 1];
      var nextRange = findStepColumnRange(sheet, nextStepName);
      if (nextRange) {
        var nextFields = getStepFields(sheet, nextRange);
        var nextPA = getPlannedActualCols(nextFields);
        if (nextPA.plannedCol) {
          var plannedDate = calculateTAT(
            now,
            WF_DEFAULT_TAT_HOURS,
            WF_DEFAULT_OFFICE_START,
            WF_DEFAULT_OFFICE_END,
            WF_DEFAULT_WEEKOFF,
            WF_DEFAULT_HOLIDAY_SHEET,
            masterSheetUrl
          );
          sheet.getRange(rowNumber, nextPA.plannedCol).setValue(plannedDate);
        }
      }
    }

    // Invalidate this user's cached Home summary so their OWN dashboard
    // reflects this submission right away rather than waiting up to 10
    // minutes for the next scheduled refresh. (The client's own optimistic
    // UI update already makes the table feel instant; this keeps the
    // server-side cache from serving stale counts on the next page load.)
    if (typeof fmsCacheRemove === 'function') {
      fmsCacheRemove(fmsCacheKey(['home', submittedByUser]));
    }

    return { success: true, message: 'Saved successfully!' };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================
// FILE UPLOAD TO DRIVE
// ============================================

function uploadFileToDrive(base64Data, fileName, mimeType) {
  var folder = getOrCreateFMSFolder();
  var bytes = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // Sharing might fail depending on domain policy - ignore, file is still saved.
  }
  return file.getUrl();
}

/**
 * Returns the Drive folder to save file-upload question answers into.
 * If FMS_DRIVE_FOLDER_ID (top of Code.gs) is set to a specific Folder ID,
 * that exact folder is used (throws a clear error if the ID is invalid/
 * inaccessible, instead of silently falling back). Otherwise, auto-
 * creates/reuses a folder named "FMS Uploads" in My Drive.
 */
function getOrCreateFMSFolder() {
  if (typeof FMS_DRIVE_FOLDER_ID !== 'undefined' && FMS_DRIVE_FOLDER_ID) {
    return DriveApp.getFolderById(FMS_DRIVE_FOLDER_ID);
  }
  var folderName = 'FMS Uploads';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

// ============================================
// MULTI-STEP "TABLE VIEW" (all of MY steps side by side, one row per item)
// ============================================
//
// The Step View (getStepTableData) shows ONE step at a time - you pick a
// step in the sidebar and work only on that step's columns.
//
// This Table View instead shows the WHOLE pipeline for a header in a
// single wide table: the base identifying columns, then - laid out left
// to right in real pipeline order - the columns of every step THIS user
// is permitted to action (from STEPS sheet Col C). Steps the user has no
// permission for are omitted entirely, so if a doer owns steps 3, 6 and
// 8, they see exactly those three step blocks and nothing else.
//
// Per row, each of the user's steps carries its own state:
//   'completed' - that step's Actual is filled
//   'pending'   - actionable NOW (previous step done, Actual still empty)
//   'locked'    - not reachable yet (an earlier step isn't done)
// and an 'overdue' flag when a pending step's Planned date has passed.
//
// The row-level status (used by the All/Pending/Overdue/Completed tabs)
// is derived from ONLY the user's own steps:
//   pending   - at least one of my steps is actionable now
//   completed - all of my steps on this row are done
//   locked    - none actionable yet (only visible under the "All" tab)
//
function getMultiStepTableData(masterSheetUrl, masterSheetName, stepsSheetName, header, userName) {
  try {
    var target = getTargetSheetInfo(masterSheetUrl, masterSheetName, header);
    if (!target) return { success: false, message: 'Target sheet info not found in MASTER for header: ' + header };

    var targetSs = SpreadsheetApp.openByUrl(target.sheetUrl);
    var sheet = targetSs.getSheetByName(target.sheetName);
    if (!sheet) return { success: false, message: 'Target sheet "' + target.sheetName + '" not found' };

    var masterSs = SpreadsheetApp.openByUrl(masterSheetUrl);
    var allSteps = getAllStepsForHeader(masterSs, stepsSheetName, header);
    if (!allSteps.length) return { success: false, message: 'No steps found in STEPS sheet for header: ' + header };

    // Which of those steps may THIS user action? (same permission source
    // as the sidebar - STEPS sheet Col C)
    var perms = getUserPermissionsInternal(masterSheetUrl, stepsSheetName, userName);
    var myStepNames = [];
    if (perms.success) {
      if (perms.permissions[header]) {
        myStepNames = perms.permissions[header];
      } else {
        for (var hk in perms.permissions) {
          if (String(hk).trim().toUpperCase() === String(header).trim().toUpperCase()) {
            myStepNames = perms.permissions[hk];
            break;
          }
        }
      }
    }
    var myStepsUpper = myStepNames.map(function (s) { return String(s).trim().toUpperCase(); });

    // ---- Base (identifying) columns: everything before the FIRST step ----
    var firstStepRange = findStepColumnRange(sheet, allSteps[0]);
    var baseColCount = firstStepRange ? Math.max(0, firstStepRange.startCol - 1) : 0;
    var baseNameVals = baseColCount ? sheet.getRange(WF_FIELD_NAME_ROW, 1, 1, baseColCount).getValues()[0] : [];
    var baseTypeVals = baseColCount ? sheet.getRange(WF_TYPE_ROW, 1, 1, baseColCount).getValues()[0] : [];

    var columns = [];          // flat list of column labels
    var columnTypes = [];      // parallel: 'DATE' etc (drives the calendar filter)
    var columnGroups = [];     // parallel: '' for base cols, else the step name
    var defaultHiddenCols = []; // audit-trail cols start hidden
    var lockedCols = [];       // Plan-date cols can NEVER be hidden (per requirement)

    for (var bc = 0; bc < baseColCount; bc++) {
      var bName = String(baseNameVals[bc]).trim() || ('Col ' + (bc + 1));
      var bType = String(baseTypeVals[bc]).trim().toUpperCase();
      var bLower = bName.toLowerCase();
      var bIsDate = bType.indexOf('DATE') !== -1 || bLower.indexOf('date') !== -1 || bLower.indexOf('timestamp') !== -1;
      columns.push(bName);
      columnTypes.push(bIsDate ? 'DATE' : bType);
      columnGroups.push('');
      if (isAuditTrailField(bName)) defaultHiddenCols.push(columns.length - 1);
    }

    // ---- One block of columns per step (only for the user's own steps) ----
    var stepMeta = [];
    for (var i = 0; i < allSteps.length; i++) {
      var stepName = allSteps[i];
      var range = findStepColumnRange(sheet, stepName);
      if (!range) continue;

      var fields = getStepFields(sheet, range);
      var pa = getPlannedActualCols(fields);
      var visibleFields = fields.filter(isFieldConfigured);
      var allowed = myStepsUpper.indexOf(String(stepName).trim().toUpperCase()) !== -1;

      var colIdxStart = -1, colIdxEnd = -1;
      if (allowed) {
        colIdxStart = columns.length;
        for (var f = 0; f < visibleFields.length; f++) {
          var fld = visibleFields[f];
          var fName = fld.name || ('Col ' + fld.col);
          var fLower = fName.toLowerCase();
          var fIsDate = fld.type.toUpperCase().indexOf('DATE') !== -1 ||
            fLower.indexOf('planned') !== -1 || fLower.indexOf('actual') !== -1 ||
            fLower.indexOf('modified at') !== -1;
          columns.push(fName);
          columnTypes.push(fIsDate ? 'DATE' : fld.type.toUpperCase());
          columnGroups.push(stepName);
          var newIdx = columns.length - 1;
          if (isAuditTrailField(fName)) defaultHiddenCols.push(newIdx);
          // Plan dates must always stay visible - they're the whole point
          // of the wide view (seeing what's due across the pipeline).
          if (fLower.indexOf('planned') !== -1) lockedCols.push(newIdx);
        }
        colIdxEnd = columns.length - 1;
      }

      stepMeta.push({
        step: stepName,
        order: i + 1,
        allowed: allowed,
        plannedCol: pa.plannedCol,
        actualCol: pa.actualCol,
        fields: visibleFields,
        colIdxStart: colIdxStart,
        colIdxEnd: colIdxEnd,
        prevActualCol: null
      });
    }

    // Each step's gate is the PREVIOUS pipeline step's Actual column
    // (regardless of whether the user owns that previous step).
    for (var s = 0; s < stepMeta.length; s++) {
      stepMeta[s].prevActualCol = s > 0 ? stepMeta[s - 1].actualCol : null;
    }

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var rows = [];
    var nowMs = Date.now();

    for (var r = WF_DATA_START_ROW; r <= lastRow; r++) {
      var rowValues = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
      if (!rowValues[0]) continue; // no base data -> not a real work item

      var cells = [];
      for (var b2 = 0; b2 < baseColCount; b2++) cells.push(formatDateSafe(rowValues[b2]));

      var stepStates = [];
      var anyPending = false, anyOverdue = false, myTotal = 0, myDone = 0;

      for (var m = 0; m < stepMeta.length; m++) {
        var sm = stepMeta[m];
        var actualVal = sm.actualCol ? rowValues[sm.actualCol - 1] : '';
        var isDone = (actualVal !== '' && actualVal !== null && actualVal !== undefined);
        var prevDone = sm.prevActualCol ? !!rowValues[sm.prevActualCol - 1] : true;
        var state = isDone ? 'completed' : (prevDone ? 'pending' : 'locked');

        var isOverdue = false;
        if (state === 'pending' && sm.plannedCol) {
          var plannedVal = rowValues[sm.plannedCol - 1];
          if (plannedVal instanceof Date && plannedVal.getTime() < nowMs) isOverdue = true;
        }

        if (sm.allowed) {
          for (var vf = 0; vf < sm.fields.length; vf++) {
            cells.push(formatDateSafe(rowValues[sm.fields[vf].col - 1]));
          }
          myTotal++;
          if (state === 'completed') myDone++;
          if (state === 'pending') anyPending = true;
          if (isOverdue) anyOverdue = true;

          stepStates.push({ step: sm.step, order: sm.order, status: state, overdue: isOverdue });
        }
      }

      var rowStatus = anyPending ? 'pending' : ((myTotal > 0 && myDone === myTotal) ? 'completed' : 'locked');
      rows.push({ row: r, status: rowStatus, overdue: anyOverdue, cells: cells, steps: stepStates });
    }

    return {
      success: true,
      columns: columns,
      columnTypes: columnTypes,
      columnGroups: columnGroups,
      defaultHiddenCols: defaultHiddenCols,
      lockedCols: lockedCols,
      baseColCount: baseColCount,
      steps: stepMeta.filter(function (s) { return s.allowed; }).map(function (s) {
        return { step: s.step, order: s.order, colIdxStart: s.colIdxStart, colIdxEnd: s.colIdxEnd };
      }),
      totalPipelineSteps: allSteps.length,
      rows: rows
    };

  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================
// CLIENT-FACING WRAPPERS (called from Dashboard via google.script.run)
// ============================================

function wfGetPendingRows(sheetUrl, masterSheet, stepsSheet, header, stepName) {
  return getPendingRowsForStep(sheetUrl, masterSheet, stepsSheet, header, stepName);
}

function wfGetStepForm(sheetUrl, masterSheet, dropdownSheet, header, stepName, rowNumber) {
  return getStepFormFields(sheetUrl, masterSheet, dropdownSheet, header, stepName, rowNumber);
}

function wfSubmitStep(sheetUrl, masterSheet, stepsSheet, header, stepName, rowNumber, formValues, fileUploads, submittedByUser) {
  return submitStepData(sheetUrl, masterSheet, stepsSheet, header, stepName, rowNumber, formValues, fileUploads, submittedByUser);
}

function wfGetStepTableData(sheetUrl, masterSheet, stepsSheet, header, stepName) {
  return getStepTableData(sheetUrl, masterSheet, stepsSheet, header, stepName);
}

function wfGetHomeSummary(sheetUrl, masterSheet, stepsSheet, userName) {
  return getHomeSummary(sheetUrl, masterSheet, stepsSheet, userName);
}

function wfGetMultiStepTableData(sheetUrl, masterSheet, stepsSheet, header, userName) {
  return getMultiStepTableData(sheetUrl, masterSheet, stepsSheet, header, userName);
}
