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
  if (val instanceof Date) {
    return formatResult(val); // reuses formatResult() from TAT_Calculator.gs
  }
  return String(val);
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

    var columnNames = [];
    for (var bc = 0; bc < baseColCount; bc++) {
      columnNames.push(String(headerRowVals[bc]).trim() || ('Col ' + (bc + 1)));
    }
    currentFields.forEach(function (f) {
      columnNames.push(f.name || ('Col ' + f.col));
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

      var rowValues = sheet.getRange(r, 1, 1, currentRange.endCol).getValues()[0];
      var cells = [];
      for (var bc2 = 0; bc2 < baseColCount; bc2++) {
        cells.push(formatDateSafe(rowValues[bc2]));
      }
      currentFields.forEach(function (f) {
        cells.push(formatDateSafe(rowValues[f.col - 1]));
      });

      rows.push({ row: r, status: status, cells: cells });
    }

    return { success: true, columns: columnNames, rows: rows };

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

      if (nameLower.indexOf('planned') !== -1) {
        formFields.push({
          name: f.name,
          type: 'PLANNED_DISPLAY',
          value: formatDateSafe(sheet.getRange(rowNumber, f.col).getValue()),
          options: []
        });
        return;
      }

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
function submitStepData(masterSheetUrl, masterSheetName, stepsSheetName, header, stepName, rowNumber, formValues, fileUploads) {
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

    fields.forEach(function (f) {
      var nameLower = f.name.toLowerCase();
      if (nameLower.indexOf('planned') !== -1 || nameLower.indexOf('actual') !== -1) return;

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

function getOrCreateFMSFolder() {
  var folderName = 'FMS Uploads';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
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

function wfSubmitStep(sheetUrl, masterSheet, stepsSheet, header, stepName, rowNumber, formValues, fileUploads) {
  return submitStepData(sheetUrl, masterSheet, stepsSheet, header, stepName, rowNumber, formValues, fileUploads);
}

function wfGetStepTableData(sheetUrl, masterSheet, stepsSheet, header, stepName) {
  return getStepTableData(sheetUrl, masterSheet, stepsSheet, header, stepName);
}
