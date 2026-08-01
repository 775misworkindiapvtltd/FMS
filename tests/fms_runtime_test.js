'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
const read = (name) => fs.readFileSync(path.join(ROOT, name), 'utf8');
const serverFiles = ['Code.gs', 'WorkflowEngine.gs'];
const optionalServerFiles = ['TAT_Calculator.gs'];
const allServerFiles = serverFiles.concat(optionalServerFiles);
const htmlFiles = ['Login.html', 'Dashboard.html', 'Scripts.html', 'Styles.html', 'ThemeEngine.html'];
const htmlSource = Object.fromEntries(htmlFiles.map((name) => [name, read(name)]));

let passed = 0;
function test(name, fn) {
  try {
    fn();
    passed += 1;
    console.log('PASS  ' + name);
  } catch (error) {
    console.error('FAIL  ' + name);
    throw error;
  }
}

class MockRange {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getValues() {
    const result = [];
    for (let r = 0; r < this.numRows; r += 1) {
      const row = [];
      for (let c = 0; c < this.numCols; c += 1) {
        row.push(this.sheet.valueAt(this.row + r, this.col + c));
      }
      result.push(row);
    }
    return result;
  }
  getValue() { return this.sheet.valueAt(this.row, this.col); }
  setValue(value) { this.sheet.setValueAt(this.row, this.col, value); return this; }
  getMergedRanges() { return []; }
  getColumn() { return this.col; }
  getNumColumns() { return this.numCols; }
}

class MockSheet {
  constructor(name, values) {
    this.name = name;
    this.values = values.map((row) => row.slice());
  }
  valueAt(row, col) {
    return this.values[row - 1] && this.values[row - 1][col - 1] !== undefined
      ? this.values[row - 1][col - 1]
      : '';
  }
  setValueAt(row, col, value) {
    while (this.values.length < row) this.values.push([]);
    while (this.values[row - 1].length < col) this.values[row - 1].push('');
    this.values[row - 1][col - 1] = value;
  }
  getDataRange() { return new MockRange(this, 1, 1, this.getLastRow(), this.getLastColumn()); }
  getRange(row, col, numRows, numCols) { return new MockRange(this, row, col, numRows, numCols); }
  getLastRow() { return this.values.length; }
  getLastColumn() { return this.values.reduce((max, row) => Math.max(max, row.length), 0); }
}

class MockSpreadsheet {
  constructor(name, sheets) {
    this.name = name;
    this.sheets = sheets;
  }
  getName() { return this.name; }
  getSheetByName(name) { return this.sheets[name] || null; }
}

function makeOutput(content) {
  return {
    content,
    title: '',
    getContent() { return this.content; },
    setTitle(value) { this.title = value; return this; },
    setXFrameOptionsMode() { return this; },
    addMetaTag() { return this; }
  };
}

const cache = new Map();
let openCount = 0;
let context;
const triggerList = [];

const masterUrl = 'https://docs.google.com/spreadsheets/d/10Aqav9bM_XQ28TkfjoXOoj8jRHenMerIEcx8gYkwo48/edit?usp=drive_web&ouid=114024121405657848674';
const targetUrl = 'mock://flow-a';
const past = new Date(2025, 0, 20, 10, 0, 0);
const completed = new Date(2025, 0, 20, 11, 0, 0);

const masterSheet = new MockSheet('MASTER', [
  ['', '', targetUrl],
  ['', '', 'Flow A'],
  ['', '', ''],
  ['LOGIN', 'PASSWORD', 'Name'],
  ['USER1', 'pass', 'User One']
]);
const stepsSheet = new MockSheet('STEPS', [
  ['Flow A', 'Step 1', 'USER1'],
  ['Flow A', 'Step 2', 'USER1']
]);
const dropdownSheet = new MockSheet('DROPDPWN', [
  ['Note,Result'],
  ['Approved'],
  ['Rejected']
]);
const targetSheet = new MockSheet('Flow A', [
  ['', '', '', '', '', '', ''],
  ['', 'Step 1', '', '', 'Step 2', '', ''],
  ['', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['', '', '', '', '', '', ''],
  ['TEXT', 'DATE', 'DATE', 'LONG TEXT', 'DATE', 'DATE', 'DROPDOWN'],
  ['Item', 'Planned', 'Actual', 'Note', 'Planned', 'Actual', 'Result'],
  ['Item A', past, '', '', '', '', ''],
  ['Item B', past, completed, 'done', past, '', '']
]);

const spreadsheets = {
  [masterUrl]: new MockSpreadsheet('FMS Master', {
    MASTER: masterSheet,
    STEPS: stepsSheet,
    DROPDPWN: dropdownSheet
  }),
  [targetUrl]: new MockSpreadsheet('Flow A Book', { 'Flow A': targetSheet })
};

function renderTemplate(name) {
  let source = htmlSource[name + '.html'];
  if (source === undefined) throw new Error('Missing HTML file: ' + name);
  source = source.replace(/<\?!=\s*include\(['"]([^'"]+)['"]\)\s*\?>/g, (_, includeName) => renderTemplate(includeName));
  source = source.replace(/<\?=\s*FMS_BUILD_ID\s*\?>/g, () => String(context.FMS_BUILD_ID));
  return source;
}

const sandbox = {
  console,
  Date,
  JSON,
  Math,
  String,
  Number,
  Boolean,
  Object,
  Array,
  RegExp,
  Error,
  parseInt,
  parseFloat,
  isNaN,
  isFinite,
  setTimeout,
  clearTimeout,
  Logger: { log() {} },
  HtmlService: {
    XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
    createHtmlOutput(content) { return makeOutput(content); },
    createHtmlOutputFromFile(name) {
      const source = htmlSource[name + '.html'];
      if (source === undefined) throw new Error('Missing HTML file: ' + name);
      return makeOutput(source);
    },
    createTemplateFromFile(name) {
      if (htmlSource[name + '.html'] === undefined) throw new Error('Missing HTML file: ' + name);
      return { evaluate: () => makeOutput(renderTemplate(name)) };
    }
  },
  SpreadsheetApp: {
    openByUrl(url) {
      openCount += 1;
      if (!spreadsheets[url]) throw new Error('Spreadsheet not found: ' + url);
      return spreadsheets[url];
    }
  },
  CacheService: {
    getScriptCache() {
      return {
        get(key) { return cache.has(key) ? cache.get(key) : null; },
        put(key, value) { cache.set(key, value); },
        remove(key) { cache.delete(key); }
      };
    }
  },
  ScriptApp: {
    getProjectTriggers() { return triggerList; },
    getService() { return { getUrl: () => 'https://script.google.com/macros/s/test/exec' }; },
    newTrigger(handler) {
      const trigger = { getHandlerFunction: () => handler };
      return {
        timeBased() { return this; },
        everyMinutes(minutes) { assert.strictEqual(minutes, 10); return this; },
        create() { triggerList.push(trigger); return trigger; }
      };
    }
  },
  Utilities: {
    base64Decode(value) { return Buffer.from(value, 'base64'); },
    newBlob(bytes, mimeType, fileName) { return { bytes, mimeType, fileName }; }
  },
  DriveApp: {
    Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK' },
    Permission: { VIEW: 'VIEW' },
    getFoldersByName() { return { hasNext: () => false }; },
    createFolder() { return { createFile: () => ({ setSharing() {}, getUrl: () => 'mock://file' }) }; },
    getFolderById() { return this.createFolder(); }
  }
};

context = vm.createContext(sandbox);
serverFiles.forEach((name) => new vm.Script(read(name), { filename: name }).runInContext(context));

const tableContainerElement = { innerHTML: '' };
const clientDocument = {
  addEventListener() {},
  getElementById(id) { return id === 'tableContainer' ? tableContainerElement : null; },
  querySelector() { return null; },
  querySelectorAll() { return []; },
  createElement() {
    let escaped = '';
    return {
      set textContent(value) {
        escaped = String(value == null ? '' : value)
          .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
          .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
      },
      get innerHTML() { return escaped; }
    };
  },
  body: { classList: { add() {}, remove() {}, toggle() {} }, insertAdjacentHTML() {} }
};
const clientStorage = new Map();
const clientWindow = {
  localStorage: {
    getItem(key) { return clientStorage.has(key) ? clientStorage.get(key) : null; },
    setItem(key, value) { clientStorage.set(key, String(value)); }
  },
  matchMedia() { return { matches: false }; },
  innerWidth: 1440,
  innerHeight: 900,
  addEventListener() {}
};
const clientContext = vm.createContext({
  console, Date, JSON, Math, String, Number, Boolean, Object, Array, RegExp, Error, URL,
  parseInt, parseFloat, isNaN, isFinite, setTimeout, clearTimeout, setInterval, clearInterval,
  document: clientDocument, window: clientWindow
});
const scriptsBody = htmlSource['Scripts.html'].match(/^<script>([\s\S]*)<\/script>\s*$/i)[1];
new vm.Script(scriptsBody, { filename: 'Scripts.html' }).runInContext(clientContext);

test('all Apps Script server files parse; production loads without optional TAT_Calculator.gs', () => {
  allServerFiles.forEach((name) => {
    assert.ok(read(name).length > 0);
    new vm.Script(read(name), { filename: name + ':parse-only' });
  });
  assert.strictEqual(typeof context.calculateTAT, 'undefined');
  assert.strictEqual(typeof context.wfCalculateTAT, 'function');
});

test('all inline HTML JavaScript blocks parse', () => {
  htmlFiles.forEach((name) => {
    const source = htmlSource[name].replace(/<\?[^?]*\?>/g, '');
    const blocks = [...source.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)];
    blocks.forEach((match, index) => new vm.Script(match[1], { filename: name + ':script-' + index }));
  });
});

test('Styles.html has balanced braces', () => {
  const css = htmlSource['Styles.html'].replace(/\/\*[\s\S]*?\*\//g, '');
  let depth = 0;
  for (const char of css) {
    if (char === '{') depth += 1;
    if (char === '}') depth -= 1;
    assert.ok(depth >= 0, 'closing brace before opening brace');
  }
  assert.strictEqual(depth, 0);
});

test('server global function declarations are unique', () => {
  const seen = new Map();
  allServerFiles.forEach((name) => {
    const matches = read(name).matchAll(/^function\s+([A-Za-z_$][\w$]*)\s*\(/gm);
    for (const match of matches) {
      assert.ok(!seen.has(match[1]), 'duplicate function ' + match[1] + ' in ' + seen.get(match[1]) + ' and ' + name);
      seen.set(match[1], name);
    }
  });
  assert.strictEqual(seen.get('doGet'), 'Code.gs');
});

test('deployment contract reports a complete runtime', () => {
  const info = context.fmsGetRuntimeInfo();
  assert.strictEqual(info.success, true);
  assert.deepStrictEqual(Array.from(info.missing), []);
  assert.strictEqual(info.buildId, '2026-07-22-column-width-2');
});

test('deployment contract is ready when calculateTAT and TAT_Calculator.gs are absent', () => {
  assert.strictEqual(typeof context.calculateTAT, 'undefined');
  const info = context.fmsGetRuntimeInfo();
  assert.strictEqual(info.success, true);
  assert.ok(!Array.from(info.missing).some((item) => item.indexOf('calculateTAT') !== -1));
  assert.match(context.doGet({ parameter: { fms_check: '1' } }).getContent(), /FMS deployment is ready/);
});

test('doGet renders FMS Login and explicit deployment status', () => {
  const login = context.doGet({ parameter: {} });
  assert.strictEqual(login.title, 'FMS - Flow Management System');
  assert.match(login.getContent(), /Flow Management System/);
  assert.match(login.getContent(), /Build 2026-07-22-column-width-2/);
  const status = context.doGet({ parameter: { fms_check: '1' } });
  assert.match(status.getContent(), /FMS deployment is ready/);
});

test('incomplete deployment is stopped before login', () => {
  const original = context.wfSubmitStep;
  context.wfSubmitStep = undefined;
  const info = context.fmsGetRuntimeInfo();
  assert.strictEqual(info.success, false);
  assert.ok(Array.from(info.missing).includes('wfSubmitStep()'));
  const output = context.doGet({ parameter: {} }).getContent();
  assert.match(output, /deployment is incomplete/);
  assert.match(output, /wfSubmitStep/);
  context.wfSubmitStep = original;
});

test('deployment preflight evaluates Dashboard includes before login', () => {
  const styles = htmlSource['Styles.html'];
  delete htmlSource['Styles.html'];
  const info = context.fmsGetRuntimeInfo();
  assert.strictEqual(info.success, false);
  assert.ok(Array.from(info.missing).some((item) => item.indexOf('Dashboard.html render') === 0));
  htmlSource['Styles.html'] = styles;
  assert.strictEqual(context.fmsGetRuntimeInfo().success, true);
});

test('login validates credentials and returns deployment identity', () => {
  const ok = context.validateLogin('user1', 'pass');
  assert.strictEqual(ok.success, true);
  assert.strictEqual(ok.loginId, 'USER1');
  assert.strictEqual(ok.buildId, context.FMS_BUILD_ID);
  assert.strictEqual(context.validateLogin('USER1', 'wrong').success, false);
  assert.strictEqual(context.validateLogin('', '').success, false);
});

test('dashboard template expands all required includes', () => {
  const output = context.getDashboardHtml();
  assert.match(output, /FMS - GLOBAL SCRIPTS/);
  assert.match(output, /PROFESSIONAL APPLICATION HEADER/);
  assert.doesNotMatch(output, /<\?!=/);
});

test('permissions cache hit performs zero additional Sheet reads', () => {
  cache.clear();
  openCount = 0;
  const first = context.getUserPermissions('USER1');
  assert.strictEqual(first.success, true);
  assert.deepStrictEqual(Array.from(first.permissions['Flow A']), ['Step 1', 'Step 2']);
  const readsAfterMiss = openCount;
  const second = context.getUserPermissions('USER1');
  assert.strictEqual(second.success, true);
  assert.strictEqual(openCount, readsAfterMiss);
});

test('home cache miss computes counts and cache hit performs zero reads', () => {
  cache.clear();
  openCount = 0;
  const first = context.wfGetHomeSummary(masterUrl, 'MASTER', 'STEPS', 'USER1');
  assert.strictEqual(first.success, true);
  assert.strictEqual(first.totalPending, 2);
  assert.strictEqual(first.totalCompleted, 1);
  const readsAfterMiss = openCount;
  const second = context.wfGetHomeSummary(masterUrl, 'MASTER', 'STEPS', 'USER1');
  assert.strictEqual(second.success, true);
  assert.strictEqual(openCount, readsAfterMiss);
});

test('step and multi-step tables load with pending/completed/locked state', () => {
  const stepOne = context.wfGetStepTableData(masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 1');
  assert.strictEqual(stepOne.success, true);
  assert.deepStrictEqual(Array.from(stepOne.columns), ['Item', 'Planned', 'Actual', 'Note']);
  assert.deepStrictEqual(Array.from(stepOne.rows, (row) => row.status), ['pending', 'completed']);
  assert.ok(Array.isArray(stepOne.formSchema));
  assert.match(stepOne.formSchemaVersion, /^schema-[0-9a-f]+$/);
  assert.ok(Array.from(stepOne.formSchema).some((field) => field.type === 'PLANNED_DISPLAY'));

  const multi = context.wfGetMultiStepTableData(masterUrl, 'MASTER', 'STEPS', 'Flow A', 'USER1');
  assert.strictEqual(multi.success, true);
  assert.strictEqual(multi.steps.length, 2);
  assert.ok(multi.formSchemas && Array.isArray(multi.formSchemas['Step 1']));
  assert.ok(Array.isArray(multi.formSchemas['Step 2']));
  assert.strictEqual(multi.formSchemaVersions['Step 1'], stepOne.formSchemaVersion);
  assert.match(multi.formSchemaVersions['Step 2'], /^schema-[0-9a-f]+$/);
  assert.deepStrictEqual(Array.from(multi.rows[0].steps, (step) => step.status), ['pending', 'locked']);
  assert.deepStrictEqual(Array.from(multi.rows[1].steps, (step) => step.status), ['completed', 'pending']);
});

test('table and form date rendering works with formatResult removed', () => {
  const original = context.formatResult;
  delete context.formatResult;
  const formatted = context.formatDateSafe(new Date(2026, 6, 24, 17, 43, 32));
  assert.strictEqual(formatted, '24 Jul 26 17:43:32');
  const dateOnly = context.formatDateSafe(new Date(2026, 6, 24, 0, 0, 0));
  assert.strictEqual(dateOnly, '24 Jul 26');
  const table = context.wfGetStepTableData(masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 1');
  assert.strictEqual(table.success, true);
  const form = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 1', 8, 'STEPS');
  assert.strictEqual(form.success, true);
  assert.match(form.schemaVersion, /^schema-[0-9a-f]+$/);
  assert.ok(Array.from(form.fields).some((field) => field.type === 'PLANNED_DISPLAY'));
  assert.deepStrictEqual(
    Array.from(form.context, (detail) => ({ name: detail.name, value: detail.value })),
    [{ name: 'Item', value: 'Item A' }]
  );
  context.formatResult = original;
});

test('self-contained workflow TAT matches office-hours, week-off and holiday rules', () => {
  const sameDay = context.wfCalculateTAT(
    new Date(2025, 0, 20, 16, 0), 2, '10:00', '18:30', [0], '', ''
  );
  assert.strictEqual(context.formatDateSafe(sameDay), '20 Jan 25 18:00:00');

  const carryNextDay = context.wfCalculateTAT(
    new Date(2025, 0, 20, 17, 30), 2, '10:00', '18:30', [0], '', ''
  );
  assert.strictEqual(context.formatDateSafe(carryNextDay), '21 Jan 25 11:00:00');

  const skipWeekend = context.wfCalculateTAT(
    new Date(2025, 0, 25, 15, 0), 2, '10:00', '18:30', [0, 6], '', ''
  );
  assert.strictEqual(context.formatDateSafe(skipWeekend), '27 Jan 25 12:00:00');

  assert.throws(
    () => context.wfCalculateTAT(new Date(), 2, '18:30', '10:00', [0], '', ''),
    /Office end time must be after office start time/
  );
});

test('server rejects stale schemas and newly-required FILE fields before writing', () => {
  cache.clear();
  const before = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 1', 8, 'STEPS');
  assert.strictEqual(before.success, true);
  assert.match(before.schemaVersion, /^schema-/);

  const originalType = targetSheet.valueAt(6, 4);
  const originalName = targetSheet.valueAt(7, 4);
  const originalAnswer = targetSheet.valueAt(8, 4);
  const originalActual = targetSheet.valueAt(8, 3);
  targetSheet.setValueAt(6, 4, 'FILE');
  targetSheet.setValueAt(7, 4, 'Proof "Image"');

  const current = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 1', 8, 'STEPS');
  assert.strictEqual(current.success, true);
  assert.notStrictEqual(current.schemaVersion, before.schemaVersion);
  assert.ok(Array.from(current.fields).some((field) => field.name === 'Proof "Image"' && field.type === 'FILE'));

  const stale = context.wfSubmitStep(
    masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 1', 8,
    { Note: 'must not write' }, {}, 'USER1', before.schemaVersion
  );
  assert.strictEqual(stale.success, false);
  assert.strictEqual(stale.code, 'FORM_SCHEMA_OUTDATED');
  assert.strictEqual(targetSheet.valueAt(8, 4), originalAnswer);
  assert.strictEqual(targetSheet.valueAt(8, 3), originalActual);

  const missingFile = context.wfSubmitStep(
    masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 1', 8,
    {}, {}, 'USER1', current.schemaVersion
  );
  assert.strictEqual(missingFile.success, false);
  assert.strictEqual(missingFile.code, 'FILE_UPLOAD_REQUIRED');
  assert.match(missingFile.message, /Proof "Image"/);
  assert.strictEqual(targetSheet.valueAt(8, 3), originalActual);

  const originalPlannedType = targetSheet.valueAt(6, 2);
  const originalPlannedName = targetSheet.valueAt(7, 2);
  const originalPlannedValue = targetSheet.valueAt(8, 2);
  targetSheet.setValueAt(6, 2, 'TEXT');
  targetSheet.setValueAt(7, 2, 'Comment');
  const malformedSchema = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 1', 8, 'STEPS');
  const malformedFile = context.wfSubmitStep(
    masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 1', 8,
    { Comment: 'must not be partially written' },
    { 'Proof "Image"': [{ base64: 'not valid base64', fileName: 'proof.txt', mimeType: 'text/plain' }] },
    'USER1', malformedSchema.schemaVersion
  );
  assert.strictEqual(malformedFile.success, false);
  assert.strictEqual(malformedFile.code, 'FILE_UPLOAD_INVALID');
  assert.strictEqual(targetSheet.valueAt(8, 2), originalPlannedValue);
  assert.strictEqual(targetSheet.valueAt(8, 3), originalActual);

  targetSheet.setValueAt(6, 2, originalPlannedType);
  targetSheet.setValueAt(7, 2, originalPlannedName);
  targetSheet.setValueAt(6, 4, originalType);
  targetSheet.setValueAt(7, 4, originalName);
  cache.clear();
});

test('form schema and dropdown options use cache on repeated loads', () => {
  cache.clear();
  openCount = 0;
  const first = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 2', 9, 'STEPS');
  assert.strictEqual(first.success, true);
  const firstReads = openCount;
  const second = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 2', 9, 'STEPS');
  assert.strictEqual(second.success, true);
  const secondReads = openCount - firstReads;
  assert.strictEqual(
    JSON.stringify(Array.from(second.fields, (field) => Array.from(field.options || []))),
    JSON.stringify(Array.from(first.fields, (field) => Array.from(field.options || [])))
  );
  assert.ok(secondReads < firstReads, 'expected cached schema/dropdown load to use fewer spreadsheet opens');
});

test('server rechecks live dropdown options before any write', () => {
  cache.clear();
  const before = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 2', 9, 'STEPS');
  const originalOption = dropdownSheet.valueAt(3, 1);
  const originalResult = targetSheet.valueAt(9, 7);
  const originalActual = targetSheet.valueAt(9, 6);
  dropdownSheet.setValueAt(3, 1, 'On Hold');

  const stale = context.wfSubmitStep(
    masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 2', 9,
    { Result: 'Rejected' }, {}, 'USER1', before.schemaVersion
  );
  assert.strictEqual(stale.success, false);
  assert.strictEqual(stale.code, 'FORM_SCHEMA_OUTDATED');
  assert.strictEqual(targetSheet.valueAt(9, 7), originalResult);
  assert.strictEqual(targetSheet.valueAt(9, 6), originalActual);

  const current = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 2', 9, 'STEPS');
  assert.notStrictEqual(current.schemaVersion, before.schemaVersion);
  const invalid = context.wfSubmitStep(
    masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 2', 9,
    { Result: 'Rejected' }, {}, 'USER1', current.schemaVersion
  );
  assert.strictEqual(invalid.success, false);
  assert.strictEqual(invalid.code, 'INVALID_FORM_OPTION');
  assert.strictEqual(targetSheet.valueAt(9, 7), originalResult);
  assert.strictEqual(targetSheet.valueAt(9, 6), originalActual);

  dropdownSheet.setValueAt(3, 1, originalOption);
  cache.clear();

  const restored = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 2', 9, 'STEPS');
  const trimmed = context.wfSubmitStep(
    masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 2', 9,
    { Result: '  Approved  ' }, {}, 'USER1', restored.schemaVersion
  );
  assert.strictEqual(trimmed.success, true, trimmed.message);
  assert.strictEqual(targetSheet.valueAt(9, 7), 'Approved');
  targetSheet.setValueAt(9, 7, originalResult);
  targetSheet.setValueAt(9, 6, originalActual);
  cache.clear();
});

test('edit/submit writes fields, Actual and next Planned without calculateTAT global', () => {
  const homeKey = context.fmsCacheKey(['home', 'USER1']);
  cache.set(homeKey, JSON.stringify({ success: true, stale: true }));
  const form = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 1', 8, 'STEPS');
  assert.strictEqual(form.success, true);
  const result = context.wfSubmitStep(
    masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 1', 8,
    { Note: 'Edited from form' }, {}, 'USER1', form.schemaVersion
  );
  assert.strictEqual(result.success, true, result.message);
  assert.strictEqual(targetSheet.valueAt(8, 4), 'Edited from form');
  assert.ok(targetSheet.valueAt(8, 3) instanceof Date, 'Actual was not written');
  assert.ok(targetSheet.valueAt(8, 5) instanceof Date, 'next Planned was not written');
  assert.strictEqual(cache.has(homeKey), false);
  const refreshed = context.wfGetStepTableData(masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 1');
  assert.strictEqual(refreshed.rows[0].status, 'completed');
});

test('10-minute trigger installs once and replaces stale cached data', () => {
  triggerList.length = 0;
  const homeKey = context.fmsCacheKey(['home', 'USER1']);
  const permsKey = context.fmsCacheKey(['perms', 'USER1']);
  cache.set(homeKey, JSON.stringify({ success: true, stale: true }));
  cache.set(permsKey, JSON.stringify({ success: true, stale: true }));

  const installed = context.installFmsAutoRefreshTrigger();
  assert.match(installed, /Installed!/);
  assert.strictEqual(triggerList.length, 1);
  assert.strictEqual(context.isFmsAutoRefreshTriggerInstalled(), true);
  assert.strictEqual(JSON.parse(cache.get(homeKey)).stale, undefined);
  assert.strictEqual(JSON.parse(cache.get(permsKey)).stale, undefined);

  const again = context.installFmsAutoRefreshTrigger();
  assert.match(again, /Already installed/);
  assert.strictEqual(triggerList.length, 1);
});

test('client renders one-line numeric state chips in a compact Steps column', () => {
  const fms = clientContext.FMS;
  fms.viewMode = 'table';
  fms.table.columns = ['Item'];
  fms.table.columnTypes = ['TEXT'];
  fms.table.columnGroups = [''];
  fms.table.hiddenCols = {};
  fms.table.colWidths = {};
  fms.table.columnPanelOpen = false;
  fms.table.filterStatus = 'all';
  fms.table.globalSearch = '';
  fms.table.colSearch = {};
  fms.table.sortCol = null;
  fms.table.page = 1;
  fms.table.pageSize = 50;
  fms.table.steps = [
    { step: 'Review', order: 1 },
    { step: 'Approve', order: 2 },
    { step: 'Dispatch', order: 3 }
  ];
  fms.table.rows = [{
    row: 8,
    status: 'pending',
    overdue: false,
    cells: ['Item A'],
    steps: [
      { step: 'Review', order: 1, status: 'pending', overdue: false },
      { step: 'Approve', order: 2, status: 'completed', overdue: false },
      { step: 'Dispatch', order: 3, status: 'locked', overdue: false }
    ]
  }];

  clientContext.renderMultiStepTable('Flow A');
  const html = tableContainerElement.innerHTML;
  assert.match(html, />Steps</);
  assert.match(html, /chip-pending[^>]*[\s\S]*?>1<\/button>/);
  assert.match(html, /chip-done[^>]*>2<\/span>/);
  assert.match(html, /chip-locked[^>]*>3<\/span>/);
  assert.match(html, /data-col="__action" style="width:78px/);
  assert.strictEqual((html.match(/role="separator"/g) || []).length, 3);
  assert.strictEqual((html.match(/ondblclick="autoFitColumn\(/g) || []).length, 3);
  assert.match(html, /Drag to resize · Double-click to Auto Fit/);
});

test('every Step/Table column has visible resize grips, Auto Fit and persisted widths', () => {
  const fms = clientContext.FMS;
  clientStorage.clear();
  fms.columnPrefs = {};
  fms.columnPrefsLoaded = false;
  clientContext.FMS_SESSION = { userName: 'USER1' };
  fms.viewMode = 'step';
  fms.currentHeader = 'Flow A';
  fms.currentStep = 'Review';
  fms.table.cacheKey = 'Flow A|||Review';
  fms.table.columns = ['Short', 'Long description'];
  fms.table.columnTypes = ['TEXT', 'TEXT'];
  fms.table.columnGroups = [];
  fms.table.lockedCols = [];
  fms.table.hiddenCols = {};
  fms.table.colWidths = {};
  fms.table.columnPanelOpen = false;
  fms.table.filterStatus = 'all';
  fms.table.globalSearch = '';
  fms.table.colSearch = {};
  fms.table.sortCol = null;
  fms.table.page = 1;
  fms.table.pageSize = 50;
  fms.table.steps = [];
  fms.table.rows = [{ row: 8, status: 'pending', overdue: false, cells: ['A', 'A substantially longer value for Auto Fit'] }];

  clientContext.renderDataTable('Flow A', 'Review');
  const stepHtml = tableContainerElement.innerHTML;
  assert.strictEqual((stepHtml.match(/role="separator"/g) || []).length, 4);
  assert.strictEqual((stepHtml.match(/ondblclick="autoFitColumn\(/g) || []).length, 4);
  assert.match(stepHtml, />Action<span class="col-resize-handle col-resize-handle-left"/);

  const originalRerender = clientContext.rerenderCurrentTable;
  clientContext.rerenderCurrentTable = function() {};
  const event = { cancelable: true, preventDefault() {}, stopPropagation() {} };
  clientContext.autoFitAllColumns(event);
  assert.ok(fms.table.colWidths.__serial >= 40);
  assert.ok(fms.table.colWidths.__action >= 72);
  assert.ok(fms.table.colWidths[1] > fms.table.colWidths[0]);
  assert.ok(fms.table.colWidths[1] <= 600);

  fms.viewMode = 'table';
  fms.table.steps = Array.from({ length: 16 }, (_, index) => ({ step: 'Step ' + (index + 1), order: index + 1 }));
  const manyStepsWidth = clientContext.measureAutoFitColumn('__action');
  assert.ok(manyStepsWidth > 320, 'Auto Fit must fit 14+ nowrap step chips');
  assert.ok(manyStepsWidth <= 800);
  fms.table.steps = Array.from({ length: 40 }, (_, index) => ({ step: 'Step ' + (index + 1), order: index + 1 }));
  const fortyStepsWidth = clientContext.measureAutoFitColumn('__action');
  assert.ok(fortyStepsWidth > 800, 'Auto Fit ceiling must grow with every nowrap step chip');
  assert.strictEqual(fortyStepsWidth, clientContext.compactStepsColumnWidth());
  fms.viewMode = 'step';
  fms.table.steps = [];

  fms.table.columns = ['Link'];
  fms.table.rows = [{ row: 8, status: 'pending', overdue: false, cells: ['https://example.com/file'] }];
  fms.currentFontSize = 'medium';
  const mediumLinkWidth = clientContext.measureAutoFitColumn(0);
  fms.currentFontSize = 'large';
  const largeLinkWidth = clientContext.measureAutoFitColumn(0);
  assert.ok(largeLinkWidth > mediumLinkWidth, 'Auto Fit must respect A++ link typography');
  fms.currentFontSize = 'medium';

  clientContext.setColumnWidth(0, 287);
  assert.strictEqual(fms.table.colWidths[0], 287);
  const stored = JSON.parse(clientStorage.get('fms_column_prefs_v2'));
  const scoped = stored['user1|||Flow A|||Review'];
  assert.ok(scoped, 'missing user/view-scoped column preferences');
  assert.strictEqual(scoped.colWidths[0], 287);
  assert.strictEqual(scoped.colWidths.__serial, fms.table.colWidths.__serial);
  assert.strictEqual(scoped.colWidths.__action, fms.table.colWidths.__action);
  clientContext.rerenderCurrentTable = originalRerender;

  const styles = htmlSource['Styles.html'];
  assert.match(styles, /\.col-resize-handle\s*\{[\s\S]*?width:\s*12px/);
  assert.match(styles, /\.col-resize-handle::before/);
  assert.match(styles, /grid-template-columns:\s*repeat\(3/);
  assert.match(htmlSource['Scripts.html'], /Auto Fit All/);
  assert.match(htmlSource['Scripts.html'], /class="column-width-autofit"/);
});

test('cell links and quote-bearing form context are HTML-attribute safe', () => {
  const good = clientContext.renderCellValue('https://example.com/file?a=1&b=2');
  assert.match(good, /^<a href="https:\/\/example\.com\/file\?a=1&amp;b=2"/);
  assert.match(good, /rel="noopener noreferrer"/);

  const quoteBearing = clientContext.renderCellValue('https://example.com/" onclick="alert(1)');
  assert.doesNotMatch(quoteBearing, /<a\s/);
  assert.match(quoteBearing, /&quot;/);
  assert.doesNotMatch(clientContext.renderCellValue('javascript:alert(1)'), /<a\s/);
  assert.strictEqual(clientContext.safeHttpUrl('http://%'), null);
  assert.strictEqual(clientContext.autoFitDisplayText('http://%'), 'http://%');

  const contextHtml = clientContext.renderFormContextHtml([
    { name: 'Order "ID"', value: 'Customer "A"' }
  ], 8);
  assert.match(contextHtml, /Order &quot;ID&quot;/);
  assert.match(contextHtml, /title="Customer &quot;A&quot;"/);
});

test('late submit responses cannot target a newly opened popup generation', () => {
  const fms = clientContext.FMS;
  fms.currentFormHeader = 'Flow A';
  fms.currentFormStep = 'Review';
  fms.currentFormRow = 8;
  fms.currentFormSchemaVersion = 'schema-new';
  fms.currentFormGeneration = 12;
  assert.strictEqual(clientContext.isSubmittedFormStillCurrent('Flow A', 'Review', 8, 'schema-old', 12), false);
  assert.strictEqual(clientContext.isSubmittedFormStillCurrent('Flow A', 'Review', 8, 'schema-new', 11), false);
  assert.strictEqual(clientContext.isSubmittedFormStillCurrent('Flow A', 'Review', 8, 'schema-new', 12), true);

  fms.formFetchLatest = {};
  fms.formFetchInFlight = {};
  fms.formFetchSequence = 0;
  const firstFetch = clientContext.beginFormFetch('Flow A|||Review');
  const newerFetch = clientContext.beginFormFetch('Flow A|||Review');
  assert.strictEqual(clientContext.completeLatestFormFetch('Flow A|||Review', firstFetch), false);
  assert.strictEqual(fms.formFetchInFlight['Flow A|||Review'], true);
  assert.strictEqual(clientContext.completeLatestFormFetch('Flow A|||Review', newerFetch), true);
  assert.strictEqual(fms.formFetchInFlight['Flow A|||Review'], undefined);

  fms.currentFormHeader = null;
  fms.currentFormStep = null;
  fms.currentFormRow = null;
  fms.currentFormSchemaVersion = '';
  fms.currentFormGeneration = 0;
});

test('client optimistic completion updates both views and can roll back', () => {
  const fms = clientContext.FMS;
  const stepRow = { row: 8, status: 'pending', overdue: true };
  const multiRow = {
    row: 8,
    status: 'pending',
    overdue: true,
    steps: [
      { step: 'Review', order: 1, status: 'pending', overdue: true },
      { step: 'Approve', order: 2, status: 'locked', overdue: false }
    ]
  };
  fms.tableCache = { 'Flow A|||Review': { rows: [stepRow] } };
  fms.multiTableCache = { 'Flow A': { rows: [multiRow] } };

  const optimistic = clientContext.applyOptimisticStepCompletion('Flow A', 'Review', 8);
  assert.strictEqual(stepRow.status, 'completed');
  assert.strictEqual(stepRow.overdue, false);
  assert.strictEqual(multiRow.steps[0].status, 'completed');
  assert.strictEqual(multiRow.steps[1].status, 'pending');
  optimistic.revert();
  assert.strictEqual(stepRow.status, 'pending');
  assert.strictEqual(stepRow.overdue, true);
  assert.strictEqual(multiRow.steps[0].status, 'pending');
  assert.strictEqual(multiRow.steps[1].status, 'locked');
});

test('form cache derives order context and planned date directly from the loaded table row', () => {
  const fms = clientContext.FMS;
  fms.tableCache = {
    'Flow A|||Step 1': {
      columns: ['Order ID', 'Customer', 'Planned', 'Actual', 'Note'],
      baseColCount: 2,
      rows: [{ row: 18, cells: ['ORD-1042', 'Sample Customer', '28 Jul 26 15:00:00', '', ''], status: 'pending' }]
    }
  };
  fms.multiTableCache = {};
  const snapshot = clientContext.getCachedFormSnapshot('Flow A', 'Step 1', 18);
  assert.deepStrictEqual(
    Array.from(snapshot.context, (detail) => ({ name: detail.name, value: detail.value })),
    [
      { name: 'Order ID', value: 'ORD-1042' },
      { name: 'Customer', value: 'Sample Customer' }
    ]
  );
  assert.strictEqual(snapshot.planned, '28 Jul 26 15:00:00');
  assert.strictEqual(
    clientContext.formIdentitySubtitle('Flow A', snapshot.context, 18),
    'Order ID: ORD-1042  •  Flow A  •  Record #18'
  );

  clientContext.cacheFormSchemasFromTable('Flow A', 'Step 1', {
    formSchema: [
      { name: 'Planned', type: 'PLANNED_DISPLAY', value: '', options: [] },
      { name: 'Payment Proof Image', type: 'FILE', options: [] }
    ],
    formSchemaVersion: 'schema-test-step-1'
  });
  assert.strictEqual(clientContext.getFormCacheEntry('Flow A|||Step 1').fields.length, 2);
  assert.strictEqual(clientContext.getFormCacheEntry('Flow A|||Step 1').version, 'schema-test-step-1');
  assert.strictEqual(clientContext.isFormCacheFresh(clientContext.getFormCacheEntry('Flow A|||Step 1')), true);

  const blankRowFields = clientContext.fieldsWithSnapshot([
    { name: 'Planned', type: 'PLANNED_DISPLAY', value: 'OLD ROW DATE', options: [] }
  ], { context: [], planned: '' });
  assert.strictEqual(blankRowFields[0].value, '', 'empty Planned must not reuse another row value');
  clientContext.storeFormStructureEntry('Flow A|||Step 1', [
    { name: 'Planned', type: 'PLANNED_DISPLAY', value: 'ROW-SPECIFIC DATE', options: [] }
  ], 'schema-test-step-1', Date.now());
  assert.strictEqual(clientContext.getFormCacheEntry('Flow A|||Step 1').fields[0].value, '');

  fms.currentFormHeader = 'Flow A';
  fms.currentFormStep = 'Step 1';
  fms.formSchemaOutdated = false;
  const schemaChanged = clientContext.cacheFormSchemasFromTable('Flow A', 'Step 1', {
    formSchema: [
      { name: 'Planned', type: 'PLANNED_DISPLAY', value: '', options: [] },
      { name: 'New proof', type: 'FILE', options: [] }
    ],
    formSchemaVersion: 'schema-test-step-2'
  });
  assert.strictEqual(schemaChanged, true);
  assert.strictEqual(fms.formSchemaOutdated, true, 'schema-only table refresh must invalidate an open form');
  fms.currentFormHeader = null;
  fms.currentFormStep = null;
  fms.formSchemaOutdated = false;
});

test('modal stays clear behind the table and uses the polished context layout', () => {
  const styles = htmlSource['Styles.html'];
  const modalCss = styles.slice(styles.indexOf('/* ===== MODAL / POPUP FORM ===== */'), styles.indexOf('/* ===== DATA TABLE'));
  const scripts = htmlSource['Scripts.html'];
  assert.doesNotMatch(modalCss, /backdrop-filter:\s*blur/);
  assert.match(modalCss, /backdrop-filter:\s*none\s*!important/);
  assert.match(modalCss, /\.form-context-grid/);
  assert.match(modalCss, /\.form-entry-card/);
  assert.match(scripts, /Order details/);
  assert.match(scripts, /function formIdentitySubtitle/);
  assert.match(scripts, /updateModalIdentity\(headerName, stepName, rowNumber, snapshot\.context\)/);
  assert.match(scripts, /FORM_SCHEMA_CACHE_MS = 30 \* 60 \* 1000/);
  assert.match(scripts, /cacheFormSchemasFromTable/);
  assert.match(scripts, /<div class="modal-footer"><div class="form-msg"/);
  assert.doesNotMatch(scripts, /Loading form\.\.\./);
});

test('client has guarded login RPC and compact numeric step chips', () => {
  const login = htmlSource['Login.html'];
  const scripts = htmlSource['Scripts.html'];
  const styles = htmlSource['Styles.html'];
  assert.match(login, /setTimeout\(function\(\)[\s\S]*45000/);
  assert.match(login, /Deployment mismatch/);
  assert.match(scripts, /var label = String\(st\.order\)/);
  assert.doesNotMatch(scripts, /var label = 'Step ' \+ st\.order/);
  assert.match(scripts, /function compactStepsColumnWidth\(\)/);
  assert.match(scripts, /12 \+ \(count \* 20\)/);
  assert.match(styles, /min-width: 20px; width: 20px; height: 20px/);
  assert.match(styles, /flex-wrap: nowrap/);
  assert.match(styles, /chip-done[\s\S]*var\(--success\)/);
  assert.match(styles, /chip-overdue[\s\S]*#ff5252/);
});

console.log('\n' + passed + ' FMS runtime tests passed.');
