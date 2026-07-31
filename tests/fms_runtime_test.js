'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.resolve(__dirname, '..');
const read = (name) => fs.readFileSync(path.join(ROOT, name), 'utf8');
const serverFiles = ['Code.gs', 'TAT_Calculator.gs', 'WorkflowEngine.gs'];
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
const clientWindow = {
  localStorage: { getItem() { return null; }, setItem() {} },
  matchMedia() { return { matches: false }; },
  innerWidth: 1440,
  innerHeight: 900,
  addEventListener() {}
};
const clientContext = vm.createContext({
  console, Date, JSON, Math, String, Number, Boolean, Object, Array, RegExp, Error,
  parseInt, parseFloat, isNaN, isFinite, setTimeout, clearTimeout, setInterval, clearInterval,
  document: clientDocument, window: clientWindow
});
const scriptsBody = htmlSource['Scripts.html'].match(/^<script>([\s\S]*)<\/script>\s*$/i)[1];
new vm.Script(scriptsBody, { filename: 'Scripts.html' }).runInContext(clientContext);

test('all Apps Script server files parse and load together', () => {
  serverFiles.forEach((name) => assert.ok(read(name).length > 0));
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
  serverFiles.forEach((name) => {
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
  assert.strictEqual(info.buildId, '2026-07-22-runtime-safe-1');
});

test('doGet renders FMS Login and explicit deployment status', () => {
  const login = context.doGet({ parameter: {} });
  assert.strictEqual(login.title, 'FMS - Flow Management System');
  assert.match(login.getContent(), /Flow Management System/);
  assert.match(login.getContent(), /Build 2026-07-22-runtime-safe-1/);
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

  const multi = context.wfGetMultiStepTableData(masterUrl, 'MASTER', 'STEPS', 'Flow A', 'USER1');
  assert.strictEqual(multi.success, true);
  assert.strictEqual(multi.steps.length, 2);
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
  const form = context.wfGetStepForm(masterUrl, 'MASTER', 'DROPDPWN', 'Flow A', 'Step 1', 8);
  assert.strictEqual(form.success, true);
  assert.ok(Array.from(form.fields).some((field) => field.type === 'PLANNED_DISPLAY'));
  context.formatResult = original;
});

test('edit/submit writes fields, Actual and next Planned, then invalidates Home cache', () => {
  const homeKey = context.fmsCacheKey(['home', 'USER1']);
  cache.set(homeKey, JSON.stringify({ success: true, stale: true }));
  const result = context.wfSubmitStep(
    masterUrl, 'MASTER', 'STEPS', 'Flow A', 'Step 1', 8,
    { Note: 'Edited from form' }, {}, 'USER1'
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
