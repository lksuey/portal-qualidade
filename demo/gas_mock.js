'use strict';
/**
 * Minimal in-memory mock of the Google Apps Script services used by Code.gs.
 * It lets the real server code run under plain Node (no dependencies) for the
 * local demo and for the tests. It is NOT a faithful emulation of Google: it
 * implements only what this project calls.
 */
process.env.TZ = process.env.TZ || 'America/Sao_Paulo';

const vm = require('vm');
const fs = require('fs');
const crypto = require('crypto');

const PAD = (n, w = 2) => String(n).padStart(w, '0');

/* ---------- Utilities ---------- */
function toBuffer(data) {
  if (Buffer.isBuffer(data)) return data;
  if (Array.isArray(data)) return Buffer.from(data.map((b) => b & 0xff));
  return Buffer.from(String(data), 'utf8');
}

function formatDate(date, tz, pattern) {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: tz, hourCycle: 'h23', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
  }).formatToParts(date).reduce((acc, p) => { acc[p.type] = p.value; return acc; }, {});
  return pattern
    .replace(/yyyy/g, parts.year).replace(/MM/g, parts.month).replace(/dd/g, parts.day)
    .replace(/HH/g, parts.hour).replace(/mm/g, parts.minute).replace(/ss/g, parts.second);
}

function makeBlob(data, contentType, name) {
  const buf = toBuffer(data);
  return {
    getBytes: () => Array.from(buf).map((b) => (b > 127 ? b - 256 : b)),
    getDataAsString: () => buf.toString('utf8'),
    getContentType: () => contentType || 'application/octet-stream',
    getName: () => name || '',
  };
}

const Utilities = {
  DigestAlgorithm: { SHA_256: 'SHA_256' },
  computeDigest: (alg, value) => Array.from(crypto.createHash('sha256').update(toBuffer(value)).digest())
    .map((b) => (b > 127 ? b - 256 : b)),
  getUuid: () => crypto.randomUUID(),
  base64Encode: (data) => toBuffer(data).toString('base64'),
  base64EncodeWebSafe: (data) => toBuffer(data).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
  base64DecodeWebSafe: (s) => Array.from(Buffer.from(String(s).replace(/-/g, '+').replace(/_/g, '/'), 'base64'))
    .map((b) => (b > 127 ? b - 256 : b)),
  newBlob: makeBlob,
  formatDate,
};

/* ---------- Spreadsheet ---------- */
class Range {
  constructor(sheet, row, col, numRows, numCols) {
    Object.assign(this, { sheet, row, col, numRows, numCols });
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const line = [];
      for (let c = 0; c < this.numCols; c++) {
        const v = (this.sheet.data[this.row - 1 + r] || [])[this.col - 1 + c];
        line.push(v === undefined ? '' : v);
      }
      out.push(line);
    }
    return out;
  }
  getValue() { return this.getValues()[0][0]; }
  setValues(values) {
    values.forEach((line, r) => line.forEach((v, c) => this.sheet._set(this.row + r, this.col + c, v)));
    return this;
  }
  setValue(v) { this.sheet._set(this.row, this.col, v); return this; }
  setBackground() { return this; }
  setFontColor() { return this; }
  setFontWeight() { return this; }
  setNumberFormat() { return this; }
}

class Sheet {
  constructor(name) { this.name = name; this.data = []; }
  _set(row, col, v) {
    while (this.data.length < row) this.data.push([]);
    const line = this.data[row - 1];
    while (line.length < col) line.push('');
    line[col - 1] = v;
  }
  getName() { return this.name; }
  setName(n) { this.name = n; }
  getLastRow() { return this.data.length; }
  getLastColumn() { return this.data.reduce((m, l) => Math.max(m, l.length), 0); }
  getRange(row, col, numRows = 1, numCols = 1) { return new Range(this, row, col, numRows, numCols); }
  getDataRange() { return new Range(this, 1, 1, this.getLastRow(), this.getLastColumn()); }
  appendRow(values) { this.data.push(values.slice()); }
  clearContents() { this.data = []; }
  clear() { this.data = []; }
  setFrozenRows() {}
  hideColumns() {}
}

class Workbook {
  constructor(name) { this.name = name; this.id = 'demo-' + crypto.randomUUID(); this.sheets = [new Sheet('Sheet1')]; }
  getId() { return this.id; }
  getUrl() { return 'https://example.invalid/d/' + this.id + '/edit'; }
  getSheets() { return this.sheets; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  getActiveSheet() { return this.sheets[0]; }
  insertSheet(n) { const s = new Sheet(n); this.sheets.push(s); return s; }
}

/* ---------- Cache (enforces the 100KB per value limit of the real service) ---------- */
function makeCache(clock) {
  const store = new Map();
  const live = (k) => {
    const e = store.get(k);
    if (!e) return null;
    if (e.expires <= clock()) { store.delete(k); return null; }
    return e.value;
  };
  const put = (k, v, ttl = 600) => {
    if (String(v).length > 100 * 1024) throw new Error('Argument too large: value');
    store.set(k, { value: String(v), expires: clock() + ttl * 1000 });
  };
  return {
    get: (k) => live(k),
    put,
    putAll: (obj, ttl) => Object.keys(obj).forEach((k) => put(k, obj[k], ttl)),
    getAll: (keys) => keys.reduce((acc, k) => { const v = live(k); if (v !== null) acc[k] = v; return acc; }, {}),
    remove: (k) => { store.delete(k); },
    _size: () => store.size,
  };
}

/**
 * Builds a vm context with the Apps Script globals and loads the given source files.
 * options: { files: [paths], now: Date|number (frozen clock) , properties: {}, htmlFiles: {name: path} }
 */
function createRuntime(options) {
  let frozen = options.now !== undefined ? new Date(options.now).getTime() : null;
  const RealDate = Date;
  const clock = () => (frozen !== null ? frozen : RealDate.now());
  const setNow = (value) => { frozen = new Date(value).getTime(); };

  class MockDate extends RealDate {
    constructor(...args) { if (args.length === 0) super(clock()); else super(...args); }
    static now() { return clock(); }
  }

  const properties = Object.assign({}, options.properties || {});
  const workbooks = new Map();
  const triggers = [];
  const logs = [];
  const cache = makeCache(clock);

  const SpreadsheetApp = {
    create: (name) => { const w = new Workbook(name); workbooks.set(w.id, w); return w; },
    openById: (id) => { if (!workbooks.has(id)) throw new Error('No item with the given ID could be found: ' + id); return workbooks.get(id); },
    flush: () => {},
  };
  const PropertiesService = {
    getScriptProperties: () => ({
      getProperty: (k) => (Object.prototype.hasOwnProperty.call(properties, k) ? properties[k] : null),
      setProperty: (k, v) => { properties[k] = String(v); },
    }),
  };
  const HtmlService = {
    XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
    _out(file) {
      const out = {
        title: '', file, getContent: () => fs.readFileSync(options.htmlFiles[file], 'utf8'),
        setTitle(t) { out.title = t; return out; },
        addMetaTag() { return out; },
        setXFrameOptionsMode() { return out; },
        evaluate() { return out; },
      };
      return out;
    },
    createTemplateFromFile(file) { return HtmlService._out(file); },
    createHtmlOutputFromFile(file) { return HtmlService._out(file); },
  };
  const ScriptApp = {
    getProjectTriggers: () => triggers.slice(),
    deleteTrigger: (t) => { const i = triggers.indexOf(t); if (i >= 0) triggers.splice(i, 1); },
    newTrigger: (fn) => ({
      timeBased: () => ({ everyMinutes: () => ({ create: () => { const t = { getHandlerFunction: () => fn }; triggers.push(t); return t; } }) }),
    }),
    getOAuthToken: () => 'mock-token',
  };
  const lock = { waitLock() {}, releaseLock() {} };

  const sandbox = {
    console, Date: MockDate, Utilities, SpreadsheetApp, PropertiesService, HtmlService, ScriptApp,
    CacheService: { getScriptCache: () => cache },
    LockService: { getScriptLock: () => lock },
    Session: { getScriptTimeZone: () => 'America/Sao_Paulo' },
    Logger: { log: (m) => logs.push(String(m)) },
    DriveApp: { getFileById: () => ({ setTrashed() {} }) },
    // The real Google export endpoint is not available offline: the demo returns the
    // sheet rows as CSV so that "export" still produces a usable file.
    UrlFetchApp: {
      fetch: (url) => {
        const m = String(url).match(/\/d\/(demo-[^/]+)\//);
        const wb = m && workbooks.get(m[1]);
        const rows = wb ? wb.sheets[0].data : [];
        const csv = rows.map((l) => l.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(';')).join('\n');
        return { getBlob: () => makeBlob(csv, 'text/csv') };
      },
    },
  };
  const context = vm.createContext(sandbox);
  (options.files || []).forEach((f) => vm.runInContext(fs.readFileSync(f, 'utf8'), context, { filename: f }));
  // `const` declarations in the sources are not properties of the context; expose
  // them (and `let`) through an evaluator so tests and the server can read them.
  const evaluate = (code) => vm.runInContext(code, context);
  return { context, evaluate, workbooks, properties, logs, cache, triggers, clock, setNow };
}

module.exports = { createRuntime, Workbook, Sheet, formatDate, PAD };
