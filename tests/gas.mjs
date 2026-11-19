/* A small stand-in for the Apps Script services Code.gs uses, enough to run
   its web handlers in Node. In the spirit of the minibus repo's tests/lib/gas.mjs. */
import crypto from "node:crypto";

class FakeRange {
  constructor(sheet, r, c, nr = 1, nc = 1) { Object.assign(this, { sheet, r, c, nr, nc }); }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(this.sheet.cell(this.r + i, this.c + j));
      out.push(row);
    }
    return out;
  }
  setValue(v) { this.sheet.set(this.r, this.c, v); return this; }
  setValues(vs) { vs.forEach((row, i) => row.forEach((v, j) => this.sheet.set(this.r + i, this.c + j, v))); return this; }
  setFontWeight() { return this; }
  setNumberFormat() { return this; }
  setFontSize() { return this; }
  setFontColor() { return this; }
  setFormula(f) { this.sheet.set(this.r, this.c, f); return this; }
  setFormulas(fs) { return this.setValues(fs); }
  getValue() { return this.sheet.cell(this.r, this.c); }
  getSheet() { return this.sheet; }
  protect() { return this.sheet.book._protect("RANGE", this); }
  setNumberFormats() { return this; }
}

export class FakeSheet {
  constructor(name, rows = []) { this.name = name; this.rows = rows.map(r => r.slice()); this.maxCols = 26; this.frozen = 0; this.writes = []; }
  getName() { return this.name; }
  cell(r, c) { const row = this.rows[r - 1]; return row && row[c - 1] !== undefined ? row[c - 1] : ""; }
  set(r, c, v) {
    this.writes.push(r + ":" + c);
    while (this.rows.length < r) this.rows.push([]);
    const row = this.rows[r - 1];
    while (row.length < c) row.push("");
    row[c - 1] = v;
  }
  getLastRow() { for (let i = this.rows.length; i > 0; i--) if (this.rows[i - 1].some(v => v !== "")) return i; return 0; }
  getLastColumn() { let m = 0; this.rows.forEach(r => r.forEach((v, j) => { if (v !== "") m = Math.max(m, j + 1); })); return m; }
  getMaxColumns() { return this.maxCols; }
  getMaxRows() { return 1000; }
  insertColumnsAfter(after, n) { this.maxCols += n; }
  insertRowsAfter() {}
  getRange(r, c, nr, nc) {
    if (typeof r === "string") {          /* "A1" or "B5:B8" */
      const cell = a => { const m = /^([A-Z]+)(\d+)$/.exec(a); let n = 0; for (const ch of m[1]) n = n * 26 + ch.charCodeAt(0) - 64; return [Number(m[2]), n]; };
      const [a, b] = r.split(":"); const [r1, c1] = cell(a); const [r2, c2] = b ? cell(b) : [r1, c1];
      return new FakeRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    return new FakeRange(this, r, c, nr, nc);
  }
  getDataRange() { return new FakeRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  appendRow(row) { this.rows.splice(this.getLastRow(), 0, row.slice()); }
  getFrozenRows() { return this.frozen; }
  clear() { this.rows = []; }
  setColumnWidth() {}
  protect() { return this.book._protect("SHEET", { getSheet: () => this }); }
  setFrozenRows(n) { this.frozen = n; }
}

export function makeGas(sheets, props = {}) {
  const protections = [];
  const book = {
    sheets, protections,
    _protect(type, range) {
      const p = { type, range, desc: "", warn: false,
        setDescription(d) { this.desc = d; return this; }, getDescription() { return this.desc; },
        setWarningOnly(w) { this.warn = w; return this; }, getRange() { return this.range; },
        remove() { protections.splice(protections.indexOf(this), 1); } };
      protections.push(p); return p;
    },
    getProtections: type => protections.filter(p => p.type === type),
    getSheetByName: n => { const sh = sheets[n] || null; if (sh) sh.book = book; return sh; },
    insertSheet: n => { sheets[n] = new FakeSheet(n); sheets[n].book = book; return sheets[n]; },
    getSpreadsheetTimeZone: () => "Europe/London",
    toast: () => {}
  };
  const cache = new Map();
  const toasts = [];
  const pad = n => String(n).padStart(2, "0");
  const g = {
    SpreadsheetApp: { getActiveSpreadsheet: () => book, getUi: () => { throw new Error("no ui"); },
                      ProtectionType: { SHEET: "SHEET", RANGE: "RANGE" } },
    ScriptApp: { getProjectTriggers: () => g._triggers, newTrigger: h => ({ timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create: () => g._triggers.push({ getHandlerFunction: () => h }) }) }) }) }) },
    _triggers: [],
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    CacheService: { getScriptCache: () => ({ get: k => cache.get(k) || null, put: (k, v) => cache.set(k, v) }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Session: { getScriptTimeZone: () => "Europe/London" },
    Logger: { log: () => {} },
    ContentService: {
      MimeType: { JSON: "json" },
      createTextOutput: text => ({ text, setMimeType() { return this; } })
    },
    Utilities: {
      formatDate: d => pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear() + ", " +
                       pad(d.getHours()) + ":" + pad(d.getMinutes()) + ":" + pad(d.getSeconds()),
      computeHmacSha256Signature: (v, k) => [...crypto.createHmac("sha256", k).update(v).digest()],
      base64EncodeWebSafe: v => Buffer.from(typeof v === "string" ? v : Uint8Array.from(v.map(b => b & 255))).toString("base64url"),
      base64DecodeWebSafe: v => [...Buffer.from(v, "base64url")],
      newBlob: bytes => ({ getDataAsString: () => Buffer.from(bytes).toString("utf8") }),
      getUuid: () => crypto.randomUUID()
    },
    _toasts: toasts
  };
  return { g, book };
}
