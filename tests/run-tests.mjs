/* node tests/run-tests.mjs   (ends READY or NOT READY, like the minibus and
   ushers repos). Covers what this round added: vehicle.js, pdf.js, the sheet's
   vehicle check handler, and that index.html still parses and has every
   element its script asks for. */
import fs from "node:fs";
import vm from "node:vm";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { FakeSheet, makeGas } from "./gas.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const read = f => fs.readFileSync(path.join(ROOT, f), "utf8");
let pass = 0, fail = 0;
function t(name, fn) {
  try { fn(); pass++; }
  catch (e) { fail++; console.log("✗ " + name + "\n    " + (e && e.message || e)); }
}
function eq(a, b, msg) { if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error((msg || "") + " expected " + JSON.stringify(b) + ", got " + JSON.stringify(a)); }
function ok(v, msg) { if (!v) throw new Error(msg || "expected truthy"); }

/* ---------- vehicle.js ---------- */
const vctx = {}; vm.createContext(vctx); vm.runInContext(read("vehicle.js"), vctx);
const V = vctx.TaxiVehicle;

const allFine = () => Object.fromEntries(V.ALL_ITEMS.map(i => [i.id, { mark: "fine", note: "" }]));

t("every item has a unique id, a name and words", () => {
  const ids = new Set();
  V.ALL_ITEMS.forEach(i => { ok(i.id && i.name && i.what, i.id); ok(!ids.has(i.id), "dup " + i.id); ids.add(i.id); });
  ok(V.ALL_ITEMS.length >= 20);
});
t("all fine is Fit", () => { const o = V.checkOutcome(allFine()); eq([o.complete, o.outcome, o.fit], [true, "Fit", true]); });
t("an unmarked item is incomplete", () => {
  const r = allFine(); delete r.tyres; const o = V.checkOutcome(r);
  eq([o.complete, o.missing], [false, ["tyres"]]);
});
t("advisory keeps it fit", () => {
  const r = allFine(); r.wash = { mark: "advisory", note: "low" };
  eq(V.checkOutcome(r).outcome, "Fit with advisories");
});
t("a defect needs a note", () => {
  const r = allFine(); r.seats = { mark: "defect", note: " " };
  const o = V.checkOutcome(r); eq([o.complete, o.needNote], [false, ["seats"]]);
});
t("non-critical defect is Defects found, still fit", () => {
  const r = allFine(); r.seats = { mark: "defect", note: "tip-up seat jams" };
  const o = V.checkOutcome(r); eq([o.complete, o.outcome, o.fit], [true, "Defects found", true]);
  eq(V.describeFindings(o.defects), "Seats: tip-up seat jams");
});
t("critical defect is Unfit", () => {
  const r = allFine(); r.brakes = { mark: "defect", note: "pedal sinks" };
  const o = V.checkOutcome(r); eq([o.outcome, o.fit], ["Unfit", false]);
  eq(V.describeFindings(o.defects), "Brakes (critical): pedal sinks");
});

const day = (y, m, d) => new Date(y, m - 1, d, 9, 30);
t("renewal levels at 61, 60, 30, 7, 1, 0, -1 days", () => {
  const today = day(2026, 11, 19);
  const at = n => { const d = new Date(2026, 10, 19 + n); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); };
  const lv = n => V.renewalStatus({ mot: at(n) }, today)[0].level;
  eq([61, 60, 30, 8, 7, 1, 0, -1].map(lv), ["ok", "upcoming", "due", "due", "soon", "soon", "today", "expired"]);
  eq(V.renewalStatus({ mot: at(1) }, today)[0].text, "MOT / compliance test expires tomorrow");
  eq(V.renewalStatus({ mot: at(-1) }, today)[0].text, "MOT / compliance test expired on 18/11/2026");
});
t("renewals across the clock change count whole days", () => {
  eq(V.renewalStatus({ insurance: "2027-03-29" }, day(2027, 3, 27))[0].days, 2);
});
t("only gated documents block, and only once expired", () => {
  const today = day(2026, 11, 19);
  const b = V.blockingRenewals({ insurance: "2026-11-18", dvla: "2026-01-01", mot: "2026-11-19" }, today);
  eq(b.map(x => x.id), ["insurance"]);
});
t("bad dates are ignored", () => {
  eq(V.renewalStatus({ mot: "2026-02-31", insurance: "nonsense" }, day(2026, 1, 1)), []);
});
t("renewal summary for the sheet", () => {
  eq(V.renewalSummary({ insurance: "2027-01-01", mot: "2027-03-14" }), "Insurance 01/01/2027; MOT / compliance test 14/03/2027");
});

/* ---------- pdf.js ---------- */
const require = createRequire(import.meta.url);
const pctx = { jspdf: require(path.join(ROOT, "vendor/jspdf.umd.min.js")) };
vm.createContext(pctx); vm.runInContext(read("pdf.js"), pctx);
const P = pctx.TaxiPdf;
const parse = s => { const m = /^(\d{2})\/(\d{2})\/(\d{4}), (\d{2}):(\d{2}):(\d{2})$/.exec(s); return m ? new Date(+m[3], m[2] - 1, +m[1], +m[4], +m[5], +m[6]) : null; };
const uk = d => d.toLocaleDateString("en-GB");

t("tax year starts 6 April", () => {
  let r = P.periodRange("tax-year", day(2026, 4, 5)); eq([uk(r.from), r.label], ["06/04/2025", "Tax year 2025-26 to date"]);
  r = P.periodRange("tax-year", day(2026, 4, 6)); eq(uk(r.from), "06/04/2026");
});
t("weeks start Monday; last month across a year", () => {
  let r = P.periodRange("this-week", day(2026, 11, 22)); eq(uk(r.from), "16/11/2026");    /* a Sunday */
  r = P.periodRange("last-week", day(2026, 11, 16)); eq([uk(r.from), uk(r.to)], ["09/11/2026", "16/11/2026"]);
  r = P.periodRange("last-month", day(2027, 1, 10)); eq([uk(r.from), r.text], ["01/12/2026", "01/12/2026 to 31/12/2026"]);
});
const trips = [
  { tripId: "a", date: "19/11/2026, 08:10:00", fare: 9.70, meteredFare: 7.40, extras: [{ label: "Tunnel toll", amount: 2.30 }], miles: "2.00", time: "00:12:00", tariff: "Day" },
  { tripId: "b", date: "19/11/2026, 23:40:00", fare: 30, meteredFare: 26.10, agreedFare: 30, extras: [], miles: "11.20", time: "00:25:00", tariff: "Night" },
  { tripId: "c", date: "31/10/2026, 12:00:00", fare: 5, meteredFare: 5, miles: "1.10", time: "00:06:00", tariff: "Day" },
  { tripId: "d", date: "garbage", fare: 100 }
];
t("statement adds up only its period", () => {
  const s = P.buildStatement(trips, P.periodRange("this-month", day(2026, 11, 20)), parse);
  eq([s.count, s.total, s.metered, s.extras, s.agreedCount, s.miles, s.average], [2, 39.7, 33.5, 2.3, 1, 13.2, 19.85]);
  eq(s.days, [{ date: "19/11/2026", trips: 2, fare: 39.7, miles: 13.2 }]);
  eq(s.tariffs.map(x => x.band), ["Day", "Night"]);
});
t("receipt PDF is made and carries the total", () => {
  const doc = P.receiptPdf(trips[0], { driver: "Test Driver", driverId: "TX-AAAAA", vehicle: "AB12 CDE", made: "19/11/2026, 09:00:00", compress: false });
  const out = doc.output();
  ok(out.startsWith("%PDF"), "not a PDF");
  ok(out.includes("Tunnel toll") && out.includes("9.70"), "text missing");
});
t("statement PDF splits a long period across pages", () => {
  const many = Array.from({ length: 160 }, (_, i) => ({ tripId: "x" + i, date: "0" + (1 + i % 9) + "/11/2026, 10:00:00", fare: 8, meteredFare: 8, miles: "2.00", time: "00:10:00", tariff: "Day" }));
  const s = P.buildStatement(many, P.periodRange("this-month", day(2026, 11, 20)), parse);
  const doc = P.statementPdf(s, { driver: "Test Driver", made: "now", compress: false });
  ok(doc.getNumberOfPages() >= 3, "pages " + doc.getNumberOfPages());
  ok(doc.output().includes("Page 1 of " + doc.getNumberOfPages()));
});
t("file names", () => {
  eq(P.fileName("receipt", trips[0]), "Taxi receipt 2026-11-19.pdf");
  eq(P.fileName("statement", P.periodRange("last-month", day(2026, 11, 20))), "Earnings statement 2026-10-01 to 2026-10-31.pdf");
});

/* ---------- Code.gs: the vehicle check ---------- */
function loadGas(sheets) {
  const { g, book } = makeGas(sheets, { APP_SECRET: "test-secret" });
  vm.createContext(g); vm.runInContext(read("Code.gs"), g);
  return { g, book };
}
const drivers = () => new FakeSheet("Drivers", [["Driver ID", "Name", "Active (TRUE/FALSE)"], ["TX-AAAAA", "Ann Driver", true], ["TX-BBBBB", "Off Driver", false]]);
const post = (g, body) => JSON.parse(g.doPost({ postData: { contents: JSON.stringify(body) } }).text);
const goodCheck = { checkId: "c-1", date: "19/11/2026, 07:55:10", vehicle: "AB12 CDE", odometer: "84210", outcome: "Unfit", defects: "Brakes (critical): pedal sinks", advisories: "", renewals: "Insurance 01/01/2027" };

t("a check is filed by header, once", () => {
  const { g, book } = loadGas({ Drivers: drivers() });
  const token = g.makeToken_("TX-AAAAA");
  eq(post(g, { type: "vehicleCheck", token, check: goodCheck }).status, "ok");
  const r2 = post(g, { type: "vehicleCheck", token, check: goodCheck });
  eq([r2.status, r2.duplicate], ["ok", true]);
  const sh = book.sheets["Vehicle Checks"];
  eq(sh.rows[0], g.CHECK_HEADERS.slice());
  eq(sh.getLastRow(), 2);
  const row = Object.fromEntries(sh.rows[0].map((h, i) => [h, sh.rows[1][i]]));
  eq([row["Driver Name"], row["Outcome"], row["Odometer"], row["Check ID"]], ["Ann Driver", "Unfit", 84210, "c-1"]);
  ok(row["Check Date"] instanceof Date || Object.prototype.toString.call(row["Check Date"]) === "[object Date]", "date not a Date");
});
t("the office's own column order and extra columns are respected", () => {
  const sh = new FakeSheet("Vehicle Checks", [["Check ID", "Office note", "Outcome"]]);
  const { g } = loadGas({ Drivers: drivers(), "Vehicle Checks": sh });
  eq(post(g, { type: "vehicleCheck", token: g.makeToken_("TX-AAAAA"), check: goodCheck }).status, "ok");
  eq(sh.rows[0].slice(0, 4), ["Check ID", "Office note", "Outcome", "Received"]);
  eq([sh.rows[1][0], sh.rows[1][1], sh.rows[1][2]], ["c-1", "", "Unfit"]);
});
t("refusals: no token, suspended driver, bad outcome, bad date, formula injection", () => {
  const { g, book } = loadGas({ Drivers: drivers() });
  eq(post(g, { type: "vehicleCheck", check: goodCheck }).code, "badtoken");
  eq(post(g, { type: "vehicleCheck", token: g.makeToken_("TX-BBBBB"), check: goodCheck }).code, "badtoken");
  const tok = g.makeToken_("TX-AAAAA");
  eq(post(g, { type: "vehicleCheck", token: tok, check: { ...goodCheck, outcome: "Fine-ish" } }).code, "invalid");
  eq(post(g, { type: "vehicleCheck", token: tok, check: { ...goodCheck, date: "31/02/2026, 10:00:00" } }).code, "invalid");
  eq(post(g, { type: "vehicleCheck", token: tok, check: { ...goodCheck, odometer: "-5" } }).code, "invalid");
  eq(post(g, { type: "vehicleCheck", token: tok, check: { ...goodCheck, checkId: "c-2", vehicle: "=HYPERLINK(1)" } }).status, "ok");
  const sh = book.sheets["Vehicle Checks"];
  ok(String(sh.rows[1][sh.rows[0].indexOf("Vehicle")]).startsWith("'="), "not escaped");
});
t("validate and doGet carry the sheet version", () => {
  const { g } = loadGas({ Drivers: drivers() });
  eq(post(g, { type: "validate", driverId: "TX-AAAAA" }).version, g.SHEET_VERSION);
  eq(JSON.parse(g.doGet().text).version, g.SHEET_VERSION);
});

/* ---------- index.html ---------- */
const html = read("index.html");
const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
t("the page script parses", () => { scripts.forEach(s => new vm.Script(s)); });
t("every $(\"id\") the script asks for is in the markup", () => {
  const ids = new Set([...html.matchAll(/\sid=["']([^"']+)["']/g)].map(m => m[1]));
  const dynamic = new Set(["sheetTitle"]);
  const want = new Set([...scripts.join("\n").matchAll(/\$\("([A-Za-z][\w-]*)"\)/g)].map(m => m[1]));
  const missing = [...want].filter(id => !ids.has(id) && !dynamic.has(id) && !id.startsWith("rn-") && !id.startsWith("page-"));
  eq(missing, []);
});
t("APP_VERSION matches the service worker's note", () => {
  const v = /const APP_VERSION = "([^"]+)"/.exec(html)[1];
  ok(read("sw.js").includes("app " + v), "sw.js CACHE_VERSION comment does not name app " + v);
});
t("the service worker precaches the new files", () => {
  const sw = read("sw.js");
  ["./vehicle.js", "./pdf.js", "./vendor/jspdf.umd.min.js"].forEach(f => ok(sw.includes('"' + f + '"'), f));
});

console.log(pass + " passed, " + fail + " failed");
console.log(fail ? "NOT READY" : "READY");
process.exit(fail ? 1 : 0);
