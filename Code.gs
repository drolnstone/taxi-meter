/*  ============================================================
    Taxi Meter backend. Complete file. Replace everything in the
    Apps Script editor with this.

    After pasting, reload the spreadsheet. A "Taxi Meter" menu
    appears next to Help with every maintenance job on it.

    FIRST RUN, in this order:
      1. Taxi Meter → Generate app secret
      2. Taxi Meter → Set OpenRouteService key   (a NEW key, see below)
      3. Taxi Meter → Repair trip dates          (once, fixes old rows)
      4. Run any function from the editor and accept the prompts.
         This version calls UrlFetchApp, which is a new permission.
         Without re-authorising, routing fails on the live URL.
      5. Deploy → Manage deployments → edit → New version.

    The old OpenRouteService key was readable in the deployed page.
    Delete it at openrouteservice.org and issue a fresh one. Storing
    the same key here does not un-publish it.

    Trips sheet layout, columns A to Q:
      A Sync Timestamp   B Driver ID       C Driver Name     D Trip Date
      E Distance (mi)    F Duration        G Fare (£)        H Tariff
      I Start Coords     J End Coords      K Trip ID
      L Destination      M Quoted Low (£)  N Quoted High (£) O Quoted Miles
      P Estimated Miles  Q Gap Events
    ============================================================  */

var TRIP_ID_COL   = 11;   /* column K. Change this too if the sheet is ever reordered. */
var TRIP_DATE_COL = 4;    /* column D. Must hold real dates or Summary reads zero. */

var TRIP_HEADERS = [
  "Sync Timestamp", "Driver ID", "Driver Name", "Trip Date",
  "Distance (mi)", "Duration", "Fare (£)", "Tariff",
  "Start Coords", "End Coords", "Trip ID",
  "Destination", "Quoted Low (£)", "Quoted High (£)", "Quoted Miles",
  "Estimated Miles", "Gap Events"
];

var DRIVER_HEADERS = ["Driver ID", "Name", "Active (TRUE/FALSE)"];
var LOG_HEADERS    = ["Timestamp", "Driver ID", "Driver Name", "Action", "App Timestamp"];
var VALID_TARIFFS  = ["Day", "Night", "Holiday", "Extra"];

var TOKEN_HOURS = 12;     /* a session covers a shift, then the driver re-verifies */

/* ============================================================
   Menu
   ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Taxi Meter")
    .addItem("Check headers", "showHeaders")
    .addItem("Fix headers", "fixHeaders")
    .addItem("Repair trip dates", "repairTripDates")
    .addSeparator()
    .addItem("Build or rebuild Summary tab", "buildSummary")
    .addSeparator()
    .addItem("Deactivate unnamed driver IDs", "deactivateUnnamedDrivers")
    .addItem("Generate 10 new driver IDs", "generateBulkIds")
    .addSeparator()
    .addItem("Generate app secret", "generateAppSecret")
    .addItem("Set OpenRouteService key", "setOrsKey")
    .addItem("Check security setup", "checkSecuritySetup")
    .addToUi();
}

/* ============================================================
   Secrets and sessions

   Both values live in Script properties. They are readable by this
   script and by you, and are never sent to the browser. That is the
   whole reason the routing key moved here.
   ============================================================ */

function props_() { return PropertiesService.getScriptProperties(); }

/* A signed, time limited proof that an ID passed a real check. Editing or
   forging one fails the signature, so the only way to hold a working token
   is to have verified a genuine, active Driver ID. */
function makeToken_(driverId) {
  var secret = props_().getProperty("APP_SECRET");
  if (!secret) throw new Error("APP_SECRET is not set. Run Taxi Meter → Generate app secret.");

  var payload = driverId + "|" + (Date.now() + TOKEN_HOURS * 3600000);
  var sig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(payload, secret));
  return Utilities.base64EncodeWebSafe(payload) + "." + sig;
}

/* Returns the driver ID if the token is genuine and unexpired, otherwise null. */
function readToken_(token) {
  if (!token || String(token).indexOf(".") < 0) return null;

  var secret = props_().getProperty("APP_SECRET");
  if (!secret) return null;

  var parts = String(token).split(".");
  var payload;
  try {
    payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
  } catch (err) {
    return null;
  }

  var expect = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(payload, secret));
  if (expect !== parts[1]) return null;

  var bits = payload.split("|");
  if (!bits[1] || Number(bits[1]) < Date.now()) return null;
  return bits[0];
}

function badToken_() {
  return respond({
    status: "error", code: "badtoken",
    message: "Session expired. Verify your Driver ID again."
  });
}

/* Crude but effective counter, keyed per hour. Apps Script web apps cannot
   see the caller's IP, so anonymous traffic can only be capped in total.
   Traffic carrying a valid token is capped per driver instead. */
function rateLimitOk_(id, bucket, perHour) {
  var cache = CacheService.getScriptCache();
  var key   = "rl_" + bucket + "_" + id + "_" + Math.floor(Date.now() / 3600000);
  var n     = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(n), 3700);
  return n <= perHour;
}

/* ============================================================
   Web endpoint
   ============================================================ */

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var data;

  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond({ status: "error", message: "Invalid JSON payload" });
  }

  var type = data.type || "sync";

  if (type === "validate")   return handleValidate(sheet, data);
  if (type === "route")      return handleRoute(sheet, data);
  if (type === "instantLog") return handleInstantLog(sheet, data);
  if (type === "sync")       return handleSync(sheet, data);

  return respond({ status: "error", message: "Unknown request type: " + type });
}

/* validate is the only open endpoint. It has to be, because it is how a
   driver gets in. Everything else now requires the token it hands out. */
function handleValidate(sheet, data) {
  var driverId = (data.driverId || "").trim().toUpperCase();

  if (!driverId) {
    return respond({ valid: false, active: false, name: "", message: "No driver ID supplied" });
  }

  var check = lookupDriver(sheet, driverId);

  if (!check.valid) {
    /* Throttle misses only, so the endpoint cannot be used to sweep the ID
       space and learn which IDs exist. Genuine drivers never see this. */
    if (!rateLimitOk_("global", "validatefail", 60)) {
      return respond({ valid: false, active: false, name: "", message: "Too many attempts. Try again later." });
    }
    return respond({ valid: false, active: false, name: "" });
  }

  /* An ID with no name against it has been generated but never issued.
     It must not start a trip, or fares land on the sheet with no owner. */
  if (!check.name) {
    return respond({
      valid: true, active: false, name: "",
      message: "This ID has not been assigned to a driver yet. Contact your manager."
    });
  }

  if (!check.active) {
    return respond({
      valid: true, active: false, name: check.name,
      message: "Account suspended. Contact your manager."
    });
  }

  return respond({ valid: true, active: true, name: check.name, token: makeToken_(driverId) });
}

/* The browser asks for a route, this asks OpenRouteService on its behalf.
   The key never leaves the server, and the token stops the public URL being
   used as a free routing service on your quota. */
function handleRoute(sheet, data) {
  var driverId = readToken_(data.token);
  if (!driverId) return badToken_();

  var check = lookupDriver(sheet, driverId);
  if (!check.valid || !check.active || !check.name) {
    return respond({ status: "error", code: "badtoken", message: "Driver account is no longer active." });
  }

  if (!rateLimitOk_(driverId, "route", 120)) {
    return respond({ status: "error", message: "Too many route requests this hour. Try again shortly." });
  }

  var key = props_().getProperty("ORS_KEY");
  if (!key) {
    return respond({ status: "error", message: "No routing key on the server. Run Taxi Meter → Set OpenRouteService key." });
  }

  if (!data.from || !data.to) {
    return respond({ status: "error", message: "Route request was missing its start or end point." });
  }

  var res = UrlFetchApp.fetch(
    "https://api.openrouteservice.org/v2/directions/driving-car/geojson",
    {
      method:      "post",
      contentType: "application/json",
      headers:     { "Authorization": key },
      payload: JSON.stringify({
        coordinates: [
          [Number(data.from.lng), Number(data.from.lat)],
          [Number(data.to.lng),   Number(data.to.lat)]
        ],
        instructions: true,
        units: "m"
      }),
      muteHttpExceptions: true
    });

  var code = res.getResponseCode();
  if (code !== 200) {
    return respond({
      status: "error",
      message: "Routing service returned HTTP " + code + ". Check the key and its daily quota."
    });
  }

  return respond({ status: "ok", route: JSON.parse(res.getContentText()) });
}

function handleInstantLog(sheet, data) {
  var driverId = readToken_(data.token);
  if (!driverId) return badToken_();

  var logSheet = sheet.getSheetByName("Log");
  if (!logSheet) {
    return respond({ status: "error", message: "Log sheet not found" });
  }

  if (!rateLimitOk_(driverId, "log", 200)) {
    return respond({ status: "error", message: "Too many log entries this hour." });
  }

  var lock     = LockService.getScriptLock();
  var acquired = false;
  try {
    lock.waitLock(10000);
    acquired = true;
    logSheet.appendRow([
      Utilities.formatDate(new Date(), "Europe/London", "dd/MM/yyyy, HH:mm:ss"),
      driverId,
      data.driverName || "",
      data.action     || "",
      data.timestamp  || ""
    ]);
  } catch (err) {
    return respond({ status: "error", message: "Log busy, entry not written" });
  } finally {
    if (acquired) lock.releaseLock();
  }

  return respond({ status: "ok" });
}

function handleSync(sheet, data) {
  /* The driver ID comes from the signed token, never from the payload, so a
     caller cannot file trips under someone else's name. */
  var driverId = readToken_(data.token);
  if (!driverId) return badToken_();

  var tripsSheet = sheet.getSheetByName("Trips");
  if (!tripsSheet) {
    return respond({ status: "error", message: "Trips sheet not found" });
  }

  var driverCheck = lookupDriver(sheet, driverId);
  if (!driverCheck.valid)  return respond({ status: "error", message: "Driver ID not recognised" });
  if (!driverCheck.name)   return respond({ status: "error", message: "Driver ID has not been assigned to a driver" });
  if (!driverCheck.active) return respond({ status: "error", message: "Driver account is inactive" });

  var driverNm = driverCheck.name;
  var trips    = data.trips || [];

  if (trips.length > 200) {
    return respond({ status: "error", message: "Too many trips in one sync. Send 200 or fewer." });
  }

  var inserted  = 0;
  var skipped   = 0;
  var errors    = [];
  var syncedIds = [];   /* only these get marked synced on the phone */

  /* One lock around the whole batch. Without it, two drivers syncing in the
     same second can read the same duplicate list and write the same trip twice. */
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return respond({ status: "error", message: "Sheet busy. Try syncing again in a moment." });
  }

  try {
    var existingTripIds = getExistingTripIds(tripsSheet);
    var rowsToWrite = [];

    trips.forEach(function (trip, idx) {
      var validation = validateTrip(trip, idx);
      if (validation.error) {
        errors.push(validation.error);
        return;                        /* stays queued on the phone, deliberately */
      }

      var tripId = String(trip.tripId || "").trim();
      if (tripId && existingTripIds[tripId]) {
        skipped++;
        syncedIds.push(tripId);        /* already on the sheet, so it is done */
        return;
      }

      var miles  = parseFloat(trip.miles) || 0;
      var fare   = parseFloat(trip.fare)  || 0;
      var tariff = VALID_TARIFFS.indexOf(trip.tariff) !== -1 ? trip.tariff : "Day";

      var startStr = trip.start ? trip.start.lat + ", " + trip.start.lng : "N/A";
      var endStr   = trip.end   ? trip.end.lat   + ", " + trip.end.lng   : "N/A";

      /* A real Date, not the phone's text. Every date formula on the Summary
         tab compares against dates, and text never matches one. */
      var tripDate = parseUkDate_(trip.date);

      rowsToWrite.push([
        Utilities.formatDate(new Date(), "Europe/London", "dd/MM/yyyy, HH:mm:ss"),
        driverId,
        driverNm,
        tripDate || String(trip.date || ""),
        round2(miles),
        trip.time || "00:00:00",
        round2(fare),
        tariff,
        startStr,
        endStr,
        tripId || "",
        String(trip.destination || ""),
        numOrBlank(trip.quotedLow),
        numOrBlank(trip.quotedHigh),
        numOrBlank(trip.quotedMiles),
        numOrBlank(trip.estimatedMiles),
        numOrBlank(trip.gapEvents)
      ]);

      if (tripId) {
        existingTripIds[tripId] = true;
        syncedIds.push(tripId);
      }
      inserted++;
    });

    /* Single block write. Faster than appendRow per trip, and a timeout
       part way through cannot leave half a trip on the sheet. */
    if (rowsToWrite.length) {
      var startRow = tripsSheet.getLastRow() + 1;
      tripsSheet.getRange(startRow, 1, rowsToWrite.length, TRIP_HEADERS.length)
                .setValues(rowsToWrite);
      tripsSheet.getRange(startRow, TRIP_DATE_COL, rowsToWrite.length, 1)
                .setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
  } finally {
    lock.releaseLock();
  }

  return respond({
    status: "ok", inserted: inserted, skipped: skipped,
    errors: errors, syncedIds: syncedIds
  });
}

/* ============================================================
   Shared helpers
   ============================================================ */

function lookupDriver(sheet, driverId) {
  var driversSheet = sheet.getSheetByName("Drivers");
  if (!driversSheet) return { valid: false, active: false, name: "" };

  var rows = driversSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rowId = String(rows[i][0]).trim().toUpperCase();
    if (rowId === driverId) {
      var isActive = (rows[i][2] === true || String(rows[i][2]).trim().toUpperCase() === "TRUE");
      return { valid: true, active: isActive, name: String(rows[i][1] || "").trim() };
    }
  }
  return { valid: false, active: false, name: "" };
}

/* The phone sends "09/08/2026, 06:40:12". JavaScript reads that as month
   nine, day eight, or as nothing at all, so it is parsed by hand. */
function parseUkDate_(s) {
  var m = String(s || "").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}),?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;

  var d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]),
                   Number(m[4]), Number(m[5]), Number(m[6] || 0));
  return isNaN(d.getTime()) ? null : d;
}

function getExistingTripIds(tripsSheet) {
  var map  = {};
  var last = tripsSheet.getLastRow();
  if (last < 2) return map;

  var ids = tripsSheet.getRange(2, TRIP_ID_COL, last - 1, 1).getValues();
  ids.forEach(function (row) {
    var id = String(row[0]).trim();
    if (id) map[id] = true;
  });
  return map;
}

function validateTrip(trip, idx) {
  var label = "Trip[" + idx + "]";

  if (!trip || typeof trip !== "object") {
    return { error: label + ": not an object" };
  }

  var miles = parseFloat(trip.miles);
  if (isNaN(miles) || miles < 0 || miles > 500) {
    return { error: label + ": invalid miles (" + trip.miles + ")" };
  }

  var fare = parseFloat(trip.fare);
  if (isNaN(fare) || fare < 0 || fare > 2000) {
    return { error: label + ": invalid fare (" + trip.fare + ")" };
  }

  if (trip.tariff && VALID_TARIFFS.indexOf(trip.tariff) === -1) {
    return { error: label + ": invalid tariff (" + trip.tariff + ")" };
  }

  var est = parseFloat(trip.estimatedMiles);
  if (!isNaN(est) && est > miles + 0.05) {
    return { error: label + ": estimated miles (" + est + ") exceeds trip distance (" + miles + ")" };
  }

  return { error: null };
}

function round2(n)      { return Math.round((parseFloat(n) || 0) * 100) / 100; }
function numOrBlank(v)  {
  if (v === undefined || v === null || v === "") return "";
  var n = parseFloat(v);
  return isNaN(n) ? "" : n;
}
function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
                       .setMimeType(ContentService.MimeType.JSON);
}
function toast(msg) {
  try { SpreadsheetApp.getUi().alert(msg); }
  catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Taxi Meter", 10);
    Logger.log(msg);
  }
}

/* ============================================================
   Security setup
   ============================================================ */

function generateAppSecret() {
  var ui = SpreadsheetApp.getUi();

  if (props_().getProperty("APP_SECRET")) {
    var r = ui.alert("An app secret already exists",
      "Replacing it signs every driver out immediately and they will each have to " +
      "verify again. Trips already on the sheet are unaffected.\n\nReplace it?",
      ui.ButtonSet.YES_NO);
    if (r !== ui.Button.YES) { toast("Nothing changed."); return; }
  }

  /* No cryptographic RNG is exposed to Apps Script. Four concatenated UUIDs
     give 128 hex characters, which is ample for signing session tokens. */
  var secret = (Utilities.getUuid() + Utilities.getUuid() +
                Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "");
  props_().setProperty("APP_SECRET", secret);

  toast("App secret stored. You never need to see or copy it.\n\n" +
        "Redeploy afterwards: Deploy → Manage deployments → edit → New version.");
}

function setOrsKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt("OpenRouteService key",
    "Paste a NEWLY GENERATED key.\n\n" +
    "Delete the old one at openrouteservice.org first. It was readable in the " +
    "deployed page, so it must be treated as public.",
    ui.ButtonSet.OK_CANCEL);

  if (res.getSelectedButton() !== ui.Button.OK) return;

  var key = res.getResponseText().trim();
  if (!key) { toast("No key entered. Nothing changed."); return; }

  props_().setProperty("ORS_KEY", key);
  toast("Routing key stored in Script properties. It is never sent to the browser.\n\n" +
        "Redeploy afterwards: Deploy → Manage deployments → edit → New version.");
}

function checkSecuritySetup() {
  var p   = props_();
  var out = [];

  out.push(p.getProperty("APP_SECRET")
    ? "App secret: set. Sign-in can issue session tokens."
    : "App secret: MISSING. Nobody can sign in. Run Generate app secret.");

  out.push(p.getProperty("ORS_KEY")
    ? "Routing key: set. Route requests are proxied through this script."
    : "Routing key: MISSING. Routing will fail, the meter still works.");

  out.push("");
  out.push("Neither value is ever sent to the browser.");
  out.push("After changing either: Deploy → Manage deployments → edit → New version.");

  toast(out.join("\n"));
}

/* ============================================================
   Headers: check and repair
   ============================================================ */

function showHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];

  out.push(checkSheetHeaders(ss, "Trips",   TRIP_HEADERS));
  out.push(checkSheetHeaders(ss, "Drivers", DRIVER_HEADERS));
  out.push(checkSheetHeaders(ss, "Log",     LOG_HEADERS));

  var report = out.join("\n\n");
  Logger.log(report);
  toast(report);
}

function checkSheetHeaders(ss, name, expected) {
  var sh = ss.getSheetByName(name);
  if (!sh) return name + ": sheet not found.";

  var row1 = sh.getRange(1, 1, 1, sh.getMaxColumns()).getValues()[0];
  var bad  = [];

  for (var i = 0; i < expected.length; i++) {
    var actual = (i < row1.length) ? String(row1[i]).trim() : "(no column)";
    if (actual !== expected[i]) {
      bad.push("  " + colLetter(i + 1) + ": expected '" + expected[i] + "', found '" + actual + "'");
    }
  }

  if (!bad.length) return name + ": all " + expected.length + " headers correct.";
  return name + ": " + bad.length + " header(s) wrong.\n" + bad.join("\n");
}

/* Overwrites row 1 on all three sheets. Safe because every writer in this
   script addresses columns by position, so the header row carries no logic. */
function fixHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];

  out.push(writeHeaders(ss, "Trips",   TRIP_HEADERS));
  out.push(writeHeaders(ss, "Drivers", DRIVER_HEADERS));
  out.push(writeHeaders(ss, "Log",     LOG_HEADERS));

  var idCheck = ss.getSheetByName("Trips");
  if (idCheck) {
    var v = String(idCheck.getRange(1, TRIP_ID_COL).getValue()).trim();
    out.push(v === "Trip ID"
      ? "Duplicate checking is aligned on column " + TRIP_ID_COL + "."
      : "WARNING: column " + TRIP_ID_COL + " reads '" + v + "'. Duplicate checking is misaligned.");
  }

  var report = out.join("\n");
  Logger.log(report);
  toast(report);
}

function writeHeaders(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) return name + ": sheet not found, skipped.";

  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  }

  /* If the sheet has data but no header row, row 1 is a real record and
     overwriting it would destroy it. Detect that before writing. */
  if (sh.getLastRow() > 0) {
    var row1 = sh.getRange(1, 1, 1, Math.min(3, sh.getMaxColumns())).getValues()[0];
    var looksLikeData = row1.some(function (c) {
      return String(c || "").trim().toUpperCase().indexOf("TX-") === 0;
    });
    if (looksLikeData) {
      return name + ": STOPPED. Row 1 looks like data, not headers. Insert a blank row at the top first.";
    }
  }

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sh.setFrozenRows(1);
  return name + ": " + headers.length + " headers written.";
}

/* Converts the Trip Date column from the phone's text into real dates.
   Run once. QUERY picks one type per column, so a column that is mostly
   text makes toDate(D) fail for every row, including the new ones. */
function repairTripDates() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips");
  if (!sh) { toast("No Trips sheet found."); return; }

  var last = sh.getLastRow();
  if (last < 2) { toast("No trip rows to repair."); return; }

  var range   = sh.getRange(2, TRIP_DATE_COL, last - 1, 1);
  var values  = range.getValues();
  var fixed   = 0, already = 0, blank = 0, failed = 0;
  var samples = [];

  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];

    if (v instanceof Date)              { already++; continue; }
    if (v === "" || v === null)         { blank++;   continue; }

    var d = parseUkDate_(v);
    if (d) { values[i][0] = d; fixed++; }
    else   { failed++; if (samples.length < 3) samples.push("row " + (i + 2) + ": '" + v + "'"); }
  }

  range.setValues(values);
  range.setNumberFormat("dd/MM/yyyy HH:mm:ss");

  var msg = "Trip dates repaired.\n\n" +
            fixed   + " converted from text\n" +
            already + " already dates\n" +
            blank   + " blank\n" +
            failed  + " could not be read";
  if (samples.length) msg += "\n\nUnreadable examples:\n" + samples.join("\n");
  toast(msg);
}

function colLetter(n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}

/* ============================================================
   Summary tab
   ============================================================ */

/* Writes formulas, not values, so the tab recalculates on every sync
   with no need to run anything again. Safe to re-run at any time.
   Everything here depends on column D holding real dates. If Today,
   This week and This month all read £0.00 while Lifetime is right,
   run Repair trip dates. */
function buildSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Trips")) { toast("No Trips sheet found."); return; }

  var sh = ss.getSheetByName("Summary");
  if (sh) sh.clear(); else sh = ss.insertSheet("Summary");

  var monday = "TODAY()-WEEKDAY(TODAY(),3)";

  sh.getRange("A1").setValue("Taxi Meter Summary").setFontSize(16).setFontWeight("bold");
  sh.getRange("A2").setValue("Live formulas. This updates itself after every sync.").setFontColor("#666666");

  sh.getRange("A4").setValue("Headline").setFontWeight("bold");
  sh.getRange("B4").setValue("Value").setFontWeight("bold");

  var labels = [
    ["Today"], ["This week from Monday"], ["This month"], ["Lifetime"],
    ["Trips recorded"], ["Average fare"], ["Total miles"],
    ["Estimated miles in background"], ["Gap events"]
  ];
  sh.getRange(5, 1, labels.length, 1).setValues(labels);

  sh.getRange(5, 2, 9, 1).setFormulas([
    ['=SUMIFS(Trips!G2:G,Trips!D2:D,">="&TODAY(),Trips!D2:D,"<"&TODAY()+1)'],
    ['=SUMIFS(Trips!G2:G,Trips!D2:D,">="&' + monday + ',Trips!D2:D,"<"&' + monday + '+7)'],
    ['=SUMIFS(Trips!G2:G,Trips!D2:D,">="&EOMONTH(TODAY(),-1)+1,Trips!D2:D,"<"&EOMONTH(TODAY(),0)+1)'],
    ['=SUM(Trips!G2:G)'],
    ['=COUNTA(Trips!B2:B)'],
    ['=IFERROR(AVERAGE(Trips!G2:G),0)'],
    ['=SUM(Trips!E2:E)'],
    ['=SUM(Trips!P2:P)'],
    ['=SUM(Trips!Q2:Q)']
  ]);

  sh.getRange("B5:B8").setNumberFormat('£#,##0.00');
  sh.getRange("B10").setNumberFormat('£#,##0.00');
  sh.getRange("B11:B12").setNumberFormat('0.00');

  sh.getRange("A15").setValue("Daily totals").setFontWeight("bold");
  sh.getRange("A16").setFormula(
    '=IFERROR(QUERY(Trips!A2:Q,' +
    '"select toDate(D), B, C, count(G), sum(G), sum(E) ' +
    'where B is not null ' +
    'group by toDate(D), B, C ' +
    'order by toDate(D) desc ' +
    "label toDate(D) 'Date', B 'Driver ID', C 'Driver', count(G) 'Trips', sum(G) 'Fare', sum(E) 'Miles' " +
    "format toDate(D) 'dd/MM/yyyy', sum(G) '£#,##0.00', sum(E) '0.00'\"" +
    ',0),"No trip data yet")'
  );

  sh.getRange("H15").setValue("Weekly totals, week beginning Monday").setFontWeight("bold");
  sh.getRange("H16").setFormula(
    '=IFERROR(QUERY({ARRAYFORMULA(IF(Trips!D2:D="","",INT(Trips!D2:D)-WEEKDAY(Trips!D2:D,3))),' +
    'Trips!B2:B,Trips!G2:G,Trips!E2:E},' +
    '"select Col1, Col2, count(Col3), sum(Col3), sum(Col4) ' +
    'where Col2 is not null ' +
    'group by Col1, Col2 ' +
    'order by Col1 desc ' +
    "label Col1 'Week beginning', Col2 'Driver ID', count(Col3) 'Trips', sum(Col3) 'Fare', sum(Col4) 'Miles' " +
    "format Col1 'dd/MM/yyyy', sum(Col3) '£#,##0.00', sum(Col4) '0.00'\"" +
    ',0),"No trip data yet")'
  );

  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 140);
  sh.setFrozenRows(4);

  toast("Summary tab built. It refreshes itself from now on.\n\n" +
        "If the dated totals read zero, run Taxi Meter → Repair trip dates.");
}

/* ============================================================
   Driver ID management
   ============================================================ */

/* Any generated but unissued ID is a live key to the app. This turns
   every nameless row off. Issue one by adding the name and setting TRUE. */
function deactivateUnnamedDrivers() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Drivers");
  if (!sh) { toast("No Drivers sheet found."); return; }

  var last = sh.getLastRow();
  if (last < 2) { toast("No driver rows found."); return; }

  var rows    = sh.getRange(2, 1, last - 1, 3).getValues();
  var changed = 0;
  var named   = 0;

  for (var i = 0; i < rows.length; i++) {
    var id   = String(rows[i][0] || "").trim();
    var name = String(rows[i][1] || "").trim();
    if (!id) continue;
    if (name) { named++; continue; }
    if (String(rows[i][2]).trim().toUpperCase() !== "FALSE") {
      rows[i][2] = "FALSE";
      changed++;
    }
  }

  sh.getRange(2, 1, rows.length, 3).setValues(rows);
  toast(changed + " unnamed ID(s) set to FALSE. " + named + " named driver(s) left untouched.");
}

function generateDriverId() {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var id = "TX-";
  for (var i = 0; i < 5; i++) id += chars.charAt(Math.floor(Math.random() * chars.length));
  Logger.log("Generated ID: " + id);
  return id;
}

/* New IDs are written inactive. Add a name and set Active to TRUE to issue one. */
function generateBulkIds() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Drivers");
  if (!sh) { toast("No Drivers sheet found."); return; }

  var existing = {};
  sh.getDataRange().getValues().forEach(function (r) {
    existing[String(r[0]).trim().toUpperCase()] = true;
  });

  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var added = 0, guard = 0, batch = [];

  while (added < 10 && guard++ < 500) {
    var id = "TX-";
    for (var j = 0; j < 5; j++) id += chars.charAt(Math.floor(Math.random() * chars.length));
    if (existing[id]) continue;          /* never issue the same ID twice */
    existing[id] = true;
    batch.push([id, "", "FALSE"]);
    added++;
  }

  if (batch.length) {
    sh.getRange(sh.getLastRow() + 1, 1, batch.length, 3).setValues(batch);
  }
  toast(added + " new driver ID(s) added, all inactive. Add a name and set Active to TRUE to issue one.");
}

/* Fresh install only. Creates all three sheets with full headers. */
function setupHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var drivers = ss.getSheetByName("Drivers") || ss.insertSheet("Drivers");
  if (drivers.getLastRow() === 0) drivers.appendRow(DRIVER_HEADERS);

  var trips = ss.getSheetByName("Trips") || ss.insertSheet("Trips");
  if (trips.getLastRow() === 0) trips.appendRow(TRIP_HEADERS);

  var log = ss.getSheetByName("Log") || ss.insertSheet("Log");
  if (log.getLastRow() === 0) log.appendRow(LOG_HEADERS);

  fixHeaders();
}
