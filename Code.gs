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

    Trips headings, in the order a fresh sheet gets them. From sheet 1.2.0
    every column is found by its heading, so the office may move, insert or
    add columns freely. Only renaming a heading matters, and the heading rows
    are protected (Taxi Meter → Protect the sheet).
      A Sync Timestamp   B Driver ID       C Driver Name     D Trip Date
      E Distance (mi)    F Duration        G Fare (£)        H Tariff
      I Start Coords     J End Coords      K Trip ID
      L Destination      M Quoted Low (£)  N Quoted High (£) O Quoted Miles
      P Estimated Miles  Q Gap Events
      R Metered Fare (£) S Extras (£)      T Extras Detail   U Agreed Fare (£)
      V Running Miles    W Waiting Time    X Vehicle Reg

    G "Fare (£)" is what the passenger actually paid: the metered fare, or an
    agreed fare where one was struck, PLUS any permitted additional charges.
    That keeps every existing Summary formula meaning what it always meant.
    R breaks out the meter reading on its own, S and T the extras, U the agreed
    fare where one applied. Columns R to W are added automatically on first sync
    against an older sheet, so no manual migration is needed.
    ============================================================  */

/* Printed on the phone's Settings screen beside the app's own version, so
   "is the new sheet code live" can be answered from the car. Bump on every
   change to this file. (The minibus app's "versions on the foot".) */
var SHEET_VERSION = "1.2.0";


var TRIP_HEADERS = [
  "Sync Timestamp", "Driver ID", "Driver Name", "Trip Date",
  "Distance (mi)", "Duration", "Fare (£)", "Tariff",
  "Start Coords", "End Coords", "Trip ID",
  "Destination", "Quoted Low (£)", "Quoted High (£)", "Quoted Miles",
  "Estimated Miles", "Gap Events",
  "Metered Fare (£)", "Extras (£)", "Extras Detail", "Agreed Fare (£)",
  "Running Miles", "Waiting Time",
  /* From sheet 1.2.0: the registration the driver entered on the phone for
     this trip. Added at the end of an existing sheet on the first sync. */
  "Vehicle Reg"
];


/* Permitted additional charges, from the council card. Soiling is a fixed figure;
   tolls vary, so they are free entry on the phone and only bounded here. */
var MAX_EXTRA_GBP       = 200;
var MAX_EXTRAS_PER_TRIP = 8;

var DRIVER_HEADERS = ["Driver ID", "Name", "Active (TRUE/FALSE)"];
var LOG_HEADERS    = ["Timestamp", "Driver ID", "Driver Name", "Action", "App Timestamp"];

/* "Peak" is the P rate on the council card. "Holiday" is kept so trips queued on a
   phone running the previous build still sync instead of being rejected. */
var VALID_TARIFFS  = ["Day", "Night", "Peak", "Extra", "Holiday"];

var TOKEN_HOURS = 12;     /* a session covers a shift, then the driver re-verifies */

/* ============================================================
   Menu
   ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Taxi Meter")
    .addItem("Check headers", "showHeaders")
    .addItem("Add missing headers", "fixHeaders")
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
    .addSeparator()
    .addItem("Create missing sheets (fresh install)", "setupHeaders")
    .addSeparator()
    .addItem("Protect the sheet", "lockSheet")
    .addItem("Unprotect the sheet", "unlockSheet")
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

/* FIX A. Any exception escaping a handler makes Apps Script return its own HTML
   error page. The phone then reports a parse failure and blames the deployment,
   which sends you looking in the wrong place. Everything is wrapped so the client
   always receives JSON, and the real reason is carried inside it. */
/* Opening the deployment URL in a browser used to return an Apps Script error
   page, which is exactly the confusion describeNonJson() on the phone exists to
   untangle. Now it answers plainly, so "is the new version live" is one tap. */
function doGet() {
  var p = props_();
  return respond({
    status:   "ok",
    service:  "taxi-meter",
    version:  SHEET_VERSION,
    columns:  TRIP_HEADERS.length,
    secret:   !!p.getProperty("APP_SECRET"),
    routing:  !!p.getProperty("ORS_KEY"),
    scriptTz: Session.getScriptTimeZone(),
    sheetTz:  SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
  });
}

function doPost(e) {
  try {
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

  } catch (fatal) {
    try { Logger.log("doPost fatal: " + (fatal && fatal.stack ? fatal.stack : fatal)); } catch (ignored) {}
    return respond({
      status: "error",
      code: "serverfault",
      message: "Server error: " + (fatal && fatal.message ? fatal.message : String(fatal))
    });
  }
}

/* validate is the only open endpoint. It has to be, because it is how a
   driver gets in. Everything else now requires the token it hands out. */
function handleValidate(sheet, data) {
  var driverId = String(data.driverId || "").trim().toUpperCase();

  if (!driverId) {
    return respond({ valid: false, active: false, name: "", message: "No driver ID supplied" });
  }

  /* FIX B. The throttle used to run only after lookupDriver had already read the
     whole Drivers sheet, so a flood of junk IDs still burned read quota and could
     stall genuine sign-ins. A total ceiling on validate traffic is checked first,
     and the tighter miss-only counter still runs afterwards. */
  if (!rateLimitOk_("global", "validateall", 600)) {
    return respond({ valid: false, active: false, name: "", message: "Server busy. Try again in a minute." });
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

  return respond({ valid: true, active: true, name: check.name, token: makeToken_(driverId),
                   version: SHEET_VERSION });
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

  /* Parsing a large GeoJSON only to re-encode it wastes execution time and turns
     a non-JSON 200 from the routing service into an unhelpful server fault. */
  return ContentService
    .createTextOutput('{"status":"ok","route":' + res.getContentText() + '}')
    .setMimeType(ContentService.MimeType.JSON);
}

function handleInstantLog(sheet, data) {
  var driverId = readToken_(data.token);
  if (!driverId) return badToken_();

  /* The logged name used to be whatever the payload claimed, so the audit trail
     could disagree with the Drivers sheet. The ID already comes from the signed
     token; the name now resolves from the same place, with no client fallback. */
  var whoIs = lookupDriver(sheet, driverId);
  if (!whoIs.valid || !whoIs.active) {
    return respond({ status: "error", code: "badtoken", message: "Driver account is no longer active." });
  }

  var logSheet = sheet.getSheetByName("Log");
  if (!logSheet) {
    Logger.log("instantLog: Log sheet missing, entry dropped for " + driverId);
    return respond({ status: "error", message: "Log sheet not found" });
  }

  if (!rateLimitOk_(driverId, "log", 200)) {
    Logger.log("instantLog: rate limited, entry dropped for " + driverId);
    return respond({ status: "error", message: "Too many log entries this hour." });
  }

  var lock     = LockService.getScriptLock();
  var acquired = false;
  try {
    lock.waitLock(10000);
    acquired = true;
    writeRowsByHeader_(logSheet, ensureHeaders_(logSheet, LOG_HEADERS), [{
      "Timestamp":     Utilities.formatDate(new Date(), "Europe/London", "dd/MM/yyyy, HH:mm:ss"),
      "Driver ID":     driverId,
      "Driver Name":   whoIs.name,
      "Action":        sanitiseCell_(data.action),
      "App Timestamp": sanitiseCell_(data.timestamp)
    }]);
  } catch (err) {
    Logger.log("instantLog: lock timeout, entry dropped for " + driverId);
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

  if (!Array.isArray(trips)) {
    return respond({ status: "error", message: "trips must be an array" });
  }
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
    var cols = ensureHeaders_(tripsSheet, TRIP_HEADERS);
    var existingTripIds = getExistingTripIds(tripsSheet, cols);
    var rowsToWrite = [];

    trips.forEach(function (trip, idx) {
      var validation = validateTrip(trip, idx);
      if (validation.error) {
        errors.push(validation.error);
        return;                        /* stays queued on the phone, deliberately */
      }

      var tripId = String(trip.tripId || "").trim();
      var dupKey = driverId + "|" + tripId;
      if (existingTripIds[dupKey]) {
        skipped++;
        syncedIds.push(tripId);        /* already on the sheet, so it is done */
        return;
      }

      var miles  = parseFloat(trip.miles) || 0;
      var tariff = VALID_TARIFFS.indexOf(trip.tariff) !== -1 ? trip.tariff : "Day";

      var startStr = trip.start ? trip.start.lat + ", " + trip.start.lng : "N/A";
      var endStr   = trip.end   ? trip.end.lat   + ", " + trip.end.lng   : "N/A";

      /* validateTrip has already refused anything unparseable, so this is a real
         Date. Text in the date column makes the Summary tab's weekly QUERY fail
         outright and report "No trip data yet", which reads as an empty week
         rather than a broken report. */
      var tripDate = parseUkDate_(trip.date);

      var ex      = summariseExtras_(trip.extras);
      var metered = round2(parseFloat(trip.meteredFare));
      var agreed  = numOrBlank(trip.agreedFare);
      /* What the passenger actually paid. An agreed fare replaces the meter
         reading; permitted extras are added on top of whichever applied. */
      var charged = round2((agreed === "" ? metered : Number(agreed)) + ex.total);

      var row = {};
      row["Sync Timestamp"]   = Utilities.formatDate(new Date(), "Europe/London", "dd/MM/yyyy, HH:mm:ss");
      row["Driver ID"]        = driverId;
      row["Driver Name"]      = driverNm;
      row["Trip Date"]        = tripDate;
      row["Distance (mi)"]    = round2(miles);
      row["Duration"]         = sanitiseCell_(trip.time || "00:00:00");
      row["Fare (£)"]         = charged;
      row["Tariff"]           = tariff;
      row["Start Coords"]     = startStr;
      row["End Coords"]       = endStr;
      row["Trip ID"]          = tripId;
      row["Destination"]      = sanitiseCell_(trip.destination);
      row["Quoted Low (£)"]   = numOrBlank(trip.quotedLow);
      row["Quoted High (£)"]  = numOrBlank(trip.quotedHigh);
      row["Quoted Miles"]     = numOrBlank(trip.quotedMiles);
      row["Estimated Miles"]  = numOrBlank(trip.estimatedMiles);
      row["Gap Events"]       = numOrBlank(trip.gapEvents);
      row["Metered Fare (£)"] = metered;
      row["Extras (£)"]       = ex.total ? ex.total : "";
      row["Extras Detail"]    = sanitiseCell_(ex.detail);
      row["Agreed Fare (£)"]  = agreed;
      row["Running Miles"]    = numOrBlank(trip.runningMiles);
      row["Waiting Time"]     = sanitiseCell_(trip.waitingTime);
      row["Vehicle Reg"]      = sanitiseCell_(normaliseReg_(trip.vehicleReg));
      rowsToWrite.push(row);

      existingTripIds[dupKey] = true;
      syncedIds.push(tripId);
      inserted++;
    });

    /* Each column found by its heading. One write per run of adjacent
       columns, so a timeout cannot leave half a trip on the sheet. */
    if (rowsToWrite.length) {
      var startRow = writeRowsByHeader_(tripsSheet, cols, rowsToWrite);
      tripsSheet.getRange(startRow, cols["Trip Date"], rowsToWrite.length, 1)
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

  /* By heading, so a column the office inserts on Drivers cannot turn a name
     into an ID or switch everyone's Active flag. */
  var map = headerMap_(driversSheet);
  var cId = map["Driver ID"], cName = map["Name"], cAct = activeCol_(map);
  if (!cId) return { valid: false, active: false, name: "" };

  var rows = driversSheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rowId = String(rows[i][cId - 1]).trim().toUpperCase();
    if (rowId === driverId) {
      var a = cAct ? rows[i][cAct - 1] : "";
      var isActive = (a === true || String(a).trim().toUpperCase() === "TRUE");
      return { valid: true, active: isActive, name: cName ? String(rows[i][cName - 1] || "").trim() : "" };
    }
  }
  return { valid: false, active: false, name: "" };
}

/* The phone sends "09/08/2026, 06:40:12". JavaScript reads that as month
   nine, day eight, or as nothing at all, so it is parsed by hand. */
/* JavaScript turns 31/02 into 3 March and hour 24 into 00:00 the following day.
   Both land a trip on the wrong date, silently. The hour-24 case is not
   hypothetical: some older WebKit builds emit "24:30:00" from toLocaleString
   with hour12:false, which would move every trip in the midnight hour. */
function parseUkDate_(s) {
  var m = String(s || "").match(/^(\d{1,2})\/(\d{1,2})\/(\d{4}),?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;

  var day = Number(m[1]), mon = Number(m[2]), yr = Number(m[3]);
  var hh  = Number(m[4]), mi  = Number(m[5]), ss = Number(m[6] || 0);

  if (mon < 1 || mon > 12 || day < 1 || day > 31) return null;
  if (hh > 23 || mi > 59 || ss > 59)              return null;
  if (yr < 2000 || yr > 2100)                     return null;

  var d = new Date(yr, mon - 1, day, hh, mi, ss);
  if (isNaN(d.getTime())) return null;
  /* Catches 31 February and friends, which the constructor rolls forward. */
  if (d.getMonth() !== mon - 1 || d.getDate() !== day) return null;
  return d;
}

/* Keyed on driver AND trip so the map stays correct under any future ID scheme,
   and built on a null prototype so a trip whose ID happens to be "constructor"
   or "__proto__" is not mistaken for a duplicate and silently dropped. */
function getExistingTripIds(tripsSheet, cols) {
  var map  = Object.create(null);
  var last = tripsSheet.getLastRow();
  if (last < 2) return map;
  cols = cols || headerMap_(tripsSheet);
  var cDrv = cols["Driver ID"], cTrip = cols["Trip ID"];
  if (!cDrv || !cTrip) return map;

  /* Retries are always recent. Reading the whole column keeps the script lock
     held for longer and longer as the sheet grows, for no extra safety. */
  var from = Math.max(2, last - 5000), n = last - from + 1;
  var drv  = tripsSheet.getRange(from, cDrv, n, 1).getValues();
  var ids  = tripsSheet.getRange(from, cTrip, n, 1).getValues();
  for (var i = 0; i < n; i++) {
    var id = String(ids[i][0]).trim();
    if (id) map[String(drv[i][0]).trim().toUpperCase() + "|" + id] = true;
  }
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

  /* A trip ID is the only thing standing between a retry and a duplicate row, so
     it is REQUIRED, not merely validated when present. Without one the phone has
     nothing to match the acknowledgement against and marks the trip synced
     whatever happened here, deleting it from the only other place it exists. */
  var tid = String(trip.tripId || "").trim();
  if (!tid)                                     return { error: label + ": missing trip ID" };
  if (tid.length > 64 || /[\r\n\t]/.test(tid)) return { error: label + ": malformed trip ID" };

  /* Text in the date column breaks the Summary tab quietly: the weekly QUERY
     fails outright and reports "No trip data yet", which reads as an empty week
     rather than a broken report. Refuse it here, where the driver sees the
     message and the trip stays safely queued. */
  if (!parseUkDate_(trip.date)) {
    return { error: label + ": unreadable trip date (" + trip.date + ")" };
  }

  if (String(trip.destination || "").length > 300) {
    return { error: label + ": destination text is too long" };
  }

  var meter = parseFloat(trip.meteredFare);
  if (isNaN(meter) || meter < 0 || meter > 2000) {
    return { error: label + ": invalid metered fare (" + trip.meteredFare + ")" };
  }

  if (trip.agreedFare !== undefined && trip.agreedFare !== null && trip.agreedFare !== "") {
    var ag = parseFloat(trip.agreedFare);
    if (isNaN(ag) || ag < 0 || ag > 2000) {
      return { error: label + ": invalid agreed fare (" + trip.agreedFare + ")" };
    }
  }

  if (trip.extras !== undefined && trip.extras !== null) {
    if (!Array.isArray(trip.extras)) return { error: label + ": extras must be a list" };
    if (trip.extras.length > MAX_EXTRAS_PER_TRIP) {
      return { error: label + ": too many extras (" + trip.extras.length + ")" };
    }
    for (var x = 0; x < trip.extras.length; x++) {
      var amt = parseFloat(trip.extras[x] && trip.extras[x].amount);
      if (isNaN(amt) || amt <= 0 || amt > MAX_EXTRA_GBP) {
        return { error: label + ": invalid extra amount (" + (trip.extras[x] || {}).amount + ")" };
      }
    }
  }

  var est = parseFloat(trip.estimatedMiles);
  if (!isNaN(est) && est > miles + 0.05) {
    return { error: label + ": estimated miles (" + est + ") exceeds trip distance (" + miles + ")" };
  }

  return { error: null };
}

/* FIX D. Text arriving from the phone lands in a cell. A value that starts with
   = + - or @ is evaluated by Sheets as a formula, so a destination label taken
   from a third party geocoder could execute inside your spreadsheet. Length is
   capped at the same time so one bad row cannot bloat the sheet. */
/* Flattens the extras list into one total and one readable cell. */
function summariseExtras_(list) {
  if (!Array.isArray(list) || !list.length) return { total: 0, detail: "" };
  var total = 0, bits = [];
  for (var i = 0; i < list.length && i < MAX_EXTRAS_PER_TRIP; i++) {
    var amt = parseFloat(list[i] && list[i].amount);
    if (isNaN(amt) || amt <= 0 || amt > MAX_EXTRA_GBP) continue;
    total += amt;
    bits.push(String(list[i].label || "Extra").substring(0, 40) + " £" + amt.toFixed(2));
  }
  return { total: round2(total), detail: bits.join("; ") };
}

function sanitiseCell_(v) {
  var s = (v === undefined || v === null) ? "" : String(v);
  if (s.length > 300) s = s.substring(0, 300);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return s;
}

/* "ab12cde" and "AB12  CDE" both become "AB12 CDE" when they look like a
   current-style UK registration; anything else is upper-cased and kept. */
function normaliseReg_(v) {
  var s = String(v == null ? "" : v).toUpperCase().replace(/[^A-Z0-9 ]/g, "").replace(/\s+/g, " ").trim();
  var m = /^([A-Z]{2}\d{2}) ?([A-Z]{3})$/.exec(s);
  return m ? m[1] + " " + m[2] : s.substring(0, 12);
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

  /* Column A is stamped explicitly in Europe/London; column D is built from the
     ambient script timezone. A bound script inherits the spreadsheet's timezone
     when created but does not follow later changes, and if the two drift, every
     trip datetime shifts and evening trips cross the day boundary. */
  var scriptTz = Session.getScriptTimeZone();
  var sheetTz  = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  out.push("");
  out.push(scriptTz === sheetTz
    ? "Timezone: " + scriptTz + " on both script and spreadsheet."
    : "TIMEZONE MISMATCH. Script is " + scriptTz + ", spreadsheet is " + sheetTz +
      ". Trip dates will be shifted. Fix in File → Settings and Project Settings.");

  /* The minibus sheet's health check names any tab left unlocked. */
  var unlocked = [];
  try { unlocked = locksMissing_(SpreadsheetApp.getActiveSpreadsheet()); } catch (err) {}
  out.push("");
  out.push(unlocked.length
    ? "UNPROTECTED: " + unlocked.join(", ") + ". Run Taxi Meter → Protect the sheet."
    : "Protection: Trips, Log, Summary and the Drivers heading row are locked (warning only).");

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
  out.push("Columns are found by heading, so their order does not matter.");

  var report = out.join("\n\n");
  Logger.log(report);
  toast(report);
}

/* Missing and doubled headings. Order is not checked: it does not matter. */
function checkSheetHeaders(ss, name, expected) {
  var sh = ss.getSheetByName(name);
  if (!sh) return name + ": sheet not found.";
  var map = headerMap_(sh);
  var missing = expected.filter(function (h) {
    return !(map[h] || (name === "Drivers" && h === DRIVER_HEADERS[2] && activeCol_(map)));
  });
  var seen = {}, doubled = [];
  if (sh.getLastColumn() > 0) {
    sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach(function (h) {
      var k = String(h || "").trim();
      if (!k) return;
      if (seen[k] && expected.indexOf(k) >= 0 && doubled.indexOf(k) < 0) doubled.push(k);
      seen[k] = true;
    });
  }
  if (!missing.length && !doubled.length) return name + ": all " + expected.length + " headings present.";
  var out = name + ":";
  if (missing.length) out += "\n  missing: " + missing.join(", ") + " (Add missing headers puts them at the end)";
  if (doubled.length) out += "\n  appears twice: " + doubled.join(", ") + " (the first is used; rename or delete the other)";
  return out;
}

/* Adds any heading a tab lacks, at the end. Never relabels or moves a
   column, because the headings are what every reader and writer goes by. */
function fixHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];

  out.push(writeHeaders(ss, "Trips",   TRIP_HEADERS));
  out.push(writeHeaders(ss, "Drivers", DRIVER_HEADERS));
  out.push(writeHeaders(ss, "Log",     LOG_HEADERS));

  var report = out.join("\n");
  Logger.log(report);
  toast(report);
}

function writeHeaders(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) return name + ": sheet not found, skipped.";

  /* If the sheet has data but no header row, row 1 is a real record and
     writing headings beside it would hide it. Detect that before writing. */
  if (sh.getLastRow() > 0) {
    var row1 = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    var looksLikeData = row1.some(function (c) {
      return String(c || "").trim().toUpperCase().indexOf("TX-") === 0;
    });
    if (looksLikeData) {
      return name + ": STOPPED. Row 1 looks like data, not headers. Insert a blank row at the top first.";
    }
  }

  var before = Object.keys(headerMap_(sh)).length;
  var want = headers.slice();
  /* Drivers' Active column may carry an older wording; do not add a second. */
  if (name === "Drivers" && activeCol_(headerMap_(sh))) want = want.filter(function (h) { return h !== DRIVER_HEADERS[2]; });
  var after = Object.keys(ensureHeaders_(sh, want)).length;
  return name + ": " + (after - before ? (after - before) + " heading(s) added at the end." : "nothing missing.");
}

/* Converts the Trip Date column from the phone's text into real dates.
   Run once. QUERY picks one type per column, so a column that is mostly
   text makes toDate(D) fail for every row, including the new ones. */
function repairTripDates() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips");
  if (!sh) { toast("No Trips sheet found."); return; }

  var last = sh.getLastRow();
  if (last < 2) { toast("No trip rows to repair."); return; }
  var col = headerMap_(sh)["Trip Date"];
  if (!col) { toast("The Trips tab has no 'Trip Date' column."); return; }

  var values  = sh.getRange(2, col, last - 1, 1).getValues();
  var fixed   = 0, already = 0, blank = 0, failed = 0;
  var samples = [];

  for (var i = 0; i < values.length; i++) {
    var v = values[i][0];

    if (v instanceof Date)      { already++; continue; }
    if (v === "" || v === null) { blank++;   continue; }

    var d = parseUkDate_(v);
    if (d) {
      /* Only cells that actually changed are written back. Writing the whole
         column would replace any formula in it with its current value. */
      sh.getRange(i + 2, col).setValue(d).setNumberFormat("dd/MM/yyyy HH:mm:ss");
      fixed++;
    } else {
      failed++;
      if (samples.length < 5) samples.push("row " + (i + 2) + ": '" + v + "'");
    }
  }

  var msg = failed
    ? "Trip dates PARTIALLY repaired.\n\n" + failed + " row(s) are still text. Until those are " +
      "corrected by hand, Today, This week and This month stay wrong and the weekly table on " +
      "Summary reads 'No trip data yet'.\n\n"
    : "Trip dates repaired.\n\n";
  msg += fixed   + " converted from text\n" +
         already + " already dates\n" +
         blank   + " blank\n" +
         failed  + " could not be read";
  if (samples.length) msg += "\n\nUnreadable examples:\n" + samples.join("\n");
  toast(msg);
}



/* ============================================================
   Summary tab
   ============================================================ */

/* Writes formulas, not values, so the tab recalculates on every sync
   with no need to run anything again. Safe to re-run at any time.
   Everything here depends on Trip Date holding real dates. If Today,
   This week and This month all read £0.00 while Lifetime is right,
   run Repair trip dates. */
function buildSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trips = ss.getSheetByName("Trips");
  if (!trips) { toast("No Trips sheet found."); return; }

  /* Every column is looked up by its heading at build time. Once written,
     Sheets keeps the formulas pointing at the same columns when Trips columns
     are moved or inserted, so the Summary follows them. */
  var cols = ensureHeaders_(trips, TRIP_HEADERS);
  var C = function (h) { var L = colLetter(colOf_(cols, h, "Trips")); return "Trips!" + L + "2:" + L; };
  var FARE = C("Fare (£)"), DATE = C("Trip Date"), ID = C("Driver ID"), NAME = C("Driver Name"),
      MILES = C("Distance (mi)");

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
    ["Estimated miles in background"], ["Gap events"],
    ["Metered fares, lifetime"], ["Permitted extras, lifetime"], ["Agreed fares, count"]
  ];
  sh.getRange(5, 1, labels.length, 1).setValues(labels);

  /* Fare (£) is what the passenger paid, so every headline here is real income.
     Metered, Extras and Agreed break that down and are blank on rows written before extras
     existed, which sums to zero rather than erroring. */
  sh.getRange(5, 2, 12, 1).setFormulas([
    ['=SUMIFS(' + FARE + ',' + DATE + ',">="&TODAY(),' + DATE + ',"<"&TODAY()+1)'],
    ['=SUMIFS(' + FARE + ',' + DATE + ',">="&' + monday + ',' + DATE + ',"<"&' + monday + '+7)'],
    ['=SUMIFS(' + FARE + ',' + DATE + ',">="&EOMONTH(TODAY(),-1)+1,' + DATE + ',"<"&EOMONTH(TODAY(),0)+1)'],
    ['=SUM(' + FARE + ')'],
    ['=COUNTA(' + ID + ')'],
    ['=IFERROR(AVERAGE(' + FARE + '),0)'],
    ['=SUM(' + MILES + ')'],
    ['=SUM(' + C("Estimated Miles") + ')'],
    ['=SUM(' + C("Gap Events") + ')'],
    ['=SUM(' + C("Metered Fare (£)") + ')'],
    ['=SUM(' + C("Extras (£)") + ')'],
    ['=COUNT(' + C("Agreed Fare (£)") + ')']
  ]);

  sh.getRange("B5:B8").setNumberFormat('£#,##0.00');
  sh.getRange("B10").setNumberFormat('£#,##0.00');
  sh.getRange("B11:B12").setNumberFormat('0.00');
  sh.getRange("B14:B15").setNumberFormat('£#,##0.00');

  sh.getRange("A18").setValue("Daily totals").setFontWeight("bold");
  /* The QUERY takes its columns as an array, {date, id, name, fare, miles},
     so it reads Col1..Col5 rather than sheet letters. */
  sh.getRange("A19").setFormula(
    '=IFERROR(QUERY({' + DATE + ',' + ID + ',' + NAME + ',' + FARE + ',' + MILES + '},' +
    '"select toDate(Col1), Col2, Col3, count(Col4), sum(Col4), sum(Col5) ' +
    'where Col2 is not null ' +
    'group by toDate(Col1), Col2, Col3 ' +
    'order by toDate(Col1) desc ' +
    "label toDate(Col1) 'Date', Col2 'Driver ID', Col3 'Driver', count(Col4) 'Trips', sum(Col4) 'Fare', sum(Col5) 'Miles' " +
    "format toDate(Col1) 'dd/MM/yyyy', sum(Col4) '£#,##0.00', sum(Col5) '0.00'\"" +
    ',0),"No trip data yet")'
  );

  sh.getRange("H18").setValue("Weekly totals, week beginning Monday").setFontWeight("bold");
  sh.getRange("H19").setFormula(
    '=IFERROR(QUERY({ARRAYFORMULA(IF(NOT(ISNUMBER(' + DATE + ')),"",INT(' + DATE + ')-WEEKDAY(' + DATE + ',3))),' +
    ID + ',' + FARE + ',' + MILES + '},' +
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

  relockIfNeeded();
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
  var map = headerMap_(sh), cId = map["Driver ID"], cName = map["Name"], cAct = activeCol_(map);
  if (!cId || !cName || !cAct) { toast("Drivers needs 'Driver ID', 'Name' and 'Active' headings. Run Add missing headers."); return; }

  var rows    = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var changed = 0;
  var named   = 0;

  for (var i = 0; i < rows.length; i++) {
    var id   = String(rows[i][cId - 1] || "").trim();
    var name = String(rows[i][cName - 1] || "").trim();
    if (!id) continue;
    if (name) { named++; continue; }
    if (String(rows[i][cAct - 1]).trim().toUpperCase() !== "FALSE") {
      /* One cell, in the one column that needs it, so no formula elsewhere on
         the row is flattened into a value. */
      sh.getRange(i + 2, cAct).setValue(false);
      changed++;
    }
  }

  toast(changed
    ? changed + " unnamed ID(s) set to FALSE. " + named + " named driver(s) left untouched."
    : "Nothing to change. " + named + " named driver(s), no unnamed active IDs.");
}

/* New IDs are written inactive. Add a name and set Active to TRUE to issue one. */
function generateBulkIds() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Drivers");
  if (!sh) { toast("No Drivers sheet found."); return; }

  /* lookupDriver reads row 1 as headings, so a tab without them would put an
     ID somewhere sign-in can never see it. */
  var map = headerMap_(sh), cId = map["Driver ID"], cAct = activeCol_(map);
  if (!cId || !cAct) {
    toast("The Drivers sheet has no 'Driver ID' and 'Active' headings. Run Taxi Meter → Add missing headers first, " +
          "otherwise new IDs could never sign in.");
    return;
  }

  var existing = Object.create(null);
  sh.getDataRange().getValues().forEach(function (r) {
    existing[String(r[cId - 1]).trim().toUpperCase()] = true;
  });

  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var added = 0, guard = 0, batch = [];
  var activeHeading = Object.keys(map).filter(function (h) { return map[h] === cAct; })[0];

  while (added < 10 && guard++ < 500) {
    var id = "TX-";
    for (var j = 0; j < 5; j++) id += chars.charAt(Math.floor(Math.random() * chars.length));
    if (existing[id]) continue;          /* never issue the same ID twice */
    existing[id] = true;
    var row = { "Driver ID": id };
    row[activeHeading] = "FALSE";
    batch.push(row);
    added++;
  }

  writeRowsByHeader_(sh, map, batch);
  toast(added + " new driver ID(s) added, all inactive. Add a name and set Active to TRUE to issue one.");
}

/* Fresh install only. Creates all three sheets with full headers. */
function setupHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureHeaders_(ss.getSheetByName("Drivers") || ss.insertSheet("Drivers"), DRIVER_HEADERS);
  ensureHeaders_(ss.getSheetByName("Trips")   || ss.insertSheet("Trips"),   TRIP_HEADERS);
  ensureHeaders_(ss.getSheetByName("Log")     || ss.insertSheet("Log"),     LOG_HEADERS);

  fixHeaders();
  lockSheet();
}

/* ============================================================
   Columns by heading, never by position

   From the minibus sheet (drolnstone/minibus-check: headerMap, colOf, and its
   rule "find a column by its heading, never by counting to it"). The office can
   move, insert or add columns on any tab and the app reads and writes straight
   past them. Headings are the contract: rename one and the app adds a fresh
   column with the right heading at the end rather than guessing.
   ============================================================ */

/* { heading: 1-based column }. If a heading appears twice the first wins,
   because that is the one the data was written under. */
function headerMap_(sh) {
  var map = {};
  if (!sh || sh.getLastColumn() < 1 || sh.getLastRow() < 1) return map;
  sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach(function (h, i) {
    var k = String(h == null ? "" : h).trim();
    if (k && map[k] === undefined) map[k] = i + 1;
  });
  return map;
}

/* Adds any wanted heading the tab lacks, at the end, and returns the map.
   Never relabels or moves an existing column. */
function ensureHeaders_(sh, wanted) {
  var map = headerMap_(sh);
  var next = 1;
  Object.keys(map).forEach(function (k) { next = Math.max(next, map[k] + 1); });
  wanted.forEach(function (h) {
    if (map[h]) return;
    if (next > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), next - sh.getMaxColumns());
    sh.getRange(1, next).setValue(h).setFontWeight("bold");
    map[h] = next++;
  });
  if (sh.getFrozenRows() < 1) sh.setFrozenRows(1);
  return map;
}

/* The column under a heading, or an error that names it. */
function colOf_(map, heading, tab) {
  var c = map[heading];
  if (!c) throw new Error("The " + tab + " tab has no '" + heading + "' column. " +
                          "Put the heading back exactly as it was, or run Taxi Meter → Add missing headers.");
  return c;
}

/* The Drivers tab's Active column has carried more than one wording. */
function activeCol_(map) {
  if (map[DRIVER_HEADERS[2]]) return map[DRIVER_HEADERS[2]];
  var k = Object.keys(map).filter(function (h) { return /^active\b/i.test(h); })[0];
  return k ? map[k] : 0;
}

/* Writes rows given as { heading: value } objects below the last row, one
   setValues per run of adjacent columns, so a column of the office's own (a
   note, a formula) is never written over. */
function writeRowsByHeader_(sh, map, objs) {
  if (!objs.length) return 0;
  var startRow = sh.getLastRow() + 1;
  var short = (startRow + objs.length - 1) - sh.getMaxRows();
  if (short > 0) sh.insertRowsAfter(sh.getMaxRows(), short + 500);

  var cols = [];
  Object.keys(objs[0]).forEach(function (h) { if (map[h]) cols.push({ h: h, c: map[h] }); });
  cols.sort(function (a, b) { return a.c - b.c; });

  var i = 0;
  while (i < cols.length) {
    var j = i;
    while (j + 1 < cols.length && cols[j + 1].c === cols[j].c + 1) j++;
    var run = cols.slice(i, j + 1);
    var values = objs.map(function (o) { return run.map(function (x) { return o[x.h]; }); });
    sh.getRange(startRow, run[0].c, objs.length, run.length).setValues(values);
    i = j + 1;
  }
  return startRow;
}

/* One row as { heading: value }, for readers. */
function rowObject_(map, row) {
  var o = {};
  Object.keys(map).forEach(function (h) { o[h] = row[map[h] - 1]; });
  return o;
}

function colLetter(n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}

/* ============================================================
   Protection

   From the minibus sheet's lock set (applyLocks, locksMissing, relock at
   3am). Warning-only, as there: the owner can still edit anything, but an
   accidental keystroke on a trip record or a heading asks first. Strict
   protection would not stop the owner at all, and would lock out anyone the
   sheet is later shared with. The script runs as the owner, so its own writes
   are never stopped.

     Trips, Log, Summary   the whole tab (the app writes them)
     Drivers               the heading row only (the office types the rest)
   ============================================================ */

var LOCK_TAG = "Taxi Meter lock";

function lockPlan_(ss) {
  var out = [];
  [["Trips", "trip records, written by the app"], ["Log", "written by the app"],
   ["Summary", "formulas, rebuilt from the menu"]].forEach(function (x) {
    var sh = ss.getSheetByName(x[0]);
    if (sh) out.push({ sh: sh, whole: true, note: x[1] });
  });
  var d = ss.getSheetByName("Drivers");
  if (d) out.push({ sh: d, whole: false, note: "heading row: the app finds columns by these names" });
  return out;
}

function applyLocks_(ss) {
  removeLocks_(ss);
  var skipped = [];
  lockPlan_(ss).forEach(function (item) {
    try {
      var p = item.whole
        ? item.sh.protect()
        : item.sh.getRange(1, 1, 1, Math.max(1, item.sh.getLastColumn())).protect();
      p.setDescription(LOCK_TAG + ": " + item.note);
      p.setWarningOnly(true);
    } catch (err) {
      skipped.push(item.sh.getName() + " (" + String((err && err.message) || err) + ")");
    }
  });
  /* Trips' heading row stays protected even if the whole-tab lock is removed
     by hand: renaming a heading is the one edit that sends data astray. */
  var t = ss.getSheetByName("Trips");
  if (t && t.getLastColumn() > 0) {
    try {
      t.getRange(1, 1, 1, t.getLastColumn()).protect()
       .setDescription(LOCK_TAG + ": heading row").setWarningOnly(true);
    } catch (err) {}
  }
  return skipped;
}

function removeLocks_(ss) {
  [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE].forEach(function (type) {
    ss.getProtections(type).forEach(function (p) {
      if (String(p.getDescription() || "").indexOf(LOCK_TAG) === 0) p.remove();
    });
  });
}

/* Tabs in the plan with no lock of ours on them. */
function locksMissing_(ss) {
  var have = {};
  [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE].forEach(function (type) {
    ss.getProtections(type).forEach(function (p) {
      if (String(p.getDescription() || "").indexOf(LOCK_TAG) !== 0) return;
      try { have[p.getRange().getSheet().getName()] = true; } catch (err) {}
    });
  });
  return lockPlan_(ss).filter(function (i) { return !have[i.sh.getName()]; })
                      .map(function (i) { return i.sh.getName(); });
}

/* The nightly timer, and after a tab is made. */
function relockIfNeeded() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try { if (locksMissing_(ss).length) applyLocks_(ss); } catch (err) { Logger.log("relock: " + err); }
}

/* One daily timer at about 3am, made once. */
function ensureRelockTrigger_() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === "relockIfNeeded"; });
  if (!has) ScriptApp.newTrigger("relockIfNeeded").timeBased().everyDays(1).atHour(3).create();
}

function lockSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var skipped = applyLocks_(ss);
  var timer = "";
  try { ensureRelockTrigger_(); timer = "It locks itself again every night at about 3am."; }
  catch (err) { timer = "The nightly relock could not be set up: " + (err && err.message ? err.message : err); }
  toast("✓ The sheet is protected.\n\n" +
        "Editing Trips, Log or Summary, or a heading on Drivers, now asks you to confirm first. " +
        "You can still go ahead when you mean to: this stops the accidental keystroke, not you.\n\n" +
        "Drivers below the heading row is edited freely.\n\n" + timer +
        (skipped.length ? "\n\nCould not lock: " + skipped.join(", ") : ""));
}

function unlockSheet() {
  removeLocks_(SpreadsheetApp.getActiveSpreadsheet());
  toast("Unlocked. Nothing will warn you now. Run Taxi Meter → Protect the sheet when you are done; " +
        "it also locks itself again overnight.");
}
