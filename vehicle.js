/*  Daily vehicle check and renewal dates.

    Ported from the minibus app (drolnstone/minibus-check, index.html: the
    walkaround with Fine / Advisory / Defect on every item, critical items that
    stop the vehicle, and the 60 / 30 / 7 / 0 day renewal warnings of its MOT
    and insurance rules). Rewritten for a hackney carriage: one vehicle, one
    driver, no coordinator, so nothing here stops the meter outright. A
    critical defect or an expired document asks the driver before a trip
    starts, and the answer goes on the office log.

    Pure functions only. index.html draws the screens; tests/run-tests.mjs
    loads this file in Node.  */

(function (root) {
  "use strict";

  /* crit: a Defect here means the vehicle is not fit to carry passengers. */
  var CHECK_SECTIONS = [
    { title: "Outside", items: [
      { id: "tyres",   crit: true, name: "Tyres", what: "Tread above the legal limit across the whole width. No cuts, bulges or cords showing. Pressures look even." },
      { id: "nuts",    crit: true, name: "Wheel nuts", what: "All present and tight. No rust weeping from a nut." },
      { id: "lights",  crit: true, name: "Lights and indicators", what: "Sidelights, dipped, main beam, indicators, hazards, brake and reverse lights, number plate lamp." },
      { id: "plates",  crit: true, name: "Licence plates and number plates", what: "Council licence plate fixed and readable. Front and rear number plates clean and secure." },
      { id: "roofsign",            name: "Roof sign", what: "TAXI sign fitted, lit when for hire, off when hired." },
      { id: "body",                name: "Body and glass", what: "No sharp or loose panels. Windscreen chips or cracks outside the wiper area only." },
      { id: "mirrors", crit: true, name: "Mirrors, wipers and washers", what: "Mirrors secure and clean. Blades not torn. Washers reach the screen." },
      { id: "leaks",               name: "Leaks underneath", what: "No fresh oil, coolant or fuel on the ground where it was parked." }
    ]},
    { title: "Under the bonnet", items: [
      { id: "oil",                 name: "Engine oil", what: "Between min and max on the dipstick." },
      { id: "coolant",             name: "Coolant", what: "Between min and max, engine cold." },
      { id: "bfluid",  crit: true, name: "Brake fluid", what: "At or near max. Report a drop, do not just top it up." },
      { id: "wash",                name: "Screenwash", what: "Topped up. Winter mix from October." }
    ]},
    { title: "Passenger area", items: [
      { id: "doors",   crit: true, name: "Doors", what: "Every door opens, closes and latches from outside and in. Child locks work." },
      { id: "belts",   crit: true, name: "Seat belts", what: "Every seat, tip-up seats too: pull fully out, no fraying, locks and releases." },
      { id: "seats",               name: "Seats", what: "Secure, not torn through, tip-up seats fold and lock." },
      { id: "ramp",                name: "Wheelchair ramp and restraints", what: "Ramp deploys and stows. Restraints and belt present and not frayed." },
      { id: "partition",           name: "Partition, intercom and hearing loop", what: "Partition secure. Intercom and hearing loop working." },
      { id: "intlights",           name: "Interior lights", what: "Passenger and step lights working." },
      { id: "clean",               name: "Clean and clear", what: "Floor, seats and footwells clean. Nothing loose that could slide or trip." },
      { id: "extinguisher",        name: "Fire extinguisher", what: "Present, in date, gauge in the green, bracket tight." },
      { id: "firstaid",            name: "First aid kit", what: "Present and stocked." }
    ]},
    { title: "In the cab", items: [
      { id: "warning", crit: true, name: "Warning lights", what: "All lamps come on at ignition and go out. A brake or ABS lamp that stays on is a defect." },
      { id: "brakes",  crit: true, name: "Brakes", what: "Firm pedal, handbrake holds. No pulling to one side." },
      { id: "steering",crit: true, name: "Steering", what: "No excess free play or knocking." },
      { id: "horn",                name: "Horn, heater and demist", what: "All working. Demist clears the whole screen." },
      { id: "meter",   crit: true, name: "Taximeter and tariff card", what: "Meter seal intact. Current tariff card on display where passengers can read it." },
      { id: "badge",               name: "Badge on display", what: "Your driver's badge visible to passengers." },
      { id: "fuel",                name: "Fuel or charge", what: "Enough for the shift. Filler cap present and tight." }
    ]}
  ];

  var ALL_ITEMS = [];
  CHECK_SECTIONS.forEach(function (s) { s.items.forEach(function (it) { ALL_ITEMS.push(it); }); });

  var MARKS = ["fine", "advisory", "defect"];

  function itemById(id) {
    for (var i = 0; i < ALL_ITEMS.length; i++) if (ALL_ITEMS[i].id === id) return ALL_ITEMS[i];
    return null;
  }

  /* What a finished check amounts to.
       results: { itemId: { mark: "fine"|"advisory"|"defect", note: "" } }
     Returns { complete, missing[], needNote[], outcome, fit, defects[], advisories[] }
     outcome is one of "Fit", "Fit with advisories", "Defects found", "Unfit". */
  function checkOutcome(results) {
    results = results || {};
    var missing = [], needNote = [], defects = [], advisories = [], critical = false;
    ALL_ITEMS.forEach(function (it) {
      var r = results[it.id];
      if (!r || MARKS.indexOf(r.mark) < 0) { missing.push(it.id); return; }
      var note = String(r.note || "").trim();
      if (r.mark === "defect") {
        if (!note) needNote.push(it.id);
        defects.push({ id: it.id, name: it.name, crit: !!it.crit, note: note });
        if (it.crit) critical = true;
      } else if (r.mark === "advisory") {
        advisories.push({ id: it.id, name: it.name, note: note });
      }
    });
    var outcome = critical ? "Unfit"
                : defects.length ? "Defects found"
                : advisories.length ? "Fit with advisories" : "Fit";
    return {
      complete: !missing.length && !needNote.length,
      missing: missing, needNote: needNote,
      outcome: outcome, fit: !critical,
      defects: defects, advisories: advisories
    };
  }

  /* One line per finding, for the sheet and the office log. */
  function describeFindings(list) {
    return (list || []).map(function (f) {
      return f.name + (f.crit ? " (critical)" : "") + (f.note ? ": " + f.note : "");
    }).join("; ");
  }

  /* ---- Renewal dates ---- */

  /* gate: when expired, the driver is asked before a trip starts. */
  var DOCUMENTS = [
    { id: "vehicleLicence", name: "Vehicle licence (plate)",  gate: true },
    { id: "badge",          name: "Driver's badge",           gate: true },
    { id: "insurance",      name: "Insurance",                gate: true },
    { id: "mot",            name: "MOT / compliance test",    gate: true },
    { id: "dvla",           name: "DVLA driving licence",     gate: false },
    { id: "meterTest",      name: "Taximeter test",           gate: false }
  ];

  /* The minibus app's warning steps: 60, 30 and 7 days out, then on the day. */
  var WARN_DAYS = [60, 30, 7, 0];

  function dayNumber(d) { return Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 86400000); }

  /* "2027-03-14" (what <input type=date> gives) to a local Date, or null. */
  function parseIsoDate(s) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ""));
    if (!m) return null;
    var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (d.getFullYear() !== Number(m[1]) || d.getMonth() !== Number(m[2]) - 1 || d.getDate() !== Number(m[3])) return null;
    return d;
  }

  function ukDate(d) {
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return p(d.getDate()) + "/" + p(d.getMonth() + 1) + "/" + d.getFullYear();
  }

  /* Every document with a date, soonest first.
       level: "expired" (past the date), "today", "soon" (7 days or less),
              "due" (30 or less), "upcoming" (60 or less), "ok" */
  function renewalStatus(dates, today) {
    dates = dates || {};
    today = today || new Date();
    var t = dayNumber(today), out = [];
    DOCUMENTS.forEach(function (doc) {
      var d = parseIsoDate(dates[doc.id]);
      if (!d) return;
      var days = dayNumber(d) - t;
      var level = days < 0 ? "expired" : days === 0 ? "today" : days <= 7 ? "soon"
                : days <= 30 ? "due" : days <= 60 ? "upcoming" : "ok";
      var when = days < 0 ? "expired on " + ukDate(d)
               : days === 0 ? "expires today"
               : days === 1 ? "expires tomorrow"
               : "expires in " + days + " days, on " + ukDate(d);
      out.push({ id: doc.id, name: doc.name, gate: doc.gate, date: ukDate(d), days: days, level: level,
                 text: doc.name + " " + when });
    });
    out.sort(function (a, b) { return a.days - b.days; });
    return out;
  }

  /* The ones worth putting on the meter screen: 60 days out or nearer. */
  function renewalWarnings(dates, today) {
    return renewalStatus(dates, today).filter(function (r) { return r.level !== "ok"; });
  }

  /* Expired documents that stop a trip until the driver says otherwise. */
  function blockingRenewals(dates, today) {
    return renewalStatus(dates, today).filter(function (r) { return r.gate && r.level === "expired"; });
  }

  /* "MOT 14/03/2027; Insurance 01/01/2027", for the sheet. */
  function renewalSummary(dates) {
    return DOCUMENTS.map(function (doc) {
      var d = parseIsoDate((dates || {})[doc.id]);
      return d ? doc.name + " " + ukDate(d) : "";
    }).filter(Boolean).join("; ");
  }

  root.TaxiVehicle = {
    CHECK_SECTIONS: CHECK_SECTIONS, ALL_ITEMS: ALL_ITEMS, MARKS: MARKS, itemById: itemById,
    checkOutcome: checkOutcome, describeFindings: describeFindings,
    DOCUMENTS: DOCUMENTS, WARN_DAYS: WARN_DAYS, parseIsoDate: parseIsoDate, ukDate: ukDate,
    renewalStatus: renewalStatus, renewalWarnings: renewalWarnings,
    blockingRenewals: blockingRenewals, renewalSummary: renewalSummary
  };
})(typeof window !== "undefined" ? window : globalThis);
