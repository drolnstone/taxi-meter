/*  PDF receipts and earnings statements.

    The layout engine is the minibus coordinator app's (drolnstone/minibus-check,
    coord/pdf.js, from its v1.95.0): the app's banner across the top, a title and
    period, each section as a table that splits cleanly across pages, and
    "Page x of y" on every page. The words are real text, so they can be
    selected and searched. vendor/jspdf.umd.min.js (MIT, the same file the
    minibus app ships) does the writing; it is loaded only when a PDF is made,
    and nothing here reaches the network.

    Exposes window.TaxiPdf:
      receiptPdf(trip, opt)        one trip, A5 portrait
      statementPdf(stmt, opt)      earnings for a period, A4 portrait
      checkPdf(check, sections, renewals, opt)   one vehicle check, A4
      periodRange(kind, today)     this-week / last-week / this-month /
                                   last-month / tax-year
      buildStatement(trips, range, parseDate)
      fileName(kind, ...)          */

(function (root) {
  "use strict";

  /* The PDF's built-in fonts know Latin-1 and no more. Anything else is put
     into its nearest plain form, so no word silently loses a letter. £ is in
     Latin-1 and survives. */
  var SWAP = [[/[‘’‛′]/g, "'"], [/[“”″]/g, '"'],
              [/[–—−]/g, "-"], [/[·•]/g, "-"], [/…/g, "..."], [/→/g, "to"],
              [/[    ]/g, " "]];
  function clean(s) {
    s = String(s == null ? "" : s);
    SWAP.forEach(function (x) { s = s.replace(x[0], x[1]); });
    return s.replace(/[^\x09\x0A\x0D\x20-\x7E\xA0-\xFF]/g, "?");
  }

  function money(n) { return "£" + (Number(n) || 0).toFixed(2); }
  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function ukDate(d) { return pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear(); }
  function isoDate(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function hasAgreed(t) { return t.agreedFare !== null && t.agreedFare !== undefined && t.agreedFare !== ""; }
  function extrasOf(t) { return Array.isArray(t.extras) ? t.extras : []; }
  function extrasSum(t) { return extrasOf(t).reduce(function (a, x) { return a + (Number(x.amount) || 0); }, 0); }

  var M = 14, FOOT = 10, PAD = 1.4, SIZE = 8.5;
  var LINE = SIZE * 0.3528 * 1.25;

  /* Column widths that fit the page: each column gets what its widest line
     needs if all of them fit, otherwise an even share first and what is left
     to the columns that want more. As in the minibus app. */
  function widths(doc, cols, rows, avail) {
    var want = cols.map(function (c, i) {
      doc.setFont("helvetica", "bold");
      var w = doc.getTextWidth(clean(c));
      doc.setFont("helvetica", "normal");
      rows.forEach(function (r) {
        String(clean(r[i])).split("\n").forEach(function (l) { w = Math.max(w, doc.getTextWidth(l)); });
      });
      return w + 2 * PAD + 0.5;
    });
    var total = want.reduce(function (a, b) { return a + b; }, 0);
    if (total <= avail) return want.map(function (w) { return w * avail / total; });
    var share = avail / cols.length;
    var got = want.map(function (w) { return Math.min(w, share); });
    var left = avail - got.reduce(function (a, b) { return a + b; }, 0);
    var need = want.map(function (w, i) { return w - got[i]; });
    var needAll = need.reduce(function (a, b) { return a + b; }, 0);
    return got.map(function (g, i) { return g + (needAll ? left * need[i] / needAll : 0); });
  }

  /* A document with the banner, title and footer, and a way to add tables. */
  function frame(format, title, what, period, opt) {
    var J = root.jspdf && root.jspdf.jsPDF;
    if (!J) throw new Error("The PDF writer has not loaded.");
    var doc = new J({ orientation: "portrait", unit: "mm", format: format, compress: opt.compress !== false });
    var W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight();
    var avail = W - 2 * M;
    doc.setProperties({ title: clean(title + (period ? " - " + period : "")),
                        author: clean(opt.driver || ""), creator: "Taxi Meter" });

    var bh = 20;
    doc.setFillColor(11, 11, 11);
    doc.rect(0, 0, W, bh, "F");
    doc.setFillColor(255, 214, 10);
    doc.rect(0, bh, W, 1.2, "F");
    doc.setTextColor(255);
    doc.setFont("helvetica", "bold"); doc.setFontSize(14);
    doc.text("Taxi Meter", M, bh / 2 - 0.5);
    doc.setFont("helvetica", "normal"); doc.setFontSize(9); doc.setTextColor(200);
    doc.text(clean(opt.place || "Hackney carriage - City of Liverpool"), M, bh / 2 + 4.5);
    doc.setTextColor(0);

    var y = bh + 10;
    doc.setFont("helvetica", "bold"); doc.setFontSize(16);
    doc.text(clean(title), M, y);
    y += 6;
    if (what) {
      doc.setFont("helvetica", "normal"); doc.setFontSize(9.5); doc.setTextColor(60);
      var wl = doc.splitTextToSize(clean(what), avail);
      doc.text(wl, M, y);
      y += wl.length * 4.4;
      doc.setTextColor(0);
    }
    if (period) {
      doc.setFont("helvetica", "bold"); doc.setFontSize(10);
      doc.text(clean(period), M, y + 0.5);
      y += 4.8;
    }
    if (opt.made) {
      doc.setFont("helvetica", "normal"); doc.setFontSize(8); doc.setTextColor(90);
      doc.text(clean("Made " + opt.made + (opt.version ? " - Taxi Meter " + opt.version : "")), M, y);
      doc.setTextColor(0);
      y += 3;
    }
    doc.setDrawColor(150); doc.setLineWidth(0.4); doc.line(M, y, W - M, y);
    y += 6;

    var f = { doc: doc, W: W, H: H, avail: avail, bottom: H - M - FOOT };
    f.y = function () { return y; };
    f.move = function (dy) { y += dy; };
    f.newPage = function () { doc.addPage(); y = M; };

    function rowHeight(cells, ws) {
      var most = 1;
      cells.forEach(function (c, i) { most = Math.max(most, doc.splitTextToSize(clean(c), ws[i] - 2 * PAD).length); });
      return most * LINE + 2 * PAD;
    }
    function drawRow(cells, ws, bold, fill, right) {
      doc.setFont("helvetica", bold ? "bold" : "normal");
      var h = rowHeight(cells, ws);
      if (fill != null) { doc.setFillColor(fill); doc.rect(M, y, avail, h, "F"); }
      var x = M;
      cells.forEach(function (c, i) {
        var lines = doc.splitTextToSize(clean(c), ws[i] - 2 * PAD);
        if (right && right[i]) doc.text(lines, x + ws[i] - PAD, y + PAD + LINE * 0.78, { lineHeightFactor: 1.25, align: "right" });
        else doc.text(lines, x + PAD, y + PAD + LINE * 0.78, { lineHeightFactor: 1.25 });
        x += ws[i];
      });
      doc.setDrawColor(210); doc.setLineWidth(0.15); doc.line(M, y + h, W - M, y + h);
      y += h;
    }

    /* sec: { head, cols[], rows[][], right[] (true = right align), empty } */
    f.table = function (sec) {
      var cols = (sec.cols || []).map(clean);
      var rows = (sec.rows || []).map(function (r) { return r.map(clean); });
      doc.setFontSize(SIZE);
      var ws = widths(doc, cols, rows, avail);
      var headed = cols.some(function (c) { return c; });   /* label/value pairs have no header row */
      var first = rows.length ? rowHeight(rows[0], ws) : LINE;
      if (y + 8 + (headed ? rowHeight(cols, ws) : 0) + first > f.bottom) f.newPage();
      if (sec.head) {
        doc.setFont("helvetica", "bold"); doc.setFontSize(11.5);
        doc.text(clean(sec.head), M, y + 4);
        y += 7;
      }
      doc.setFontSize(SIZE);
      if (!rows.length) {
        doc.setFont("helvetica", "normal"); doc.setTextColor(90);
        doc.text(clean(sec.empty || "None."), M, y + 3.5);
        doc.setTextColor(0);
        y += 9;
        return;
      }
      if (headed) drawRow(cols, ws, true, 232, sec.right);
      rows.forEach(function (r, i) {
        if (y + rowHeight(r, ws) > f.bottom) { f.newPage(); doc.setFontSize(SIZE); if (headed) drawRow(cols, ws, true, 232, sec.right); }
        drawRow(r, ws, !!(sec.boldLast && i === rows.length - 1), i % 2 ? 247 : null, sec.right);
      });
      y += 6;
    };

    /* Two columns of label and value, no header row. */
    f.pairs = function (head, list) {
      f.table({ head: head, cols: ["", ""], rows: list, right: [false, true], boldLast: false });
    };

    f.note = function (text) {
      doc.setFont("helvetica", "normal"); doc.setFontSize(8.5); doc.setTextColor(70);
      var l = doc.splitTextToSize(clean(text), avail);
      if (y + l.length * 3.8 > f.bottom) f.newPage();
      doc.text(l, M, y);
      y += l.length * 3.8 + 2;
      doc.setTextColor(0);
    };

    f.finish = function (footTitle) {
      var n = doc.getNumberOfPages();
      for (var p = 1; p <= n; p++) {
        doc.setPage(p);
        doc.setFont("helvetica", "normal"); doc.setFontSize(7.5); doc.setTextColor(90);
        doc.setDrawColor(190); doc.setLineWidth(0.2); doc.line(M, H - M - 4, W - M, H - M - 4);
        doc.text(clean(footTitle), M, H - M);
        doc.text("Page " + p + " of " + n, W - M, H - M, { align: "right" });
        doc.setTextColor(0);
      }
      return doc;
    };
    return f;
  }

  /* ---- One trip ---- */

  function receiptPdf(t, opt) {
    opt = opt || {};
    var f = frame("a5", "Taxi receipt", "", t.date || "", opt);
    var doc = f.doc;

    doc.setFont("helvetica", "normal"); doc.setFontSize(10); doc.setTextColor(70);
    doc.text("Total to pay", M, f.y());
    doc.setFont("helvetica", "bold"); doc.setFontSize(26); doc.setTextColor(0);
    doc.text(clean(money(t.fare)), M, f.y() + 10);
    f.move(17);

    var trip = [
      ["Driver", (t.driverName || "") + (t.driverId ? " (" + t.driverId + ")" : "")],
      ["Date", t.date || ""],
      ["Distance", (t.miles || "0.00") + " mi"],
      ["Time", t.time || ""],
      ["Tariff", (t.tariff || "Day") + " rate"]
    ];
    if (t.pickupLabel) trip.push(["From", t.pickupLabel]);
    if (t.destination) trip.push(["To", t.destination]);
    if (opt.vehicle) trip.push(["Vehicle", opt.vehicle]);
    f.pairs("Journey", trip);

    var fare = [["Metered fare", money(t.meteredFare)]];
    if (hasAgreed(t)) fare.push(["Agreed fare (journey ending 4+ miles outside Liverpool)", money(t.agreedFare)]);
    extrasOf(t).forEach(function (x) { fare.push([x.label, money(x.amount)]); });
    fare.push(["Total to pay", money(t.fare)]);
    f.table({ head: "Fare", cols: ["", ""], rows: fare, right: [false, true], boldLast: true });

    f.note("Fares are set by the City of Liverpool Hackney Carriage Rates of Fare. " +
           "Keep this receipt if you need to query the fare." +
           (t.tripId ? " Reference: " + String(t.tripId).slice(0, 8).toUpperCase() + "." : ""));
    return f.finish("Taxi receipt - " + (t.date || ""));
  }

  /* ---- Periods ---- */

  function startOfDay(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
  function monday(d) { var s = startOfDay(d), wd = (s.getDay() + 6) % 7; return addDays(s, -wd); }

  /* { kind, label, from, to } with to exclusive. The UK tax year runs from
     6 April. */
  function periodRange(kind, today) {
    today = startOfDay(today || new Date());
    var from, to, label;
    if (kind === "last-week") {
      to = monday(today); from = addDays(to, -7); label = "Last week";
    } else if (kind === "this-month") {
      from = new Date(today.getFullYear(), today.getMonth(), 1); to = addDays(today, 1); label = "This month";
    } else if (kind === "last-month") {
      to = new Date(today.getFullYear(), today.getMonth(), 1);
      from = new Date(to.getFullYear(), to.getMonth() - 1, 1); label = "Last month";
    } else if (kind === "tax-year") {
      var y = today.getFullYear();
      from = new Date(y, 3, 6);
      if (today < from) from = new Date(y - 1, 3, 6);
      to = addDays(today, 1);
      label = "Tax year " + from.getFullYear() + "-" + String(from.getFullYear() + 1).slice(2) + " to date";
    } else {
      kind = "this-week"; from = monday(today); to = addDays(today, 1); label = "This week";
    }
    return { kind: kind, label: label, from: from, to: to,
             text: ukDate(from) + " to " + ukDate(addDays(to, -1)) };
  }

  /* Everything a statement prints, worked out once so the screen and the PDF
     agree. parseDate turns a trip's "dd/mm/yyyy, hh:mm:ss" into a Date. */
  function buildStatement(trips, range, parseDate) {
    var rows = [], days = {}, tariffs = {};
    var total = 0, metered = 0, extras = 0, miles = 0, agreed = 0;
    (trips || []).forEach(function (t) {
      var d = parseDate(t.date);
      if (!d || d < range.from || d >= range.to) return;
      var fare = Number(t.fare) || 0, ex = extrasSum(t), mi = Number(t.miles) || 0;
      total += fare; extras += ex; miles += mi;
      metered += Number(t.meteredFare !== undefined ? t.meteredFare : t.fare) || 0;
      if (hasAgreed(t)) agreed++;
      var key = isoDate(d);
      if (!days[key]) days[key] = { date: ukDate(d), trips: 0, fare: 0, miles: 0 };
      days[key].trips++; days[key].fare += fare; days[key].miles += mi;
      var band = t.tariff || "Day";
      if (!tariffs[band]) tariffs[band] = { trips: 0, fare: 0 };
      tariffs[band].trips++; tariffs[band].fare += fare;
      rows.push({ when: d, trip: t });
    });
    rows.sort(function (a, b) { return a.when - b.when; });
    var r2 = function (n) { return Math.round(n * 100) / 100; };
    return {
      range: range, count: rows.length,
      total: r2(total), metered: r2(metered), extras: r2(extras), miles: r2(miles), agreedCount: agreed,
      average: rows.length ? r2(total / rows.length) : 0,
      days: Object.keys(days).sort().map(function (k) { var x = days[k]; return { date: x.date, trips: x.trips, fare: r2(x.fare), miles: r2(x.miles) }; }),
      tariffs: Object.keys(tariffs).sort().map(function (k) { return { band: k, trips: tariffs[k].trips, fare: r2(tariffs[k].fare) }; }),
      trips: rows.map(function (r) { return r.trip; })
    };
  }

  function statementPdf(s, opt) {
    opt = opt || {};
    var f = frame("a4", "Earnings statement",
      "Fares taken, as recorded by this phone's meter. Trips removed from the phone after 90 days are not included.",
      s.range.label + ": " + s.range.text, opt);

    var head = [
      ["Driver", (opt.driver || "") + (opt.driverId ? " (" + opt.driverId + ")" : "")],
      ["Trips", String(s.count)],
      ["Total taken", money(s.total)],
      ["of which metered fares", money(s.metered)],
      ["of which permitted extras (tolls, soiling)", money(s.extras)],
      ["Trips on an agreed fare", String(s.agreedCount)],
      ["Average fare", money(s.average)],
      ["Distance with passengers", s.miles.toFixed(2) + " mi"]
    ];
    if (opt.vehicle) head.splice(1, 0, ["Vehicle", opt.vehicle]);
    f.pairs("Summary", head);

    f.table({ head: "By day", cols: ["Date", "Trips", "Miles", "Taken"], right: [false, true, true, true],
      rows: s.days.map(function (d) { return [d.date, String(d.trips), d.miles.toFixed(2), money(d.fare)]; }),
      empty: "No trips in this period." });

    f.table({ head: "By tariff", cols: ["Tariff", "Trips", "Taken"], right: [false, true, true],
      rows: s.tariffs.map(function (b) { return [b.band, String(b.trips), money(b.fare)]; }),
      empty: "No trips in this period." });

    f.table({ head: "Every trip", cols: ["Date", "Miles", "Time", "Tariff", "Extras", "Taken"],
      right: [false, true, true, false, true, true],
      rows: s.trips.map(function (t) {
        var ex = extrasSum(t);
        return [t.date || "", String(t.miles || "0.00"), t.time || "", (t.tariff || "Day") + (hasAgreed(t) ? " (agreed)" : ""),
                ex ? money(ex) : "", money(t.fare)];
      }),
      empty: "No trips in this period." });

    return f.finish("Earnings statement - " + s.range.label + ": " + s.range.text);
  }

  /* ---- One vehicle check, for the driver's own records ----
     c: the saved check. sections: TaxiVehicle.CHECK_SECTIONS. renewals: the
     output of TaxiVehicle.renewalStatus, or nothing. */
  var MARK_WORD = { fine: "Fine", advisory: "Advisory", defect: "Defect" };

  function checkPdf(c, sections, renewals, opt) {
    opt = opt || {};
    var f = frame("a4", "Vehicle check", "The driver's own daily check, saved on their phone.", c.date || "", opt);
    var head = [
      ["Driver", (opt.driver || "") + (opt.driverId ? " (" + opt.driverId + ")" : "")],
      ["Vehicle", c.vehicle || ""],
      ["Odometer", c.odometer ? Number(c.odometer).toLocaleString("en-GB") + " miles" : ""],
      ["Outcome", c.outcome || ""]
    ];
    f.pairs("Summary", head.filter(function (r) { return r[1]; }));
    var found = [];
    if (c.defects) found.push(["Defects", c.defects]);
    if (c.advisories) found.push(["Advisories", c.advisories]);
    if (found.length) f.table({ head: "Findings", cols: ["", ""], rows: found, right: [false, false] });
    var items = c.items || null;
    (sections || []).forEach(function (sec) {
      f.table({ head: sec.title, cols: ["Item", "Result", "Note"], right: [false, false, false],
        rows: sec.items.map(function (it) {
          var r = items && items[it.id] || {};
          return [it.name + (it.crit ? " (critical)" : ""), MARK_WORD[r.mark] || (items ? "Not marked" : "-"), r.note || ""];
        }) });
    });
    if (renewals && renewals.length) {
      f.table({ head: "Renewal dates", cols: ["Document", "Date", ""], right: [false, false, false],
        rows: renewals.map(function (r) { return [r.name, r.date, r.days < 0 ? "Expired" : r.days + " days"]; }) });
    }
    if (!items) f.note("This check was saved before item-by-item results were kept, so only its findings are shown.");
    return f.finish("Vehicle check - " + (c.vehicle ? c.vehicle + " - " : "") + (c.date || ""));
  }

  function fileName(kind, a) {
    if (kind === "check") {
      var cd = String(a && a.date || "").split(", ")[0].split("/").reverse().join("-");
      return "Vehicle check " + (cd || "") + (a && a.vehicle ? " " + String(a.vehicle).replace(/[^\w -]/g, "") : "") + ".pdf";
    }
    if (kind === "receipt") {
      var d = String(a && a.date || "").split(", ")[0].split("/").reverse().join("-");
      return "Taxi receipt " + (d || "trip") + ".pdf";
    }
    return "Earnings statement " + isoDate(a.from) + " to " + isoDate(addDays(a.to, -1)) + ".pdf";
  }

  root.TaxiPdf = { receiptPdf: receiptPdf, statementPdf: statementPdf, checkPdf: checkPdf, periodRange: periodRange,
                   buildStatement: buildStatement, fileName: fileName, clean: clean };
})(typeof window !== "undefined" ? window : globalThis);
