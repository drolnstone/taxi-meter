# Upgrade pass — what changed, and what to do on deploy day

Five files. `index.html` was rewritten; `Code.gs` took 26 patches; `sw.js` and
`manifest.json` are new; `tariff-bench.html` is unchanged from the last round.

I ran two independent review passes over the previous build — one on the UI and
markup, one on the Apps Script — and both found blockers I had missed. Those, plus
the outstanding items you asked me to close, are all in this build.

**Verification: 106 automated checks, all passing.** The fare engine and tariff
calendar were spliced into the rewritten `index.html` verbatim rather than retyped,
so the 84 checks that already covered them still apply unchanged, and the tests
extract the real declarations out of the shipped file rather than a copy.

---

## 1. Two ways the app could have failed completely

**A missing map library took the whole meter down.** `L.map()` ran at the top level
of the script. If unpkg was unreachable — captive portal, tunnel, blocked CDN, slow
DNS — that line threw, *every statement after it never ran*, and since no page
carried the `active` class in the markup, the router never got to add one. The
driver saw a blank black screen. No sign-in, no Start button, no error, nothing.
A mapping library the meter does not need in order to bill a fare could take the
meter offline.

Now: the sign-in page is visible in the markup before a line of JavaScript runs, the
map is constructed inside a guard, every `map.` call site is null-checked, and the
map panel shows "Map unavailable — the meter is working normally" when it can't load.

**A hung request stranded the driver mid-job.** No `fetch` had a timeout. On a
marginal signal the promise may never settle — and `startTrip` has already disabled
its own button by then, so the driver is left staring at a dead "Checking…" with a
passenger in the car and no way out but reloading the page, which they won't think
to do. Now every call aborts at 12 seconds (30 for a sync batch) and says which of
"no connection" or "not answering" happened.

## 2. The sync outage that was already scheduled

`getRange().setValues()` does not grow a sheet the way `appendRow()` does. A default
Google Sheet has 1000 rows. Once `getLastRow()` reached that, **every sync from every
driver would have failed permanently** with an out-of-bounds error, and trips would
have piled up in localStorage until a phone was wiped. At five drivers doing ten
trips a day that is about four months out. Fixed: the grid is extended before the
write, 500 rows at a time.

## 3. Silent fare loss, from two directions

**Text dates poisoned the Summary tab.** An unparseable date was written to column D
as raw text. `SUMIFS` ignores it, so Today/Week/Month quietly undercount — but worse,
the weekly `QUERY` calls `INT()` on it, one bad cell errors the whole array, and the
`IFERROR` wrapper replaces the entire weekly table with **"No trip data yet"**. That
reads as a quiet week, not a broken report. Now such a trip is refused at validation,
where the driver sees the message and the trip stays safely queued.

`parseUkDate_` was also fabricating dates rather than refusing them — `31/02/2026`
became 3 March, `24:30` became half past midnight the *next day*. That last one is
not hypothetical: some older WebKit builds emit hour 24 from `toLocaleString`, which
would have moved every midnight-hour trip — night tariff, the expensive ones — onto
the wrong day. Both refused now, and the phone builds its timestamp by hand instead
of trusting `toLocaleString`. There is a test asserting the phone's output parses
with the backend's parser, so the two can't drift apart.

**Trips without an ID were marked synced regardless.** The trip ID is the only thing
matching an acknowledgement to a trip. It was validated only when present, and the
phone had a line marking any ID-less trip synced no matter what the backend did — so
a trip the backend *rejected* was deleted from the only other place it existed. The
ID is now required, and that line is gone.

## 4. Permitted additional charges — the gap I said I would not go live without

The card allows tolls (tunnel, motorway, bridge, airport), £40 for soiling inside the
vehicle and £10 outside. There was no way to add any of them. In Liverpool that is the
Mersey Tunnel and the airport run — the meter, the receipt and the office sheet were
all wrong by the toll, every time.

Now: an Extras card on the meter screen with one-tap buttons for the two fixed soiling
charges (they are fixed by the card, so they are not typed) and a small form for tolls,
which vary. Each extra is itemised on the receipt, in the summary, in the CSV and on the
sheet. The fare display keeps showing the meter reading as the headline and adds a line
underneath: *Meter £7.40 + £2.30 extras — charge £9.70*.

**Agreed fares** are in too. The card allows a fare agreed before the journey for trips
ending four or more miles beyond the boundary. Enter it on the job setup screen; the
meter still runs and its reading is still recorded, but the charge is the agreed amount
plus any permitted extras, and the receipt says so plainly.

### The sheet schema changed — read this bit

`Trips` goes from 17 columns to 23. **Column G keeps its meaning**: it is what the
passenger actually paid. That was already true for every existing row, so every
Summary formula you have stays correct and nothing needs migrating. The six new
columns break the total down:

| | |
|---|---|
| **R** Metered Fare (£) | the meter reading on its own |
| **S** Extras (£) | total of permitted additional charges |
| **T** Extras Detail | itemised, e.g. `Tunnel toll £2.30; Soiling inside £40.00` |
| **U** Agreed Fare (£) | set only where a fare was agreed |
| **V** Running Miles | distance actually billed as running |
| **W** Waiting Time | time actually billed as waiting |

**The migration is automatic.** The first sync after you deploy checks the layout, adds
the missing columns and writes their headers. You do not need to touch the sheet. If
columns A–Q have been *reordered* at some point, it refuses to write rather than
misaligning rows, and tells you which column is wrong.

Rebuild the Summary tab afterwards to pick up three new headline rows (metered fares,
extras, agreed-fare count).

## 5. The UI pass

The trip screen had a real problem: the fare — the one number the product exists to
show — sat *after* the navigation and idle banners in the source. With the idle prompt
up on a small phone, it was pushed below the fold at exactly the moment the app was
asking the driver to make a billing decision. The fare is now sticky at the top of the
trip screen and nothing can displace it.

**GPS loss was 15px of dim grey text** in the second card, below the fold. If
`watchPosition` errors mid-trip the meter stops accruing distance and bills waiting
time only, and the driver had no way to notice. It is now a red banner with a spoken
alert.

**`confirm()` and `alert()` are gone.** Ending a trip went through the OS dialog:
small buttons outside a dash-mounted phone's thumb reach, a bright white slab at 3am,
no way to tell the committal action from the escape — and it blocks the event loop, so
the meter's own timer stopped while it was open. Replaced with an in-page sheet with
two full-width buttons, focus trapped, Escape to cancel.

**Accessibility** was nil — no ARIA, no roles, no focus management anywhere. Now: focus
moves to each page's heading on navigation (it was being dropped to `<body>`, so a
screen reader announced nothing when the page changed), every status line is a live
region, the fare is labelled but deliberately *not* live (it updates once a second;
announcing that would be unusable), address suggestions and history rows are real
buttons rather than clickable divs — there was previously **no keyboard path to setting
a destination at all** — and the tab bar carries `aria-current`.

Contrast: white on the old blue and red was 3.65:1 and 3.55:1, failing AA on the two
most important controls in the app, in a vehicle, behind a windscreen. All sixteen
colour pairs now pass, verified numerically.

Also: `charset` and `lang` were missing entirely, which on a plain static host turns
every `£` into `Â£` — on a fare display. Three `env(safe-area-inset-*)` calls were dead
because the viewport meta lacked `viewport-fit=cover`. Leaflet's zoom control painted
over the sticky header. Buttons gave no press feedback because the global tap-highlight
reset had no replacement, so a driver stabbing at Stop over a pothole got nothing. The
focus ring was blue-on-blue, i.e. invisible, on the primary button. All fixed.

## 6. Offline, and the rest

`sw.js` and `manifest.json` make it a proper installable app that opens without a
signal. The shell is network-first so a redeploy is picked up as soon as there is a
signal; Leaflet is cache-first since it is pinned; **map tiles and the backend are never
cached** — a stale driver record could let a suspended ID start a job.

Smaller things: a restored in-progress trip now re-checks the account in the background
(it was trusting a saved token indefinitely); rejected trips can be discarded instead of
failing on every sync forever, with the actual reason shown per trip; synced trips older
than 90 days can be pruned; the audit log takes the driver name from the sheet rather
than the payload and refuses a deactivated account; `repairTripDates` no longer reports
success when it repaired nothing; `deactivateUnnamedDrivers` writes one cell instead of
flattening the whole A:C block; `generateBulkIds` refuses a headerless sheet that would
have put an ID somewhere sign-in can never see; duplicate lookup is keyed per driver on
a null-prototype map and reads recent rows only; `doGet` answers with a health check so
"is the new version live" is one tap; the timezone check tells you whether script and
spreadsheet agree, because if they drift every trip datetime shifts.

---

## Deploy day

Order matters a little now.

1. **Have drivers sync and clear their queues on the old build first.** Anything still
   queued was metered at the old figures.
2. Paste `Code.gs`, then **Deploy → Manage deployments → edit → New version**.
3. Open the deployment URL in a browser. It now returns JSON — check `secret`,
   `routing`, `columns: 23`, and that `scriptTz` and `sheetTz` match.
4. Publish `index.html`, `sw.js` and `manifest.json` **to the same folder**. The service
   worker only controls pages at or below its own path.
5. First sync migrates the sheet automatically. Then **Taxi Meter → Build or rebuild
   Summary tab**.
6. **Taxi Meter → Check security setup** — confirms both secrets and the timezone pairing.

Still outstanding, and still only you can close them:

- **Email licensing about the E rate.** The card reads "30p each succeeding 15 yards",
  which is £35 a mile. The app uses 165 yd and says so on the Settings tab. **Confirm
  before 24 December**, when the E rate first goes live.
- **Confirm the Saturday Peak window** in the same email — the card's second clause is
  elliptical and I implemented the parallel reading, Sat 21:00 → Sun 08:00.
- **Rotate the OpenRouteService key** — the old one was readable in a previously
  deployed page.
- **Test on one handset of each model your drivers use.** The single most important
  check is still that the fare ticks on *distance* while moving. If it only ever ticks
  on time, that handset reports null speed and the fallback needs looking at. Also worth
  confirming: the "Trip started" announcement is audible on iPhone, and the app opens in
  airplane mode after one online visit.

## What I have not done

- **No light theme.** The UI is permanently dark. A black screen is a mirror; on a dash
  mount in daylight it reflects the windscreen. The app already knows the time of day —
  it switches tariff on it — but never switches theme. Worth doing; it is a real piece
  of work rather than a tweak.
- **No archival.** Trips and Log grow in one spreadsheet indefinitely. Fine for a year
  or two, wants a plan eventually.
- **No way to void or correct a trip** after it has synced.
- **The Summary formulas still address columns by letter** while the script uses named
  constants. The layout guard now catches a reorder at write time, which was the
  dangerous case, but the two conventions still coexist.
