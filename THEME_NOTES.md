# Light theme — and the first time anyone actually looked at this UI

Two things happened in this pass. The theme got built, and I put the app in a real
browser and took screenshots — which is how I found two defects that no amount of
reading the source would have surfaced.

**121 checks passing**, up from 106.

---

## 1. The theme system

Every colour in the app is now a named token declared in exactly one place, with a
value per theme. A theme is a swap of that block rather than a hunt through the
stylesheet. Two colour decisions that had been living in JavaScript — the tariff
chip's per-band swatch and the history sync badges — moved into CSS, because the
stylesheet is the only thing that knows which theme is in force. `updateTariff()`
now sets a class and nothing else.

**72 colour pairs are verified numerically against WCAG AA, in both themes.** Not
sampled — every foreground/background combination the app can produce, checked
against the tokens as they actually shipped rather than against my draft.

### The fare number

Amber on white is **1.4:1**. It is the obvious instinct and it is unusable — and it
happens to be the one number the entire product exists to display. The light theme
uses a deep amber, `#8A5600`, at **5.5:1**. It still reads as a taximeter rather
than as body text, and it survives direct sunlight.

### "Daylight" means actual daylight

You said the app already knows the time of day. It does — but a fixed light/dark
schedule would be wrong for half the year in Liverpool, where sunset runs from about
15:50 in December to 21:45 in June. A 07:00–19:00 rule would leave a driver on a
white screen at 5pm in December, an hour after dark, and on a black one at 8pm in
June in full sun.

So Auto computes real sunrise and sunset for Liverpool's coordinates — NOAA's
standard algorithm, about 25 lines. Verified against published times at the
solstices and both equinoxes; worst case is 4 minutes out, which is far inside
what matters for turning a screen over. Comparison is done in UTC on both sides,
so British Summer Time needs no special case.

The picker in Settings offers **Daylight / Light / Dark**, and the note under it
tells the driver what it is doing and why: *"Following daylight in Liverpool. Today
the sun rises at 05:52 and sets at 20:38, so the screen is currently light."*

It re-evaluates on the meter tick, when the app returns to the foreground, and on a
60-second timer so a parked driver still sees the screen turn over at dawn.

### Map tiles

OSM tiles are a bright white rectangle, which in dark mode was the brightest thing
on the screen at 3am. A filter is applied to the tile pane only, so the route line
and the destination marker keep their real colours.

---

## 2. What rendering it found

I had been reasoning about this interface statically for several rounds without once
looking at it. Two defects were invisible that way and obvious the moment a browser
drew them:

**Every page heading was painting a stray focus box.** Headings are given focus
programmatically on navigation so a screen reader announces the new page — that part
was a fix from the previous round and it was right. But the browser then draws its
focus ring around them, so every single page had what looked like a broken outlined
rectangle around its title. Suppressed for `tabindex="-1"` only, which never affects
a keyboard user because those elements are not in the tab order.

**Placeholder contrast was whatever the browser felt like.** `TX-A7K3M` in the Driver
ID field is instructional text — it tells the driver the format — and it was rendering
at the UA default grey. Pinned to a token: 4.9:1 dark, 6.1:1 light.

The screenshots also confirmed things I had only asserted: the sticky fare survives
both banners appearing at once, the confirm sheet sits correctly over a scrim, the
extras line reads *Meter £7.40 + £2.30 extras — charge £9.70* under the fare, the
GPS-lost banner is genuinely impossible to miss, and the three sync badge states are
distinguishable in both themes.

Contact sheets for both themes are attached — eight screens each, at iPhone size,
rendered from the actual file.

---

## 3. Deploy

No change to the deployment steps from the last note. `index.html` replaces the
previous one; `Code.gs`, `sw.js` and `manifest.json` are unchanged in this pass.

Bump `CACHE_VERSION` in `sw.js` when you publish, or phones with the old service
worker will keep serving the previous shell until it next revalidates.

---

## Still outstanding

Unchanged, and still only you can close them:

- **Email licensing about the E rate** — the card reads "30p each succeeding 15
  yards", which is £35 a mile. The app uses 165 yd and says so on the Settings tab.
  **Before 24 December.**
- **Confirm the Saturday Peak window** in the same email.
- **Rotate the OpenRouteService key.**
- **Test on one handset of each model your drivers use** — above all, that the fare
  ticks on *distance* while moving.

And the honest remainder, none of it blocking: there is still no way to void or
correct a trip after it has synced; Trips and Log grow in one spreadsheet with no
archival plan; and the Summary formulas still address columns by letter while the
script uses named constants, though the layout guard now catches a reorder at write
time, which was the dangerous half of that.
