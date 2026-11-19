# taxi-meter
Taxi Meter: a hackney carriage taximeter for the City of Liverpool tariff, as a
phone web app with a Google Sheet behind it.

| File | What it is |
|---|---|
| `index.html` | The app: meter, history, earnings, vehicle check, settings |
| `vehicle.js` | Daily vehicle check items and renewal-date rules |
| `pdf.js`, `vendor/jspdf.umd.min.js` | PDF receipts and earnings statements (jsPDF, MIT) |
| `sw.js`, `manifest.json`, icons | Installable app that opens without a signal |
| `Code.gs` | Apps Script bound to the Google Sheet |
| `tests/run-tests.mjs` | `TZ=Europe/London node tests/run-tests.mjs` (ends READY or NOT READY) |

Deploy steps for the latest change are in `DEPLOY.txt`.
