# MSP Ops Center: new interface (preview)

React + TypeScript, built with Vite. The build output goes to `../static/app` and is served by
`msp_runbook.py` at **/app/**, offline, under the same security policy as the classic view.
The server never needs Node.js; only the PC that builds this does.

## Build
- Windows: double-click `Build_Frontend.bat`.
- Any system: `npm ci` then `npm run build`.

## Develop
Start MSP Ops Center on port 8080, then run `npm run dev` and open http://localhost:5173/app/.
API and font requests are passed through to the running server.

## What has moved
| Screen | Status |
|---|---|
| Sign-in | New interface |
| My work | New interface |
| CMDB, with a detail panel per CI (facts, special care, runbook, incidents) | New interface (read-only; add and edit in the classic view) |
| Ask assistant, on-shift strip, light and dark themes | New interface |
| Every other screen | Opens in the classic view (marked "classic" in the sidebar) |

## Layout
- `src/lib`: API client (session cookie, CSRF header), types, hash router, helpers
- `src/components`: app shell, sign-in, on-shift strip, Ask panel, icons
- `src/pages`: one file per screen
- `src/styles/app.css`: the design system (shared tokens with the classic view, contrast-checked)

Dependencies are pinned to exact versions in `package.json` and `package-lock.json`.
