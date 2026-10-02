# Office Task Report V.86 — Deep Review Checklist

## Code / syntax
- Code.gs syntax: PASS
- worker.js syntax: PASS
- api/office-task.js syntax: PASS
- Prebuild regression checker: PASS

## Integration
- Frontend API actions: 42/42 matched to Code.gs routes
- Code.gs route cases: 55 unique / 55 total
- Cloudflare Worker -> Apps Script redirect: redirect=follow
- Manual redirect replay: absent
- Worker upstream timeout: 30 seconds
- Browser API timeout: 40 seconds
- Same-origin API path: /api/office-task

## Fixed in V.86
- Historical attendance/task year now resolves to the correct employee yearly spreadsheet.
- Attendance save/punch/break operations resolve the employee file from the attendance date.
- Approved leave/weekoff-adjustment rules resolve from the relevant year's employee file.
- Bulk attendance import groups records by employee + year.
- Employee profile email is preserved when only mobile number is edited.
- Task minutes are limited to the task date range and blocked after completion/cancellation.
- Task mutation requests can target the selected task year.
- Duplicate setRamadanBreakFreeze route removed.
- api/office-task.js converted to ESM to match package.json type=module.
- Stale V.86/V.74/V.62/V.54/V.39/V.18 version labels cleaned from production UI/code comments.
- Undefined helper detection retained and strengthened.

## Known regression guards
- Dynamic XLSX import blocked.
- Unsafe declaration pattern blocked.
- break3NamazType regression blocked.
- 3 Din Qafila regression blocked.
- Worker manual redirect blocked.
- Worker timeout guard required.
- Date-aware employee file helper required.
- Profile email preservation guard required.
- Duplicate setRamadanBreakFreeze route blocked.
- CommonJS module.exports in api/office-task.js blocked.

## Cloudflare configuration
- Root directory: /
- Build command: npm run build
- Deploy command: npx wrangler deploy
- Wrangler main: worker.js
- Assets directory: ./dist

## Verification limitation
The local environment did not have the npm dependency tree installed and external npm installation timed out. Therefore an actual Vite production bundle was not executed locally. All available static syntax, integration, route, regression, and Worker checks passed. Cloudflare's connected build environment remains the final Vite compilation check.
