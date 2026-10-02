# Office Task Report V.83 — Deep Review Checklist

## Static / integration checks
- Code.gs helper references reviewed; undefined underscore-suffixed helpers now fail the prebuild check.
- Worker syntax reviewed.
- API proxy syntax reviewed.
- Frontend JSX parsed with the installed TypeScript parser.
- Frontend API actions cross-checked against Code.gs route actions.
- Route action handlers cross-checked against function declarations.
- Known Cloudflare single-statement declaration regression blocked.
- Dynamic XLSX import regression blocked.
- Manual Apps Script redirect replay blocked; Worker must use redirect: follow.

## Logic fixes in V.83
- Added missing `formatAttendanceDate_()` used by `taskProgress_()`.
- Task COMPLETE now preserves cumulative manually entered `actualMinutes`; it no longer replaces them with elapsed clock duration.
- Removed obsolete My Requests CSV helper; Reports CSV remains available.
- Version labels updated to V.83.

## Cloudflare configuration
- Build command: `npm run build`
- Deploy command: `npx wrangler deploy`
- Root directory: `/`
- Wrangler main: `worker.js`
- Assets directory: `./dist`

## Verification limitation
The current environment could not complete npm dependency installation because external npm registry access timed out. Therefore an actual Cloudflare production build cannot be truthfully claimed as executed here. Static syntax, JSX parsing, route/action matching, helper-reference checks, and known regression checks were executed locally.
