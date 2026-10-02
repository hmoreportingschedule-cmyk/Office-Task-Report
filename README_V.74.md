# Office Task Report V.74

This update is based on the V.73 codebase and addresses the issues identified in the full review.

## Included fixes
- Frontend lazy XLSX import syntax corrected.
- Employee self profile photo upload now uses the profileUpdate flow, with server-side self-photo permission fixed.
- Profile photo loading no longer downloads full base64 photo data on every profile read; profile responses use the saved Drive URL.
- Office In/Out fallback from EmployeeSettings remains supported in Users.
- Monthly template assignment now creates ONE task row for the selected Date-To-Date period, not one row per day.
- Existing duplicate daily rows from the older assignment model are migrated once per employee file by combining actual minutes/progress and removing duplicate rows.
- Task month filtering includes tasks whose Date-To-Date period overlaps the selected month.
- Task completion / 100% progress is blocked before the assignment end date.
- Employee yearly sheet schema version is advanced to V.74 while preserving the existing five-sheet structure and matching columns by name.
- Master support-sheet schema version is advanced to V.74; the existing controlled schema normalizer preserves data by column name.

## Deployment
1. Replace the root Google Apps Script `Code.gs` and deploy a new Apps Script version.
2. Replace `src/main.jsx` in GitHub.
3. Do not use `Apps Script/Code.gs`; the root `Code.gs` is the canonical backend.

## Validation
`Code.gs` and `main.jsx` were syntax-checked in the available runtime. A full Vite production build was not run because dependency installation/build tooling was not available in the runtime.
