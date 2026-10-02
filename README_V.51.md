# Office Task Report V.51 — Cumulative Release

This is the cumulative implementation of V.40 through V.51+ in one package.

## Included
- V.40 Leave + Weekoff + Weekoff Adjustment
- V.41 Task Assignment + Task Status
- V.42 Task Timing + Progress
- V.43 Approval Center
- V.44 Notifications + Reminders
- V.45 Employee Dashboard
- V.46 HOD Dashboard / department-scoped analytics
- V.47 Admin Dashboard + Analytics
- V.48 Reports + Excel/CSV/PDF print export
- V.49 Audit Log + server-side role validation and login-failure audit
- V.50 Performance + low-internet optimizations (cache, compact payloads, same-origin Worker API, polling)
- V.51+ Hourly automation trigger for reminders and yearly rollover access

## Google Apps Script
There is ONLY ONE Apps Script file in this project:

`Code.gs`

Paste/deploy this file in the Google Apps Script project connected to the master Google Sheet.

Do NOT use an `Apps Script/Code.gs` duplicate.

## Master Sheet
Existing master Google Sheet remains:
`Dashboard Working / office-task-report / office-task-report`

The existing `Users` A:S schema is preserved exactly.

## Worker
Keep the existing `worker.js`, `api/office-task.js`, `wrangler.toml`, and same-origin API path:
`/api/office-task`

## Deployment order
1. Google Apps Script: replace the existing script with root `Code.gs` and deploy a new Web App version.
2. Cloudflare Worker: deploy the full project so `worker.js` continues proxying to the Apps Script Web App URL configured in `api/office-task.js`.
3. Open the dashboard and sign in.
4. MASTER_ADMIN can open Settings & Automation and enable the hourly trigger.

## Notes
- Existing Users/password data is not rebuilt or reordered.
- Employee yearly files remain `Name_EmployeeId_Year` under Employees Task Files.
- Employee photos remain under Employees Photo.
- Plain-text password column E remains because it is part of the finalized user requirement.
