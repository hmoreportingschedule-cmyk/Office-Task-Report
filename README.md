# Office Task Report — V.81 Full Stable Build

Complete Cloudflare Workers + Google Apps Script project.

## Cloudflare
- Build command: `npm run build`
- Deploy command: `npx wrangler deploy`
- Root directory: `/`

The build runs a prebuild regression check before Vite. It checks required files, frontend/backend API action parity, the known JSX declaration failure, XLSX import mode, Wrangler assets, and Apps Script redirect handling.

## Backend
Deploy the root `Code.gs` in the existing Google Apps Script Web App project. Keep the existing `/exec` deployment URL configured in `worker.js`.

## Drive
`My Drive / Dashboard Working / office-task-report / Employees Task Files`

`My Drive / Dashboard Working / office-task-report / Employees Photo`

Employee yearly files: `Name_EmployeeId_Year` with only `Profile`, `Attendance`, `Tasks`, `Requests`, `Activities`.

Master spreadsheet: `office-task-report`; Users remains the finalized A:S schema.
