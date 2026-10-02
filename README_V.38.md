# Office Task Report — FULL V.38

## This version
- Manual employee IN/OUT time entry.
- Employee attendance date is restricted to Today and Previous 1 Day; future dates are rejected server-side.
- Weekoff is automatically applied per employee. If employee Weekoff is blank, the System Weekoff is used.
- Approved Leave and approved Weekoff Adjustment dates are automatically treated as non-working; employee does not select them manually.
- Admin User Create/Edit includes Office City, Office Address, Office In Time, Office Out Time and Weekoff.
- The existing Users A:S structure is preserved. Employee-specific settings are stored in a separate `EmployeeSettings` sheet.
- The Employee ID table uses an Edit pencil action instead of the old Settings button.
- Existing Cloudflare Worker login/proxy remains included.

## Google Apps Script
Replace/deploy the included `Code.gs` in the Apps Script project. Deploy as Web App (Execute as Me, access as Anyone).

Master Google Sheet:
`My Drive / Dashboard Working / office-task-report`

The backend automatically creates/uses an `EmployeeSettings` sheet. Existing Users data/passwords are not rebuilt for these settings.

## Cloudflare Worker
Deploy the complete project:

```bash
npm install
npm run build
npx wrangler login
npx wrangler deploy
```

After deployment, verify:
`https://YOUR-WORKER-DOMAIN/api/office-task/health`

Expected worker version: V.37 (the proxy layer is unchanged from the last working full project).
