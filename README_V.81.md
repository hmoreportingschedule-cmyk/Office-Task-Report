# V.81 — Full Stable / Reviewed Build

This is the complete project ZIP.

Reviewed areas:
- Cloudflare/Vite build syntax and XLSX import.
- Apps Script ContentService redirect proxy.
- Login timeout and session-response validation.
- Frontend/backend API action parity.
- Attendance break preservation and Ramadan lunch freeze.
- No unintended break3NamazType column.
- Finalized Outdoor task categories.
- Task start/minutes date-range rules without silent completion.
- Admin user-management permissions while Master Admin stays protected.
- Employee folder-path fallback.

Deploy using Cloudflare Build `npm run build`, Deploy `npx wrangler deploy`, Root `/`.
Update the existing Apps Script Web App from the root `Code.gs`.
