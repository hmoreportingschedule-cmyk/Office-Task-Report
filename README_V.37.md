# Office Task Report — FULL V.37

Complete project package.

## Main login fix
The Cloudflare Worker now forwards the login POST directly to the existing Apps Script Web App with `redirect: "follow"` instead of manually replaying the Apps Script redirect chain.

## Existing backend
Apps Script URL:
https://script.google.com/macros/s/AKfycbyfL_lBPCkUlHzvv0iWvbUyNmhjSE7dVh6yqHp0L4E9JAs8S6e9jDx3v2dKku3ClK6M/exec

Existing Google Sheet:
My Drive / Dashboard Working / office-task-report

Do not change the Users sheet or passwords for this deployment.

## Deploy
npm install
npm run build
npx wrangler login
npx wrangler deploy

## Verify
After deployment open:
https://YOUR-WORKER-DOMAIN/api/office-task/health

Expected:
{
  "ok": true,
  "app": "Office Task Report",
  "worker": "V.37",
  "proxy": true
}

Then test the dashboard login using the credentials already present in the Users sheet.
