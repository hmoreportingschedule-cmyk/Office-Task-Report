# Office Task Report V.85 — Deep Cloudflare/Integration Review

V.85 is the reviewed baseline. It keeps the existing Wrangler + Vite + Google Apps Script architecture and does not introduce the Cloudflare Vite plugin, so the existing Worker/assets deployment path remains stable.

## Cloudflare
- Root directory: `/`
- Build: `npm run build`
- Deploy: `npx wrangler deploy`
- Wrangler source of truth: `wrangler.toml`
- Worker entry: `worker.js`
- Static assets: `dist/` via ASSETS binding

## Apps Script
Replace the existing root `Code.gs` in the same Apps Script project and deploy a new Web App version.

## Validation
Prebuild checks cover required files, API action/route parity, undefined Apps Script underscore helpers, known syntax regressions, Worker redirect handling, assets configuration, version consistency, and duplicate critical routes.

Note: final `vite build` must still be executed in the connected Cloudflare build environment because this offline environment cannot complete npm registry installation.
