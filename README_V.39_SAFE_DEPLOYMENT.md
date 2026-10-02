# Office Task Report V.39 — Safe Core Testing & Production Approval

This version adds a safer GitHub deployment workflow so code is tested before it can reach production.

## Flow

feature/* -> develop -> Core Tests -> Staging -> test -> main -> Core Tests -> Production Approval -> Production

A failed core test stops deployment.

## GitHub setup (one time)

Create two GitHub Environments:
- `staging`
- `production`

For `production`, enable **Required reviewers** and select the Master Admin/reviewer who must approve a production deployment.

Add these GitHub Actions secrets to the repository/environment:
- `CLOUDFLARE_API_TOKEN`
- `CLOUDFLARE_ACCOUNT_ID`

Add `APPS_SCRIPT_URL` as an **environment secret** in both environments:
- `staging`: use a TEST Apps Script Web App URL connected to a TEST Google Sheet.
- `production`: use the real Apps Script Web App URL connected to the real master Sheet.

Do NOT put the Apps Script URL in frontend code or commit a production secret.

## Branches

- `feature/*`: development
- `develop`: staging
- `main`: production

## Automatic checks

Every push/PR to `develop` or `main` runs:
- required-file checks
- frontend API route check
- Worker route check
- health-route check
- JSON response check
- production build

## Production rule

Pushing to `main` runs the tests first. Only after all tests pass does GitHub enter the `production` environment. If Required Reviewers are configured, the deployment waits for approval.

## Important

Production and staging must use different Apps Script Web Apps and different Google Sheets. Never run staging tests against the real `office-task-report` Sheet.

Current production master path:
My Drive / Dashboard Working / office-task-report
