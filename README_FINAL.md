# Office Task Report — Final Stable Build

## Canonical Google Apps Script
Use **root `Code.gs` only**. The `Apps Script/Code.gs` duplicate is intentionally not included.

## Google Drive structure
- `My Drive/Dashboard Working/office-task-report/office-task-report` — master spreadsheet
- `My Drive/Dashboard Working/office-task-report/Employees Task Files` — yearly employee files
- `My Drive/Dashboard Working/office-task-report/Employees Photo` — employee profile photos

## Yearly employee spreadsheet
Each employee gets one yearly file: `Name_EmployeeId_Year` with only:
- Profile
- Attendance
- Tasks
- Requests
- Activities

## Deployment
1. Replace root `Code.gs` in the existing Apps Script project and deploy a new version.
2. Deploy the complete Worker/frontend project.
3. Keep the existing master spreadsheet and Users A:S schema.
4. Do not create a second Apps Script project or use `Apps Script/Code.gs`.

## Final review note
The backend and frontend source files were syntax-checked. A full production Vite build should be run in an environment with npm dependencies installed before production deployment.
