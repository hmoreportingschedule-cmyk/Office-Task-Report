# Office Task Report V.71

## Employee Photo Upload Compatibility Fix

Fixes `Unknown action: uploadProfilePhoto` when Admin/Master Admin uploads an employee profile photo from Employee ID/Profile.

The Admin profile photo action now uses the existing `updateEmployee` backend action and sends the photo in the same update payload. This keeps compatibility with deployments where the standalone `uploadProfilePhoto` action is not yet exposed.

### Changed file
- `main.jsx` → replace the existing `src/main.jsx`

### Google Apps Script
No `Code.gs` change is required for this fix. Keep the currently deployed canonical root `Code.gs`.
