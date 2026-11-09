# Blue Ridge Entry Level Driver Training Website

GitHub Pages serves `blueridgeeldt.com` from `main`. The website uses its own Google Apps Script project and the existing **Blue Ridge ELDT website** workbook. Use a blank test workbook in a separate Apps Script project for acceptance testing.

## Secure deployment

The browser expects Apps Script to report `blue-ridge-v7-secure` from `?action=getVersion`. The older production endpoint reported `2026-09-05-blue-ridge-v5-license-state` during the November 9, 2026 audit. The secure branch blocks credential and request submission to that version.

1. In a separate Apps Script V8 project, add `app.js` and `bcrypt.js` as server-side files. The bcryptjs 3.0.3 license is in `bcrypt.LICENSE`. The vendored UMD file guards its asynchronous timer lookup because Apps Script only uses the synchronous methods.
2. Set script properties for that environment.

   | Property | Value |
   | --- | --- |
   | `BLUE_RIDGE_SPREADSHEET_ID` | ID of the blank acceptance workbook. At cutover, use the existing business workbook only. |
   | `BLUE_RIDGE_ADMIN_EMAIL` | A mailbox controlled by the tester. At cutover, use the business owner's notification mailbox. |
   | `BLUE_RIDGE_INTAKE_OPEN` | Leave unset or `false` until the target environment is ready for test submissions or customer intake. |
   | `BLUE_RIDGE_VIDEO_SECONDS` | JSON object mapping each assigned YouTube video ID to its verified duration in seconds, such as `{"example1234":1200}`. Enter measured values for both courses. |

3. Run `setupSheets_` from the Apps Script editor against the blank test workbook and configure synthetic admins, course modules, and questions. Use the editor-only `migrateBlueRidgePasswords` command to hash a synthetic admin password in the test workbook. Deploy the test project as a web app and point a local copy of `config.js` to its URL.
4. With synthetic records, verify each course's request and inbox notification, admin approval and assignment, student sign-in, successful embedded playback and save, locked final test before completion, failing and passing scores, and retained best score after reloading. Confirm that anonymous requests cannot read student data, lessons, or test answers. The browser's playback checks and the backend's elapsed-time check provide different safeguards. YouTube playback itself is observable only in the browser.
5. Before production cutover, verify the business workbook ID and tabs, back it up with restricted access, and check the actual modules, question banks, and passing scores for both courses. Confirm authorization to offer online ELDT for the applicable courses. For a production credential migration, take the legacy deployments out of service before running `migrateBlueRidgePasswords`. The migration updates only password cells, skips existing bcrypt hashes, and processes up to 25 credentials per sheet per run. Repeat until there are no remaining eligible plaintext cells. Passwords exceeding bcrypt's 72-byte limit require a reset. Rotate credentials previously used with the legacy URL-based login and handle restricted backups and sheet history as containing legacy credentials.
6. Publish the tested backend to the business deployment and verify its reported version. Merge the frontend only after acceptance and align its deployment URL if it changed. Keep customer intake closed until the endpoint, notification mailbox, workbook, and courses have been confirmed.

New training requests omit driver-license number and date of birth. Existing records and sheet headers remain compatible. If MailApp fails after the request is saved, the application preserves the pending request and records the failure in Apps Script execution logs. The admin dashboard remains the source for pending requests.

## Local checks

Run `node tests/workflow.cjs` from the repository root. It exercises the Apps Script request handlers with a synthetic in-memory workbook and mail sink. It does not contact Google or send messages.
