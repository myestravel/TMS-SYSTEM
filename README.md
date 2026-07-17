# E.S. TRAVEL TMS Enterprise v2.0

## Files
- `index.html` — Netlify entry page
- `style.css` — preserved premium gold/white UI and responsive layout
- `script.js` — Booking, Assignment, Dashboard, WhatsApp, Calendar, Payment and sync logic
- `netlify.toml` — Netlify deployment configuration
- `Code.gs` — Google Sheets / Apps Script web app backend
- `Google_Sheet_Template.xlsx` — workbook template with all required sheets

## Netlify deployment
1. Upload the whole folder or ZIP to Netlify Deploys.
2. Netlify will publish `index.html` automatically.
3. HTTPS is required for normal browser storage, maps and external integrations.

## Google Sheets setup
1. Upload `Google_Sheet_Template.xlsx` to Google Drive and open it as Google Sheets.
2. Open **Extensions > Apps Script**.
3. Replace the editor content with `Code.gs`.
4. Set project timezone to **Asia/Kuala_Lumpur**.
5. Deploy as **Web app**:
   - Execute as: Me
   - Who has access: Anyone
6. Copy the `/exec` URL.
7. In `script.js`, replace `COMPANY_SHEET_WEB_APP_URL` with the deployed URL if the existing URL is not your production deployment.

## Main workflow
Confirm Booking performs:
1. Save Booking
2. Google Sheets sync
3. Dashboard refresh
4. Open Google Calendar event
5. Open Assignment page
6. Auto-select the current Booking

Assignment changes the status to **In Progress**, syncs Google Sheets, refreshes the Dashboard, opens the Calendar update and prepares Customer/Driver WhatsApp messages.

## Dashboard rules
- TODAY'S TRIPS: Departure or Return today, with Driver and Vehicle assigned
- UPCOMING TRIPS: future Departure and status In Progress
- PENDING: confirmed but not assigned
- COMPLETED: trip completed and payment paid
- WAITING FOR PAYMENT: trip completed, unpaid, 30 days or less
- TAKE ACTION: trip completed, unpaid, more than 30 days; sorted by Outstanding Days descending

## Important
The application retains browser local storage for offline continuity and uses Google Sheets cloud sync for multi-device shared data. Do not rename required sheet headers unless the corresponding code is updated.

## v2.0.1 Calendar Automation
Deploy the latest `Code.gs` as a new Web App version and approve Google Calendar permission. The default calendar is used unless the `SETTINGS` sheet contains a `CALENDAR_ID` value.

Release: v2.0.1 Booking-Calendar-Assignment Automation

IMPORTANT: After replacing `Code.gs`, run `setupTmsProject()` once in Apps Script, then deploy a New version. See `docs/SHEET_SYNC_SETUP.md`.
