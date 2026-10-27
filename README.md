# Dues & WhatsApp Reminders

A shareable customer dues tracker with WhatsApp reminder links and Google Sheets backup.

## 1. Create the Google Sheet

1. Create a Google Sheet. Keep the sheet open in Google Apps Script.
2. In Google Sheets, choose **Extensions > Apps Script**.
3. Replace the script editor contents with [apps-script.gs](apps-script.gs).
4. Select **Deploy > New deployment**.
5. Choose **Web app**.
6. Set **Execute as** to **Me** and **Who has access** to **Anyone**.
7. Deploy and copy the Web app URL.

## 2. Connect the app

1. Open the deployed website.
2. Choose **Settings** and paste the Web app URL in **Google Sheets Apps Script web app URL**.
3. Save Settings. The app writes a full backup to the `Payment Tracker Data` sheet and readable customer rows to `Customers`.

Changes are saved in the browser first, then sent to Google Sheets. A browser still needs to be configured before it can show cloud data. Re-opening a configured browser loads the latest cloud copy.

## 3. Publish the website

### GitHub Pages

1. Push this repository to GitHub.
2. Open **Settings > Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**.
4. Select the repository's default branch and `/ (root)`.
5. Open the Pages URL shown by GitHub and share that URL.

### Other static hosting

Upload or publish the repository root on Netlify, Cloudflare Pages, or any static host. The `index.html` file is the website entry point.

## Important data note

WhatsApp reminders open a pre-written message; you still press Send in WhatsApp. Spreadsheet backup protects data from browser loss, but the deployed Apps Script and Google Sheet remain the authority, so do not grant edit access to untrusted people.