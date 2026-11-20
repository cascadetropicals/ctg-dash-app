# CTG Dash App (Millas)

Weekly scoring app for the Millas program. Managers score employees; the home screen shows the race and the performance tables (English and Spanish). Performance also has Weekly, Today, Data accuracy, Timeline and Daily summary views, and there is an Ask tab.

Live app: https://cascadetropicals.github.io/ctg-dash-app/ (sign in with Google, then pick your name).

## How it fits together

- **Page (this repo, GitHub Pages):** `index.html` is the whole front end in one file. It is served from the repo root.
- **Backend (Google Apps Script web app):** `apps-script/Code.gs` reads and writes the Google Sheet "CTG Millas App". The page talks to it with POST requests (`doPost`).
- **Sign-in:** Google Identity Services on the page. The backend checks the Google ID token, then hands out its own signed session (valid 7 days). Any Google account can sign in. Client ID and setup are the shared ones from the CTG Planting app (project `ctg-planting-app`); authorized origin `https://cascadetropicals.github.io`.
- **Old address:** the Apps Script web app address still serves the previous version of the page through `doGet`, and the reminder mails still link to it until they are switched.

## Files

| File | What it is |
| --- | --- |
| `index.html` | Hosted page (served by GitHub Pages) |
| `apps-script/Code.gs` | Backend: sheet reads and writes, `doGet` (old page), `doPost` (sign-in and function calls) |
| `apps-script/Index.html` | The page as served by Apps Script (source of `index.html`) |
| `apps-script/Mail.gs` | Friday reminder mail |
| `apps-script/Sheets.gs` | Pastes each employee's DASH table into their own sheet |
| `apps-script/appsscript.json` | Apps Script manifest (web app runs as the owner, access: anyone) |

## Keeping the two pages in sync

`index.html` is `apps-script/Index.html` with three changes: a viewport meta tag and title in the head; a small script at the top that replaces `google.script.run` with calls to the backend after Google sign-in; and `DEEP` (reminder link parameters) read from the page address instead of being filled in by the server. Any change to the app must be made in both files.

## Updating the backend

Apps Script editor, then Deploy, Manage deployments, edit the existing deployment, New version. This keeps the web app address the same.
