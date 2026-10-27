# CTG Dash App (Millas)

Weekly scoring app for the Millas program. Managers score employees; the home screen shows the race and the performance tables (English and Spanish).

It is a Google Apps Script web app that reads and writes the Google Sheet "CTG Millas App".

## Files

| File | What it is |
|---|---|
| `src/Code.gs` | Server side: reads the sheet, saves scores |
| `src/Index.html` | The app screen (one file, English and Spanish texts inside) |
| `src/appsscript.json` | Apps Script settings (add it once, see below) |
| `.github/workflows/deploy.yml` | Sends the code to Apps Script on every push to `main` |

## How the sheet is wired

- `DASH` is the main tab. Tasks come from `DD`, the master list (Spanish text in A to E, English text in J to L, same row).
- `IMPORT ...` tabs bring in order seconds, finished planting, plugs and labels. Goals per hour are in row 4 of each planting tab.
- The sheet ID and tab names are at the top of `src/Code.gs`.

## Working on it

1. Change files in `src/`, open a pull request, merge to `main`.
2. The deploy action updates the live app. The app link does not change.

Editing directly in the Apps Script editor is no longer needed. If you do, copy the change back here, or the next push overwrites it.

## One-time setup

1. Apps Script: turn on the Apps Script API at https://script.google.com/home/usersettings.
2. Apps Script: Project Settings. Copy the **Script ID** into `.clasp.json`. Tick "Show appsscript.json manifest file in editor", then copy that file's content into `src/appsscript.json`.
3. On a computer with Node: `npm install -g @google/clasp@2.5.0`, then `clasp login`. This creates `~/.clasprc.json`.
4. GitHub: Settings > Secrets and variables > Actions > New repository secret. Name `CLASPRC_JSON`, value = the full content of `~/.clasprc.json`.
5. Commit and push to `main`. Check the Actions tab for a green run.

If the Google login expires, repeat steps 3 and 4.
