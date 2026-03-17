# Plankton Settlement Uploader — Mac Build Instructions

## 1. Clone the repo

```bash
git clone https://github.com/HendrikPlankton/PlanktonSettlementUploader.git
cd PlanktonSettlementUploader
```

## 2. Set up Python environment

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
pip install pyinstaller
```

## 3. Add Google credentials

Place `client_secret.json` in the project root (next to `sheets_uploader.py`).

Get it from: Google Cloud Console → APIs & Credentials → OAuth 2.0 Client ID (Desktop app) → Download JSON.

Make sure the **Google Sheets API** is enabled on the project.

## 4. Run (to verify it works)

```bash
python sheets_uploader.py
```

First run will open a browser for Google sign-in. After that, `token.json` is cached.

## 5. Build standalone .app

```bash
pyinstaller --onefile --windowed \
  --name "Plankton Uploader" \
  --add-data "client_secret.json:." \
  sheets_uploader.py
```

The built app will be at:

```
dist/Plankton Uploader.app
```

## 6. Sign and distribute

```bash
codesign --force --deep --sign "Developer ID Application: YOUR NAME (TEAM_ID)" "dist/Plankton Uploader.app"
```

Or for ad-hoc (no Apple Developer account):

```bash
codesign --force --deep --sign - "dist/Plankton Uploader.app"
```

The `.app` can then be zipped and sent to the end user — they just double-click to run.

## Notes

- If macOS blocks the unsigned app: System Settings → Privacy & Security → "Open Anyway"
- `token.json` is generated on first run (Google OAuth). If bundled inside the app, the user won't need to re-authenticate. If not bundled, they'll sign in once via browser.
- To bundle `token.json` too (after first auth): add `--add-data "token.json:."` to the pyinstaller command
