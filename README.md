# Parts Order App

Flask + Socket.IO app (same architecture as the Water Spider dispatch
system): reads `inventory.xlsx`, builds a pick ticket that syncs live
across every open device, and exports a 2-column order file.

## Deploy to Render (same as the dispatch system)

1. Push this folder to a GitHub repo — include `inventory.xlsx`.
2. Render -> New -> Web Service -> connect the repo.
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn -c gunicorn.conf.py`
   (render.yaml sets these automatically if Render picks it up.)
3. Open the URL from any device. Add to Home Screen on iPhone/iPad for
   an app-style icon.

Updating inventory or the EMAIL_ADDRESSES autofill list = edit, commit,
push; Render redeploys.

## How sync works

Every edit (quantities, on-hand counts, removed lines, email box) saves
to the server and is pushed instantly to all connected devices over
Socket.IO — refresh-proof and device-proof. Exports also push, so the PO
number and history stats update everywhere at once.

## Known limits on Render's free tier

- The service sleeps when idle (first load takes ~30-60s), and its disk
  resets on every deploy/restart — so the PO counter, history stats, and
  any in-progress draft reset then too (same behavior as the dispatch
  system's delivered logs).
- "Send via Email" hides itself on Render: Outlook automation needs
  Windows, and Render free blocks SMTP. Use "Export to Excel" — on an
  iPhone the downloaded file can be shared straight into Mail.

## Run locally instead (Windows)

`pip install -r requirements.txt` then `python app.py` ->
http://localhost:5000. On Windows with Outlook desktop signed in, the
"Send via Email" button appears and sends through Outlook.

## PO reset

Delete `order_history.json` next to app.py (or remove today's entries
from it). Resets to 1 automatically each new day regardless.
