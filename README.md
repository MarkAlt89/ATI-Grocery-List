# Parts Order App

Flask app that reads `inventory.xlsx`, lets you build a pick ticket, and
exports/emails a 2-column order file (Material Number, Quantity Ordered).

## Repo layout

```
parts-order-app/
├── app.py
├── inventory.xlsx        <- you add this (same file you use today)
├── requirements.txt
├── render.yaml
├── templates/
│   └── index.html
└── static/
    ├── app.js
    └── style.css
```

## Run locally

```
pip install -r requirements.txt
python app.py
```
Then open http://localhost:5000. On Windows with Outlook desktop, also
`pip install pywin32` if you want the email button to use Outlook locally.

## Deploy to Render (via GitHub)

1. Create a new GitHub repo and push this folder to it (include
   `inventory.xlsx`).
2. On https://render.com → **New → Web Service** → connect the repo.
   Render reads `render.yaml` automatically; if you set it up manually,
   use build command `pip install -r requirements.txt` and start command
   `gunicorn app:app`.
3. In the service's **Environment** tab, set:
   - `SMTP_HOST` — e.g. `smtp.gmail.com` (Gmail) or `smtp.office365.com` (Outlook/M365)
   - `SMTP_PORT` — `587`
   - `SMTP_USER` — the address the orders should come FROM
   - `SMTP_PASS` — an **app password** for that account (see below)
   - `APP_PASSWORD` — any password you choose; the site will require it
     before loading. Recommended, since the URL is public.
4. Deploy. To update inventory later: edit `inventory.xlsx`, commit,
   push — Render redeploys automatically.

### Getting an app password (Gmail example)

Google account → Security → turn on 2-Step Verification → search
"App passwords" → create one for "Mail". Use that 16-character string as
`SMTP_PASS` (your normal password won't work). Office 365 has an
equivalent, though some company tenants disable SMTP — if so, a free
Gmail just for sending orders is the easy workaround.

## PO numbering, order history, and resetting

The PO number is `PO-YYYY.MM.DD-N`, where **N = number of orders already
recorded today + 1**, based on `order_history.json` next to `app.py`.

**To reset the counter:** delete `order_history.json` (wipes the counter
AND the day/week/month history), or open it and delete just today's
entries (keeps older history; counter falls back accordingly). The number
also resets to 1 automatically every new day.

**Render caveat:** on the free tier the server's disk is wiped on every
deploy or restart, so `order_history.json` — and with it the PO counter
and history stats — resets whenever you push an update or the service
spins down from inactivity. If that matters, add a Render persistent disk
(requires a paid instance) and point `ORDERS_PATH` in `app.py` at the
disk's mount path, e.g. `/data/order_history.json`.

## Email autofill

Edit the `EMAIL_ADDRESSES` list near the top of `app.py` — those
addresses appear as suggestions when you type in the "Email to" box.
Commit + push to update the list on Render.
