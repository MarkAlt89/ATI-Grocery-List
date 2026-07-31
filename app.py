"""
Parts Order App — Flask backend

Reads the inventory catalog from a FIXED file sitting next to this script
(see SOURCE_FILENAME below), using FIXED column positions, serves the
parsed parts to the frontend as JSON, and exports the final order as a
2-column Excel file (Material Number, Quantity Ordered).

Column mapping (0-indexed from column A = 0):
    Column A (idx 0)  -> Part / material number
    Column C (idx 2)  -> Store location
    Column D (idx 3)  -> Description
    Column H (idx 7)  -> Extra field
    Column J (idx 9)  -> Qty on hand
    Column O (idx 14) -> Container usage per 16 hours

If your sheet's layout changes, update COLUMN_MAP below — nothing else
needs to change.

--------------------------------------------------------------------------
RUNNING IT (laptop + iPhone/iPad)
--------------------------------------------------------------------------
- Run `python app.py` on the Windows laptop (with Outlook installed and
  signed in). On the laptop, open http://localhost:5000.
- From an iPhone/iPad on the SAME Wi-Fi network, open
  http://<laptop-ip>:5000 (find the IP with `ipconfig` -> IPv4 Address).
  The first time, Windows will ask to allow Python through the firewall —
  allow it on Private networks, or the phone can't connect.
- Emails always send from the laptop's Outlook, no matter which device
  builds the order.
- Optionally set the APP_PASSWORD environment variable to require a
  login (useful on shared Wi-Fi).
"""

import io
import json
import os
import tempfile
from datetime import datetime, timedelta
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
from flask import Flask, request, jsonify, render_template, send_file, Response
from flask_socketio import SocketIO

# Outlook automation (Windows + Outlook desktop only). Not used on Render —
# there the app sends through SMTP instead (see send_via_smtp below).
try:
    import win32com.client as win32
    import pythoncom
    OUTLOOK_AVAILABLE = True
except ImportError:
    win32 = None
    pythoncom = None
    OUTLOOK_AVAILABLE = False

app = Flask(__name__)
app.config["SECRET_KEY"] = "parts-order-local-key"
# Real-time sync engine — same setup as the Water Spider dispatch system.
# Whenever the draft or history changes, the server pushes the update to
# every connected device instantly over Socket.IO.
socketio = SocketIO(app, cors_allowed_origins="*")

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

# Name of the inventory file. It must sit in the same folder as this script.
SOURCE_FILENAME = "inventory.xlsx"
SOURCE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), SOURCE_FILENAME)

# Column letter -> 0-indexed position (A=0, B=1, ... Q=16)
COLUMN_MAP = {
    "part_number": 0,   # A
    "location": 2,      # C
    "description": 3,   # D
    "extra_field": 7,   # H
    "qty_on_hand": 9,   # J
    "usage_16h": 14,    # O
}

# --------------------------------------------------------------------------
# EMAIL AUTOFILL LIST — edit this!
# Put the addresses you send orders to here. They show up as autocomplete
# suggestions in the "Email to" box (start typing and the browser offers
# matches; clicking the empty box shows the whole list in most browsers).
# --------------------------------------------------------------------------
EMAIL_ADDRESSES = [
    # "purchasing@yourcompany.com",
    # "jane.doe@supplier.com",
    # "warehouse@yourcompany.com",
]

# Optional app-wide password (HTTP Basic Auth). If APP_PASSWORD is set,
# the browser prompts for a login before showing the app — username can be
# anything, password must match. If unset (e.g. running on your laptop),
# no login is required.
APP_PASSWORD = os.environ.get("APP_PASSWORD", "").strip()

# Server-side order history: powers the day/week/month summary and the
# sequential PO numbers.
ORDERS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "order_history.json")

# The in-progress ("draft") order, shared by every device. Each edit on
# any device saves here; every page load restores it; open pages poll it
# so laptop and phone stay in sync. "rev" is a counter that bumps on
# every save, so clients can tell when something changed.
DRAFT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "draft.json")



def send_via_smtp(file_path, to_email, subject, body):
    msg = MIMEMultipart()
    msg['From'] = os.environ["EMAIL_FROM"]
    msg['To'] = to_email
    msg['Subject'] = subject

    # Email body
    msg.attach(MIMEText(body, "plain"))

    # File attachment
    with open(file_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
        msg.attach(part)

    # Connect securely (SSL)
    server = smtplib.SMTP_SSL(os.environ["SMTP_HOST"], int(os.environ["SMTP_PORT"]))
    server.login(os.environ["SMTP_USER"], os.environ["SMTP_PASS"])
    server.send_message(msg)
    server.quit()
def load_draft():
    if not os.path.exists(DRAFT_PATH):
        return {"rev": 0, "state": {}}
    try:
        with open(DRAFT_PATH, "r") as f:
            d = json.load(f)
        if isinstance(d, dict) and isinstance(d.get("rev"), int):
            return d
    except (json.JSONDecodeError, OSError):
        pass
    return {"rev": 0, "state": {}}


@app.before_request
def require_password():
    if not APP_PASSWORD:
        return None
    auth = request.authorization
    if auth and auth.password == APP_PASSWORD:
        return None
    return Response(
        "Login required.", 401,
        {"WWW-Authenticate": 'Basic realm="Parts Order"'},
    )


# --------------------------------------------------------------------------
# Catalog parsing
# --------------------------------------------------------------------------

def parse_catalog(file_stream):
    """Read the workbook and pull out the fixed columns."""
    df = pd.read_excel(file_stream, header=0, engine="openpyxl")

    max_idx = max(COLUMN_MAP.values())
    if df.shape[1] <= max_idx:
        raise ValueError(
            f"This sheet only has {df.shape[1]} column(s), but the app "
            f"requires at least {max_idx + 1} columns (up to Column O). "
            f"Double check your file layout.")

    def as_str(val, default=""):
        if pd.isna(val):
            return default
        return str(val).strip()

    def as_num(val):
        try:
            if pd.isna(val):
                return 0
            return float(val)
        except (TypeError, ValueError):
            return 0

    parts = []
    for _, row in df.iterrows():
        part_number = as_str(row.iloc[COLUMN_MAP["part_number"]])
        if not part_number:
            continue  # skip blank rows

        parts.append({
            "sku": part_number,
            "location": as_str(row.iloc[COLUMN_MAP["location"]], "Unassigned"),
            "description": as_str(row.iloc[COLUMN_MAP["description"]]),
            "extra_field": as_str(row.iloc[COLUMN_MAP["extra_field"]]),
            "qty_on_hand": as_num(row.iloc[COLUMN_MAP["qty_on_hand"]]),
            "usage_16h": as_num(row.iloc[COLUMN_MAP["usage_16h"]]),
        })

    return parts


# --------------------------------------------------------------------------
# Order history / PO numbering
# --------------------------------------------------------------------------

def load_order_history():
    if not os.path.exists(ORDERS_PATH):
        return []
    try:
        with open(ORDERS_PATH, "r") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return []


def save_order_history(history):
    with open(ORDERS_PATH, "w") as f:
        json.dump(history, f, indent=2)


def record_order_history(order_no, lines, extra=None):
    history = load_order_history()
    entry = {
        "order_no": order_no,
        "date": datetime.now().isoformat(),
        "line_count": len(lines),
        "total_units": sum(l["qty"] for l in lines),
        "lines": lines,
    }
    if extra:
        entry.update(extra)
    history.append(entry)
    save_order_history(history)
    # Nudge all devices to refresh their PO number and history stats.
    socketio.emit("history_update", {})


# --------------------------------------------------------------------------
# Excel + email helpers
# --------------------------------------------------------------------------

def build_order_excel(lines, path_or_buffer):
    out_df = pd.DataFrame([
        {"Material Number": l["sku"], "Quantity Ordered": l["qty"]}
        for l in lines
    ])
    with pd.ExcelWriter(path_or_buffer, engine="openpyxl") as writer:
        out_df.to_excel(writer, index=False, sheet_name="Order")
        ws = writer.sheets["Order"]
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 20


def clean_lines(raw_lines):
    return [
        {"sku": l.get("sku"), "qty": l.get("qty")}
        for l in raw_lines
        if l.get("sku") and isinstance(l.get("qty"), (int, float)) and l.get("qty") > 0
    ]


def send_via_outlook(file_path, to_email, subject, body):
    """Local-only fallback: drive the Outlook desktop app on Windows."""
    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = to_email
        mail.Subject = subject
        mail.Body = body
        mail.Attachments.Add(file_path)
        mail.Send()
    finally:
        pythoncom.CoUninitialize()


# --------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/parts", methods=["GET"])
def get_parts():
    """Read SOURCE_PATH fresh on every request."""
    if not os.path.exists(SOURCE_PATH):
        return jsonify({
            "error": f"Couldn't find '{SOURCE_FILENAME}' in the app folder. "
                     f"Put your inventory file there (or update SOURCE_FILENAME "
                     f"in app.py) and refresh."
        }), 404

    try:
        with open(SOURCE_PATH, "rb") as f:
            parts = parse_catalog(f)
    except Exception as exc:  # noqa: BLE001 — surface the message to the user
        return jsonify({"error": f"Couldn't read '{SOURCE_FILENAME}': {exc}"}), 400

    if not parts:
        return jsonify({
            "error": "No parts found. Check that part numbers are in column A "
                     "and the first row is a header row."
        }), 400

    locations = sorted({p["location"] for p in parts})

    return jsonify({
        "parts": parts,
        "locations": locations,
        "filename": SOURCE_FILENAME,
    })


@app.route("/api/emails", methods=["GET"])
def get_emails():
    """The saved email list, used for autofill in the 'Email to' box."""
    return jsonify({"emails": EMAIL_ADDRESSES})


@app.route("/api/draft", methods=["GET"])
def get_draft():
    """Current shared draft order (quantities, on-hand entries, etc.)."""
    return jsonify(load_draft())


@app.route("/api/draft", methods=["POST"])
def save_draft():
    """Overwrite the shared draft with this device's latest state.
    Last write wins — fine for one person moving between devices."""
    data = request.get_json(force=True) or {}
    state = data.get("state")
    if not isinstance(state, dict):
        return jsonify({"error": "Missing draft state."}), 400
    draft = {
        "rev": load_draft().get("rev", 0) + 1,
        "state": state,
        "updated": datetime.now().isoformat(),
    }
    with open(DRAFT_PATH, "w") as f:
        json.dump(draft, f, indent=2)
    # Push the new draft to every connected device immediately.
    socketio.emit("draft_update", draft)
    return jsonify({"rev": draft["rev"]})


@app.route("/api/order-no", methods=["GET"])
def next_order_no():
    """Next PO number = (# of orders already exported today) + 1, based on
    the server-side history file."""
    history = load_order_history()
    now = datetime.now()
    today_key = now.strftime("%Y-%m-%d")
    count_today = sum(1 for o in history if o.get("date", "").startswith(today_key))
    next_count = count_today + 1
    order_no = f"PO-{now.strftime('%Y.%m.%d')}-{next_count}"
    return jsonify({"order_no": order_no})


@app.route("/api/order-summary", methods=["GET"])
def order_summary():
    history = load_order_history()
    now = datetime.now()
    windows = {"day": timedelta(days=1), "week": timedelta(days=7), "month": timedelta(days=30)}

    summary = {}
    for label, delta in windows.items():
        cutoff = now - delta
        relevant = []
        for o in history:
            try:
                order_date = datetime.fromisoformat(o["date"])
            except (KeyError, ValueError):
                continue
            if order_date >= cutoff:
                relevant.append(o)

        distinct_skus = set()
        total_units = 0
        for o in relevant:
            for l in o.get("lines", []):
                distinct_skus.add(l.get("sku"))
                total_units += l.get("qty", 0)

        summary[label] = {
            "orders": len(relevant),
            "total_units": total_units,
            "distinct_parts": len(distinct_skus),
        }

    return jsonify(summary)



@app.route("/api/email-status", methods=["GET"])
def email_status():
    # On Render, use SMTP. On Windows, use Outlook.
    smtp_ready = (
        os.environ.get("SMTP_HOST") and
        os.environ.get("SMTP_USER") and
        os.environ.get("SMTP_PASS")
    )

    return jsonify({
        "available": bool(smtp_ready or OUTLOOK_AVAILABLE),
        "method": "smtp" if smtp_ready else ("outlook" if OUTLOOK_AVAILABLE else None)
    })



@app.route("/api/export", methods=["POST"])
def export_order():
    data = request.get_json(force=True) or {}
    lines = clean_lines(data.get("lines", []))
    order_no = (data.get("order_no") or "order").strip()

    if not lines:
        return jsonify({"error": "No parts in the order."}), 400

    buffer = io.BytesIO()
    build_order_excel(lines, buffer)
    buffer.seek(0)

    record_order_history(order_no, lines)

    safe_name = "".join(c for c in order_no if c.isalnum() or c in ("-", "_")) or "order"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{safe_name}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/export-email", methods=["POST"])
def export_email():
    data = request.get_json(force=True) or {}
    lines = clean_lines(data.get("lines", []))
    order_no = (data.get("order_no") or "order").strip()
    email = (data.get("email") or "").strip()

    if not email:
        return jsonify({"error": "Enter an email address to send to."}), 400
    if not lines:
        return jsonify({"error": "No parts in the order."}), 400

    safe_name = "".join(c for c in order_no if c.isalnum() or c in ("-", "_")) or "order"
    tmp_path = os.path.join(tempfile.gettempdir(), f"{safe_name}.xlsx")
    build_order_excel(lines, tmp_path)

    subject = f"Parts Order {order_no}"
    body = f"Parts order {order_no} is attached.\n\nLine items: {len(lines)}\nTotal units: {sum(l['qty'] for l in lines)}"

    try:
        send_via_smtp(tmp_path, email, subject, body)
    except Exception as exc:
        return jsonify({"error": f"SMTP send failed: {exc}"}), 500
    finally:
        try:
            os.remove(tmp_path)
        except:
            pass

    record_order_history(order_no, lines, extra={"emailed_to": email})
    return jsonify({"sent": True, "email": email})


if __name__ == "__main__":
    socketio.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
