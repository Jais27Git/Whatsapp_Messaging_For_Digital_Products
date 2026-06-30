"""
===============================================================================
Project      : Daalchini Voucher WhatsApp Automation
Module       : Metabase → Google Sheets Sync + WhatsApp Sender
Author       : Ankit Kumar
Company      : Daalchini Technologies Pvt. Ltd.
Created On   : 2026-06-26
Python       : 3.11+
Version      : 2.2.0

Description
-----------
Single-script end-to-end pipeline:
  Step 1 → Fetch voucher records from Metabase
  Step 2 → Append only new rows to Google Sheets (dedup by Mobile + Order ID)
  Step 3 → Send WhatsApp messages to all PENDING + FAILED rows
  Step 4 → Update each row with SENT / FAILED status + timestamp + message ID

Google Sheet Columns
--------------------
A  Mobile
B  Order ID
C  Code          ← coupon_code
D  CTA
E  Claimed At
F  Status        ← written by sender (PENDING → SENT / FAILED)
G  Template      ← never overwritten by sender
H  Sent At       ← written by sender
I  Message ID    ← written by sender
J  Error         ← written by sender

Templates
---------
Discovery  → discovery_voucher
             Image via Firebase Storage URL link
             Body vars: order_id, coupon_code
             offer_desc is hardcoded in Meta template — not passed as variable
             copy_code button: hyphen stripped from coupon_code (Meta rejects hyphens)

Healthify  → healthify_msg
             Image via WhatsApp media ID
             Body vars: order_id, coupon_code

Known Constraints
-----------------
- Meta copy_code button rejects hyphens in coupon_code → strip before sending
- Meta copy_code button has a 15-char limit → always pass hyphen-stripped code
- Sheet updates use separate F and H:J ranges to avoid overwriting col G (Template)
- Sender retries all non-SENT rows (PENDING + FAILED) on every run

Dependencies
------------
pip install requests gspread google-auth tabulate

Change Log
----------
v2.2.0
- Removed offer_desc body parameter — now hardcoded in Meta template
- Strip hyphen from coupon_code for copy_code button only; body retains hyphen
v2.1.0
- Fixed sheet.update wiping col G (Template) — now updates F and H:J separately
- Sender retries FAILED rows in addition to PENDING (skips only SENT)
- Set HEALTHIFY_IMAGE_ID = 1558213535959134
v2.0.0
- Combined metabase_to_sheets.py + whatsapp_sender.py into single script
v1.1.0
- Fixed header creation for blank sheets (gspread phantom empty row bug)
- Switched console output to tabular format using tabulate
v1.0.0
- Initial implementation
===============================================================================
"""
import os
import json
import time
import requests
import gspread
from datetime import datetime
from google.oauth2.service_account import Credentials
from tabulate import tabulate

# =============================================================
# METABASE CONFIG
# =============================================================

METABASE_URL = os.environ["METABASE_URL"]
METABASE_API_KEY = os.environ["METABASE_API_KEY"]
CARD_ID = int(os.environ["CARD_ID"])

# =============================================================
# WHATSAPP CLOUD API CONFIG
# =============================================================

WA_PHONE_NUMBER_ID = os.environ["WA_PHONE_NUMBER_ID"]
WA_ACCESS_TOKEN = os.environ["WA_ACCESS_TOKEN"]

WA_API_URL = (
    f"https://graph.facebook.com/v25.0/"
    f"{WA_PHONE_NUMBER_ID}/messages"
)

WA_HEADERS = {
    "Authorization": f"Bearer {WA_ACCESS_TOKEN}",
    "Content-Type": "application/json",
}

# Delay between each WhatsApp API call (seconds) to avoid rate limits
SEND_DELAY_SECONDS = 0.5

# =============================================================
# DISCOVERY TEMPLATE CONFIG
# =============================================================

DISCOVERY_TEMPLATE_NAME = "discovery_voucher"
DISCOVERY_IMAGE_URL     = (
    "https://firebasestorage.googleapis.com/v0/b/daalchini-variant-portal.firebasestorage.app/o/Campaign%20Images%2FAnnual%20Plan.png?alt=media&token=847ee36e-7e9a-4053-8ddf-cd961ec098be"
)
DISCOVERY_OFFER_DESC    = "₹100 OFF Voucher code on 1-year Discovery channel subscription"

# =============================================================
# HEALTHIFY TEMPLATE CONFIG
# =============================================================

HEALTHIFY_TEMPLATE_NAME = "healthify_msg"
HEALTHIFY_IMAGE_ID      = "1558213535959134"

# =============================================================
# GOOGLE SERVICE ACCOUNT
# =============================================================

SERVICE_ACCOUNT_INFO = json.loads(os.environ["SERVICE_ACCOUNT_INFO"])
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
WORKSHEET_NAME = os.environ.get("WORKSHEET_NAME", "Sheet1")

# =============================================================
# SHEET COLUMN INDEX MAP  (0-based)
# =============================================================

COL_MOBILE     = 0   # A
COL_ORDER_ID   = 1   # B
COL_CODE       = 2   # C  ← coupon_code
COL_CTA        = 3   # D
COL_CLAIMED_AT = 4   # E
COL_STATUS     = 5   # F
COL_TEMPLATE   = 6   # G
COL_SENT_AT    = 7   # H
COL_MESSAGE_ID = 8   # I
COL_ERROR      = 9   # J

SHEET_HEADERS = [
    "Mobile", "Order ID", "Code", "CTA", "Claimed At",
    "Status", "Template", "Sent At", "Message ID", "Error",
]

# =============================================================
# HELPERS
# =============================================================

def format_phone(raw: str) -> str:
    """Normalize mobile to 91XXXXXXXXXX (No '+' prefix) for Meta WhatsApp API."""
    digits = "".join(ch for ch in str(raw) if ch.isdigit())
    if len(digits) == 10:
        return f"91{digits}"
    return digits


def now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def section(title: str, width: int = 100) -> None:
    print("\n" + "=" * width)
    print(f"  {title}")
    print("=" * width)


# =============================================================
# PAYLOAD BUILDERS
# =============================================================

def build_discovery_payload(phone: str, order_id: str, coupon_code: str) -> dict:
    button_code = coupon_code.replace("-", "")   # strip hyphen for copy_code button

    return {
        "messaging_product": "whatsapp",
        "to": phone,
        "type": "template",
        "template": {
            "name": DISCOVERY_TEMPLATE_NAME,
            "language": {"code": "en"},
            "components": [
                {
                    "type": "header",
                    "parameters": [
                        {"type": "image", "image": {"link": DISCOVERY_IMAGE_URL}}
                    ],
                },
                {
                    "type": "body",
                    "parameters": [
                        {"type": "text", "parameter_name": "order_id",    "text": order_id},
                        {"type": "text", "parameter_name": "coupon_code", "text": coupon_code},  # with hyphen — shown in message
                    ],
                },
                {
                    "type": "button",
                    "sub_type": "copy_code",
                    "index": "0",
                    "parameters": [
                        {"type": "coupon_code", "coupon_code": button_code}  # no hyphen — for copy button
                    ],
                },
            ],
        },
    }

def build_healthify_payload(phone: str, order_id: str, coupon_code: str) -> dict:
    return {
        "messaging_product": "whatsapp",
        "to": phone,
        "type": "template",
        "template": {
            "name": HEALTHIFY_TEMPLATE_NAME,
            "language": {"code": "en"},
            "components": [
                {
                    "type": "header",
                    "parameters": [
                        {"type": "image", "image": {"id": HEALTHIFY_IMAGE_ID}}
                    ],
                },
                {
                    "type": "body",
                    "parameters": [
                        {"type": "text", "parameter_name": "order_id",    "text": order_id},
                        {"type": "text", "parameter_name": "coupon_code", "text": coupon_code},
                    ],
                },
                {
                    "type": "button",
                    "sub_type": "copy_code",
                    "index": "0",
                    "parameters": [{"type": "coupon_code", "coupon_code": coupon_code}],
                },
            ],
        },
    }


# =============================================================
# WHATSAPP SEND
# =============================================================

def send_whatsapp(payload: dict) -> tuple[bool, str, str]:
    """Returns (success, message_id, error)."""
    try:
        resp = requests.post(WA_API_URL, headers=WA_HEADERS, json=payload, timeout=15)
        data = resp.json()
        if resp.status_code == 200 and "messages" in data:
            return True, data["messages"][0].get("id", ""), ""
        error_msg = (
            data.get("error", {}).get("message", "")
            or data.get("error", {}).get("error_data", {}).get("details", "")
            or str(data)
        )
        return False, "", error_msg[:200]
    except requests.exceptions.RequestException as e:
        return False, "", str(e)[:200]


# =============================================================
# STEP 0 — CONNECT TO GOOGLE SHEETS
# =============================================================

scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
credentials = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=scopes)
gc    = gspread.authorize(credentials)
sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(WORKSHEET_NAME)

# =============================================================
# STEP 1 — ENSURE HEADER ROW EXISTS
# =============================================================

existing = sheet.get_all_values()
existing = [row for row in existing if any(cell.strip() for cell in row)]

if not existing:
    sheet.insert_row(SHEET_HEADERS, 1)
    existing = [SHEET_HEADERS]
elif existing[0] != SHEET_HEADERS:
    sheet.insert_row(SHEET_HEADERS, 1)
    existing = [SHEET_HEADERS] + existing

# =============================================================
# STEP 2 — FETCH FROM METABASE
# =============================================================

section("STEP 1 / 3  —  METABASE FETCH")

url      = f"{METABASE_URL}/api/card/{CARD_ID}/query/json"
response = requests.post(url, headers={"X-API-Key": METABASE_API_KEY})
response.raise_for_status()
records  = response.json()

print(f"  Fetched {len(records)} records from Metabase card {CARD_ID}")

# =============================================================
# STEP 3 — SYNC NEW ROWS TO SHEET
# =============================================================

section("STEP 2 / 3  —  GOOGLE SHEETS SYNC")

# Build existing key set from current sheet data
existing_keys = set()
for row in existing[1:]:
    if len(row) >= 2:
        existing_keys.add(f"{row[0]}|{row[1]}")

print(f"  Rows already in sheet : {len(existing_keys)}")

# Auto-detect Metabase column names
first       = records[0]
phone_key   = next(k for k in first if "Mobile"   in k)
order_key   = next(k for k in first if "Order ID" in k)
code_key    = "Code"
cta_key     = next(k for k in first if "Cta"      in k)
claimed_key = next(k for k in first if "Claimed"  in k)

rows_to_insert = []
preview_table  = []

for r in records:
    phone      = str(r[phone_key]).strip()
    order_id   = str(r[order_key]).strip()
    code       = str(r[code_key]).strip()
    cta        = str(r[cta_key]).strip()
    claimed_at = str(r[claimed_key]).strip()
    key        = f"{phone}|{order_id}"

    if key in existing_keys:
        continue

    template = (
        "Discovery" if cta == "Claim Now"
        else "Healthify" if cta == "Go to HealthyfyMe"
        else "Unknown"
    )

    sheet_row = [phone, order_id, code, cta, claimed_at, "PENDING", template, "", "", ""]
    rows_to_insert.append(sheet_row)
    existing_keys.add(key)   # prevent duplicates within this batch
    preview_table.append([phone, order_id, code, cta, claimed_at, "PENDING", template, "-", "-", "-"])

if preview_table:
    print(tabulate(preview_table, headers=SHEET_HEADERS, tablefmt="simple", stralign="left"))
else:
    print("  (no new records to add)")

if rows_to_insert:
    sheet.append_rows(rows_to_insert)
    print(f"\n  ✅ Added {len(rows_to_insert)} new rows to sheet.")
else:
    print(f"\n  ✅ No new rows to add.")

# =============================================================
# STEP 4 — SEND WHATSAPP TO ALL PENDING ROWS
# =============================================================

section("STEP 3 / 3  —  WHATSAPP SENDER")

# Re-read sheet so newly appended rows are included
all_rows = sheet.get_all_values()
data_rows = all_rows[1:]   # skip header

pending = [
    (sheet_row_num, row)
    for sheet_row_num, row in enumerate(data_rows, start=2)
    if len(row) > COL_STATUS and row[COL_STATUS].strip().upper() not in ("SENT", "")
]

print(f"  PENDING / FAILED rows to send : {len(pending)}\n")

send_summary = []

for sheet_row_num, row in pending:
    raw_phone   = row[COL_MOBILE].strip()    if len(row) > COL_MOBILE   else ""
    order_id    = row[COL_ORDER_ID].strip()  if len(row) > COL_ORDER_ID else ""
    coupon_code = row[COL_CODE].strip()      if len(row) > COL_CODE     else ""
    template    = row[COL_TEMPLATE].strip()  if len(row) > COL_TEMPLATE else ""
    phone       = format_phone(raw_phone)

    # Build payload
    if template.strip().lower() == "discovery":
        payload = build_discovery_payload(phone, order_id, coupon_code)
    elif template.strip().lower() == "healthify":
        payload = build_healthify_payload(phone, order_id, coupon_code)
    else:
        sheet.update(range_name=f"F{sheet_row_num}", values=[["FAILED"]])
        sheet.update(range_name=f"H{sheet_row_num}:J{sheet_row_num}", values=[["", "", f"Unknown template: {template}"]])
        send_summary.append([sheet_row_num, phone, order_id, coupon_code, template, "FAILED", "-", f"Unknown template: {template}"])
        print(f"  Row {sheet_row_num} → SKIPPED  (unknown template: '{template}')")
        continue

    # Send
    success, message_id, error = send_whatsapp(payload)

    if success:
        sent_at = now_str()
        sheet.update(range_name=f"F{sheet_row_num}", values=[["SENT"]])
        sheet.update(range_name=f"H{sheet_row_num}:J{sheet_row_num}", values=[[sent_at, message_id, ""]])
        send_summary.append([sheet_row_num, phone, order_id, coupon_code, template, "SENT", sent_at, message_id])
        print(f"  Row {sheet_row_num} → ✅ SENT    {phone}  {coupon_code}  {message_id}")
    else:
        sheet.update(range_name=f"F{sheet_row_num}", values=[["FAILED"]])
        sheet.update(range_name=f"H{sheet_row_num}:J{sheet_row_num}", values=[["", "", error]])
        send_summary.append([sheet_row_num, phone, order_id, coupon_code, template, "FAILED", "-", error])
        print(f"  Row {sheet_row_num} → ❌ FAILED  {phone}  {coupon_code}  {error}")

    time.sleep(SEND_DELAY_SECONDS)

# =============================================================
# FINAL SUMMARY
# =============================================================

section("PIPELINE SUMMARY")

sent_count   = sum(1 for r in send_summary if r[5] == "SENT")
failed_count = sum(1 for r in send_summary if r[5] == "FAILED")

if send_summary:
    print(tabulate(
        send_summary,
        headers=["Sheet Row", "Mobile", "Order ID", "Code", "Template", "Status", "Sent At", "Message ID / Error"],
        tablefmt="simple",
        stralign="left",
    ))

print(f"""
  Metabase records fetched  : {len(records)}
  New rows added to sheet   : {len(rows_to_insert)}
  WhatsApp messages sent    : {sent_count}
  WhatsApp messages failed  : {failed_count}
""")
