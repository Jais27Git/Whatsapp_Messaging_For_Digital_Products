"""
===============================================================================
Project      : Daalchini Voucher WhatsApp Automation
Module       : Metabase → Google Sheets Sync + WhatsApp Sender
Author       : Ankit Kumar
Company      : Daalchini Technologies Pvt. Ltd.
Created On   : 2026-06-26
Python       : 3.11+
Version      : 2.5.1 (GitHub Actions Ready)
===============================================================================
"""
import os
import json
import time
import sys
import requests
import gspread
from datetime import datetime
from google.oauth2.service_account import Credentials
from tabulate import tabulate

# =============================================================
# ENVIRONMENT & SECRETS CONFIG
# =============================================================

METABASE_URL     = os.environ.get("METABASE_URL", "https://reporting.daalchini.co.in").rstrip("/")
METABASE_API_KEY = os.environ.get("METABASE_API_KEY", "")
CARD_ID          = int(os.environ.get("CARD_ID", "16663"))

WA_PHONE_NUMBER_ID = os.environ.get("WA_PHONE_NUMBER_ID", "1111751922025459")
WA_ACCESS_TOKEN    = os.environ.get("WA_ACCESS_TOKEN", "")

WA_API_URL = f"https://graph.facebook.com/v25.0/{WA_PHONE_NUMBER_ID}/messages"
WA_HEADERS = {
    "Authorization": f"Bearer {WA_ACCESS_TOKEN}",
    "Content-Type": "application/json",
}

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1rWieGHFnumQRZ3mUOvE6hf6x165xjTLhNBbii3bqM70")
WORKSHEET_NAME = os.environ.get("WORKSHEET_NAME", "Sheet1")

SERVICE_ACCOUNT_INFO_RAW = os.environ.get("SERVICE_ACCOUNT_INFO", "")

SEND_DELAY_SECONDS = 2

# Verify essential secrets before running
if not METABASE_API_KEY:
    sys.exit("❌ Error: Missing METABASE_API_KEY environment variable.")
if not WA_ACCESS_TOKEN:
    sys.exit("❌ Error: Missing WA_ACCESS_TOKEN environment variable.")
if not SERVICE_ACCOUNT_INFO_RAW:
    sys.exit("❌ Error: Missing SERVICE_ACCOUNT_INFO environment variable.")

try:
    SERVICE_ACCOUNT_INFO = json.loads(SERVICE_ACCOUNT_INFO_RAW)
except json.JSONDecodeError as e:
    sys.exit(f"❌ Error: Invalid SERVICE_ACCOUNT_INFO JSON payload: {e}")

# =============================================================
# COLUMN DEFINITIONS & TEMPLATE MAPS
# =============================================================

REQUIRED_COLUMNS = {
    "Mobile":     ("mobile no", "mobile"),
    "Order ID":   ("order id",),
    "Code":       ("code", "coupon code"),
    "CTA":        ("cta text", "cta"),
    "Claimed At": ("claimed at", "claimed"),
    "Provider":   ("provider",),
}

DISCOVERY_TEMPLATE_NAME = "discovery_voucher"
DISCOVERY_IMAGE_URL     = (
    "https://firebasestorage.googleapis.com/v0/b/daalchini-variant-portal.firebasestorage.app/o/Campaign%20Images%2FAnnual%20Plan.png?alt=media&token=847ee36e-7e9a-4053-8ddf-cd961ec098be"
)

KREDMINT_TEMPLATE_NAME = "kredmint_voucher"
KREDMINT_LANGUAGE_CODE = "en"
KREDMINT_PRODUCT_NAME  = "Kredmint Voucher"

HEALTHIFY_TEMPLATE_NAME = "healthify_msg"
HEALTHIFY_IMAGE_ID      = "1558213535959134"

# STRICT: Keyword match from Metabase Provider value
PROVIDER_TEMPLATE_MAP = {
    "discovery": "Discovery",
    "kredmint":  "Kredmint",
    "healthify": "Healthify",
}

COL_MOBILE     = 0   # A
COL_ORDER_ID   = 1   # B
COL_CODE       = 2   # C
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


def find_col(columns, *names: str) -> str:
    def short(c: str) -> str:
        return c.split("→")[-1].strip().lower().replace("_", " ")

    wanted = [n.lower().replace("_", " ") for n in names]
    for n in wanted:
        for c in columns:
            if short(c) == n:
                return c
    for n in wanted:
        for c in columns:
            if n in short(c):
                return c
    raise RuntimeError(
        f"Column '{names[0]}' not found in Metabase data.\n"
        f"Available columns: {list(columns)}"
    )


def fetch_metabase_records() -> list:
    url  = f"{METABASE_URL}/api/card/{CARD_ID}/query/json"
    resp = requests.post(url, headers={"X-API-Key": METABASE_API_KEY}, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    if isinstance(data, dict):
        raise RuntimeError(f"Metabase query failed: {data.get('error') or data}")
    return data


def map_required_columns(columns) -> dict:
    mapping, missing = {}, []
    for field, names in REQUIRED_COLUMNS.items():
        try:
            mapping[field] = find_col(columns, *names)
        except RuntimeError:
            missing.append(field)
    if missing:
        raise RuntimeError(
            f"Metabase card {CARD_ID} is missing required column(s): {', '.join(missing)}\n"
            f"Available columns: {list(columns)}"
        )
    return mapping


def resolve_template(provider: str) -> str:
    if not provider:
        return "Unknown"
    p = str(provider).strip().lower()
    if p in ("", "none", "null", "nan"):
        return "Unknown"
    for keyword, template in PROVIDER_TEMPLATE_MAP.items():
        if keyword in p:
            return template
    return "Unknown"


# =============================================================
# PAYLOAD BUILDERS
# =============================================================

def build_discovery_payload(phone: str, order_id: str, coupon_code: str) -> dict:
    button_code = coupon_code.replace("-", "")
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
                    "parameters": [{"type": "image", "image": {"link": DISCOVERY_IMAGE_URL}}],
                },
                {
                    "type": "body",
                    "parameters": [
                        {"type": "text", "parameter_name": "order_id", "text": order_id},
                        {"type": "text", "parameter_name": "coupon_code", "text": coupon_code},
                    ],
                },
                {
                    "type": "button",
                    "sub_type": "copy_code",
                    "index": "0",
                    "parameters": [{"type": "coupon_code", "coupon_code": button_code}],
                },
            ],
        },
    }


def build_kredmint_payload(phone: str, order_id: str, coupon_code: str) -> dict:
    return {
        "messaging_product": "whatsapp",
        "to": phone,
        "type": "template",
        "template": {
            "name": KREDMINT_TEMPLATE_NAME,
            "language": {"code": KREDMINT_LANGUAGE_CODE},
            "components": [
                {
                    "type": "body",
                    "parameters": [
                        {"type": "text", "parameter_name": "product_name", "text": KREDMINT_PRODUCT_NAME},
                        {"type": "text", "parameter_name": "voucher_code", "text": coupon_code},
                    ],
                },
                {
                    "type": "button",
                    "sub_type": "url",
                    "index": "0",
                    "parameters": [{"type": "text", "text": coupon_code}],
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
                    "parameters": [{"type": "image", "image": {"id": HEALTHIFY_IMAGE_ID}}],
                },
                {
                    "type": "body",
                    "parameters": [
                        {"type": "text", "parameter_name": "order_id", "text": order_id},
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


TEMPLATE_BUILDERS = {
    "discovery": build_discovery_payload,
    "kredmint":  build_kredmint_payload,
    "healthify": build_healthify_payload,
}


def send_whatsapp(payload: dict) -> tuple[bool, str, str]:
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
# MAIN PIPELINE
# =============================================================

def main():
    # ---------------------------------------------------------
    # STEP 1 — METABASE FETCH
    # ---------------------------------------------------------
    section("STEP 1 / 3  —  METABASE FETCH")
    records = fetch_metabase_records()
    print(f"  Fetched {len(records)} records from Metabase card {CARD_ID}")

    col_map = {}
    if records:
        col_map = map_required_columns(list(records[0].keys()))
        print("\n  Column mapping:")
        for field, col in col_map.items():
            print(f"    {field:<11} ← {col}")

    # ---------------------------------------------------------
    # STEP 2 — CONNECT TO GOOGLE SHEETS
    # ---------------------------------------------------------
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials = Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=scopes)
    gc = gspread.authorize(credentials)
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(WORKSHEET_NAME)

    existing = sheet.get_all_values()
    existing = [row for row in existing if any(cell.strip() for cell in row)]

    if not existing:
        sheet.insert_row(SHEET_HEADERS, 1)
        existing = [SHEET_HEADERS]
    elif existing[0] != SHEET_HEADERS:
        sheet.insert_row(SHEET_HEADERS, 1)
        existing = [SHEET_HEADERS] + existing

    # ---------------------------------------------------------
    # STEP 3 — SYNC NEW ROWS TO SHEET
    # ---------------------------------------------------------
    section("STEP 2 / 3  —  GOOGLE SHEETS SYNC")
    existing_keys = set()
    for row in existing[1:]:
        if len(row) >= 2:
            existing_keys.add(f"{row[0]}|{row[1]}")

    print(f"  Rows already in sheet : {len(existing_keys)}")

    rows_to_insert = []
    preview_table  = []

    if records:
        phone_key    = col_map["Mobile"]
        order_key    = col_map["Order ID"]
        code_key     = col_map["Code"]
        cta_key      = col_map["CTA"]
        claimed_key  = col_map["Claimed At"]
        provider_key = col_map["Provider"]

        for r in records:
            phone      = str(r[phone_key]).strip()
            order_id   = str(r[order_key]).strip()
            code       = str(r[code_key]).strip()
            cta        = str(r[cta_key]).strip()
            claimed_at = str(r[claimed_key]).strip()
            provider   = str(r.get(provider_key) or "").strip()
            key        = f"{phone}|{order_id}"

            if key in existing_keys:
                continue

            template = resolve_template(provider)

            if template == "Unknown":
                initial_status = "FAILED"
                initial_error  = f"Skipped: Unknown/missing provider '{provider}'"
            else:
                initial_status = "PENDING"
                initial_error  = ""

            sheet_row = [phone, order_id, code, cta, claimed_at, initial_status, template, "", "", initial_error]
            rows_to_insert.append(sheet_row)
            existing_keys.add(key)
            preview_table.append([phone, order_id, code, cta, claimed_at, initial_status, template, "-", "-", initial_error or "-"])

    if preview_table:
        print(tabulate(preview_table, headers=SHEET_HEADERS, tablefmt="simple", stralign="left"))
    else:
        print("  (no new records to add)")

    if rows_to_insert:
        sheet.append_rows(rows_to_insert)
        print(f"\n  ✅ Added {len(rows_to_insert)} new rows to sheet.")
    else:
        print(f"\n  ✅ No new rows to add.")

    # ---------------------------------------------------------
    # STEP 4 — RE-EVALUATE PROVIDER & SEND WHATSAPP
    # ---------------------------------------------------------
    section("STEP 3 / 3  —  RE-EVALUATE PROVIDER + WHATSAPP SENDER")

    metabase_provider_lookup = {}
    if records:
        phone_key    = col_map["Mobile"]
        order_key    = col_map["Order ID"]
        provider_key = col_map["Provider"]

        for r in records:
            mb_phone    = str(r.get(phone_key) or "").strip()
            mb_order    = str(r.get(order_key) or "").strip()
            mb_provider = str(r.get(provider_key) or "").strip()
            if mb_phone and mb_order:
                metabase_provider_lookup[f"{mb_phone}|{mb_order}"] = mb_provider

    print(f"  Provider records available from Metabase : {len(metabase_provider_lookup)}")

    all_rows = sheet.get_all_values()
    data_rows = all_rows[1:]

    pending = [
        (sheet_row_num, row)
        for sheet_row_num, row in enumerate(data_rows, start=2)
        if len(row) > COL_STATUS
        and row[COL_STATUS].strip().upper() in ("PENDING", "FAILED")
    ]

    print(f"  PENDING / FAILED rows to evaluate : {len(pending)}\n")

    send_summary = []

    for sheet_row_num, row in pending:
        raw_phone    = row[COL_MOBILE].strip()    if len(row) > COL_MOBILE   else ""
        order_id     = row[COL_ORDER_ID].strip()  if len(row) > COL_ORDER_ID else ""
        coupon_code  = row[COL_CODE].strip()      if len(row) > COL_CODE     else ""
        old_template = row[COL_TEMPLATE].strip() if len(row) > COL_TEMPLATE else ""

        phone = format_phone(raw_phone)

        # Lookup from Metabase
        provider = metabase_provider_lookup.get(f"{raw_phone}|{order_id}")
        if provider is None:
            normalized_sheet_phone = format_phone(raw_phone)
            for mb_key, mb_provider in metabase_provider_lookup.items():
                mb_phone, mb_order = mb_key.split("|", 1)
                if format_phone(mb_phone) == normalized_sheet_phone and mb_order == order_id:
                    provider = mb_provider
                    break

        provider_clean = str(provider).strip() if provider is not None else ""

        # RULE 1: Missing Provider -> NO SEND, NO FALLBACK
        if not provider_clean or provider_clean.lower() in ("none", "null", "nan"):
            error_msg = "Skipped: Missing/empty provider in Metabase (no fallback)"
            print(f"  Row {sheet_row_num} → ❌ SKIPPED — {error_msg}")

            sheet.update(range_name=f"F{sheet_row_num}", values=[["FAILED"]])
            sheet.update(range_name=f"H{sheet_row_num}:J{sheet_row_num}", values=[["", "", error_msg]])

            send_summary.append([sheet_row_num, phone, order_id, coupon_code, old_template or "Unknown", "FAILED", "-", error_msg])
            continue

        # RULE 2: Unknown Provider -> NO SEND, NO FALLBACK
        template = resolve_template(provider_clean)
        if template == "Unknown" or template.lower() not in TEMPLATE_BUILDERS:
            error_msg = f"Skipped: Unknown provider '{provider_clean}' (no fallback)"
            print(f"  Row {sheet_row_num} → Provider='{provider_clean}' → ❌ SKIPPED — {error_msg}")

            if old_template != "Unknown":
                sheet.update(range_name=f"G{sheet_row_num}", values=[["Unknown"]])

            sheet.update(range_name=f"F{sheet_row_num}", values=[["FAILED"]])
            sheet.update(range_name=f"H{sheet_row_num}:J{sheet_row_num}", values=[["", "", error_msg]])

            send_summary.append([sheet_row_num, phone, order_id, coupon_code, "Unknown", "FAILED", "-", error_msg])
            continue

        # Valid provider found: update col G if template changed
        print(f"  Row {sheet_row_num} → Provider='{provider_clean}' → Template='{template}'")
        if template != old_template:
            sheet.update(range_name=f"G{sheet_row_num}", values=[[template]])
            print(f"             Template updated: '{old_template}' → '{template}'")

        # Send WhatsApp
        builder = TEMPLATE_BUILDERS[template.lower()]
        payload = builder(phone, order_id, coupon_code)
        success, message_id, error = send_whatsapp(payload)

        if success:
            sent_at = now_str()
            sheet.update(range_name=f"F{sheet_row_num}", values=[["SENT"]])
            sheet.update(range_name=f"H{sheet_row_num}:J{sheet_row_num}", values=[[sent_at, message_id, ""]])

            send_summary.append([sheet_row_num, phone, order_id, coupon_code, template, "SENT", sent_at, message_id])
            print(f"  Row {sheet_row_num} → ✅ SENT  {phone}  {template}  {coupon_code}  {message_id}")
        else:
            sheet.update(range_name=f"F{sheet_row_num}", values=[["FAILED"]])
            sheet.update(range_name=f"H{sheet_row_num}:J{sheet_row_num}", values=[["", "", error]])

            send_summary.append([sheet_row_num, phone, order_id, coupon_code, template, "FAILED", "-", error])
            print(f"  Row {sheet_row_num} → ❌ FAILED  {phone}  {template}  {coupon_code}  {error}")

        time.sleep(SEND_DELAY_SECONDS)

    # ---------------------------------------------------------
    # SUMMARY
    # ---------------------------------------------------------
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

if __name__ == "__main__":
    main()
