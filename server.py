import json
import hashlib
import hmac
import math
import os
import re
import secrets
import sqlite3
import time
from datetime import datetime, timezone
from http.cookies import CookieError, SimpleCookie
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote, urlsplit
from urllib.request import Request, urlopen

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "3000"))
AUTH_SESSION_TTL = 8 * 60 * 60
PASSWORD_HASH_ITERATIONS = 600_000
MAX_LOGIN_ATTEMPTS = 5
LOGIN_WINDOW_SECONDS = 900
AUTH_SESSIONS = {}
LOGIN_ATTEMPTS = {}
ROOT = Path(__file__).resolve().parent
WORKBOOK_PATH = Path(os.environ.get("REGISTRATION_WORKBOOK", ROOT / "data" / "registrations.xlsx"))
ACCOUNT_DATABASE_PATH = Path(os.environ.get("ACCOUNT_DATABASE", ROOT / "data" / "accounts.sqlite3"))
PAGES = {
    "/": "index.html",
    "/index.html": "index.html",
    "/assistant.html": "assistant.html",
    "/startup-help.html": "startup-help.html",
    "/brand-login.html": "brand-login.html",
    "/startup-chat.html": "startup-chat.html",
}
MAX_BODY_SIZE = 16_000
REGISTRATION_HEADERS = (
    "Submitted At (UTC)",
    "Name",
    "Surname",
    "Phone Number",
    "Gmail",
    "Business Type",
    "Assistant Type",
    "Budget Amount",
    "Currency",
    "Budget Period",
)


def ask_gemini_ai(messages, system_instruction, api_key):
    model = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash").strip()
    url = (
        f"https://googleapis.com"
        f"{quote(model)}:generateContent?key={quote(api_key)}"
    )

    contents = []
    for msg in messages:
        role = "model" if msg.get("role") == "assistant" else "user"
        text = msg.get("content", "").strip()
        if text:
            contents.append({"role": role, "parts": [{"text": text}]})

    payload = {
        "systemInstruction": {"parts": [{"text": system_instruction}]},
        "contents": contents,
        "generationConfig": {
            "maxOutputTokens": 700,
            "temperature": 0.7,
        },
    }

    request = Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    with urlopen(request, timeout=45) as response:
        result = json.loads(response.read())

    candidates = result.get("candidates", [])
    if not candidates:
        return ""
    parts = candidates[0].get("content", {}).get("parts", [])
    return "\n".join(part.get("text", "") for part in parts if isinstance(part.get("text"), str)).strip()


def generate_basic_reply(messages, assistant_mode, business_type, budget_details):
    last_message = ""
    for msg in reversed(messages):
        if msg.get("role") == "user":
            last_message = msg.get("content", "").strip()
            break

    text = last_message.lower()
    biz_label = f" for your {business_type} business" if business_type else ""

    if assistant_mode == "budget" and isinstance(budget_details, dict):
        amount = budget_details.get("amount", 0)
        currency = budget_details.get("currency", "INR").upper()
        period = "monthly operating" if budget_details.get("period") == "monthly" else "total startup"
        ops = round(amount * 0.40, 2)
        mkt = round(amount * 0.25, 2)
        tools = round(amount * 0.20, 2)
        reserve = round(amount - (ops + mkt + tools), 2)
        return (
            f"Here is a practical {period} budget plan{biz_label} ({currency} {amount:,.2f}):\n\n"
            f"• Core Operations & Inventory (40%): {currency} {ops:,.2f}\n"
            f"• Marketing & Customer Acquisition (25%): {currency} {mkt:,.2f}\n"
            f"• Software, Website & Tools (20%): {currency} {tools:,.2f}\n"
            f"• Emergency Reserve (15%): {currency} {reserve:,.2f}\n\n"
            f"Tip: Keep fixed monthly costs separate from one-time setup expenses."
        )

    return (
        f"Here is a clear 4-step action plan{biz_label}:\n\n"
        f"1. Validate Demand: Talk to 10 potential customers and confirm they will pay for your solution.\n"
        f"2. Launch a Simple Offer: Start with one core product or service instead of building everything at once.\n"
        f"3. Set Up Online Presence: Create a clean landing page, WhatsApp Business profile, and Google Business listing.\n"
        f"4. Track Cash Flow Weekly: Record every expense and sale from day one."
    )


def valid_chat_request(payload):
    assistant_mode = payload.get("assistantMode")
    messages = payload.get("messages")
    business_type = payload.get("businessType")
    budget_details = payload.get("budgetDetails")
    if assistant_mode not in {"startup", "budget"}:
        return False
    if not isinstance(messages, list) or not 1 <= len(messages) <= 16:
        return False
    if not isinstance(business_type, str) or len(business_type) > 120:
        return False
    if assistant_mode == "budget":
        if not isinstance(budget_details, dict):
            return False
        if not isinstance(budget_details.get("amount"), (int, float)) or isinstance(budget_details["amount"], bool):
            return False
        if budget_details["amount"] <= 0:
            return False
        if not isinstance(budget_details.get("currency"), str) or len(budget_details["currency"]) != 3:
            return False
        if budget_details.get("period") not in {"startup", "monthly"}:
            return False

    return all(
        isinstance(message, dict)
        and message.get("role") in {"user", "assistant"}
        and isinstance(message.get("content"), str)
        and 0 < len(message["content"].strip()) <= 3000
        for message in messages
    )


def valid_registration_request(payload):
    text_limits = {
        "firstName": 100,
        "surname": 100,
        "phone": 40,
        "email": 254,
        "businessType": 120,
    }
    for field, max_length in text_limits.items():
        value = payload.get(field)
        if not isinstance(value, str) or not value.strip() or len(value.strip()) > max_length:
            return False

    password = payload.get("password")
    if not isinstance(password, str) or not 8 <= len(password) <= 128:
        return False

    if not re.fullmatch(r"[A-Za-z0-9._%+-]+@gmail\.com", payload["email"].strip(), re.IGNORECASE):
        return False
    if payload.get("assistantMode") not in {"startup", "budget"}:
        return False

    if payload["assistantMode"] == "budget":
        details = payload.get("budgetDetails")
        if not isinstance(details, dict):
            return False
        amount = details.get("amount")
        if not isinstance(amount, (int, float)) or isinstance(amount, bool):
            return False
        if not math.isfinite(amount) or amount <= 0:
            return False
        if not isinstance(details.get("currency"), str) or not re.fullmatch(r"[A-Z]{3}", details["currency"]):
            return False
        if details.get("period") not in {"startup", "monthly"}:
            return False

    return True


def spreadsheet_text(value):
    text = value.strip()
    if text.lstrip(" \t\r\n")[:1] in {"=", "+", "-", "@"}:
        return f"'{text}"
    return text


def append_registration(payload):
    WORKBOOK_PATH.parent.mkdir(parents=True, exist_ok=True)
    if WORKBOOK_PATH.exists():
        workbook = load_workbook(WORKBOOK_PATH)
        worksheet = workbook["Registrations"] if "Registrations" in workbook.sheetnames else workbook.active
        worksheet.title = "Registrations"
    else:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Registrations"

    try:
        if worksheet.max_row == 1 and worksheet["A1"].value is None:
            for column, header in enumerate(REGISTRATION_HEADERS, start=1):
                worksheet.cell(row=1, column=column, value=header)
            for cell in worksheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(fill_type="solid", fgColor="087E83")
            worksheet.freeze_panes = "A2"

        details = payload.get("budgetDetails") or {}
        worksheet.append((
            datetime.now(timezone.utc).isoformat(timespec="seconds"),
            spreadsheet_text(payload["firstName"]),
            spreadsheet_text(payload["surname"]),
            spreadsheet_text(payload["phone"]),
            spreadsheet_text(payload["email"]),
            spreadsheet_text(payload["businessType"]),
            "Budget planning" if payload["assistantMode"] == "budget" else "Startup help",
            details.get("amount"),
            details.get("currency", ""),
            details.get("period", ""),
        ))
        if details:
            worksheet.cell(row=worksheet.max_row, column=8).number_format = "#,##0.00"
        worksheet.auto_filter.ref = worksheet.dimensions
        for column, width in enumerate((24, 18, 18, 20, 30, 28, 20, 18, 14, 22), start=1):
            worksheet.column_dimensions[worksheet.cell(row=1, column=column).column_letter].width = width
        workbook.save(WORKBOOK_PATH)
    finally:
        workbook.close()


def open_account_database():
    ACCOUNT_DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(ACCOUNT_DATABASE_PATH)
    connection.execute(
        "CREATE TABLE IF NOT EXISTS accounts ("
        "email TEXT PRIMARY KEY COLLATE NOCASE, "
        "password_salt BLOB NOT NULL, "
        "password_hash BLOB NOT NULL, "
        "created_at TEXT NOT NULL)"
    )
    return connection


def create_account(payload):
    salt = secrets.token_bytes(16)
    password_hash = hashlib.pbkdf2_hmac(
        "sha256",
        payload["password"].encode("utf-8"),
        salt,
        PASSWORD_HASH_ITERATIONS,
    )
