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
from urllib.parse import urlsplit
from urllib.request import Request, urlopen

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


HOST = "127.0.0.1"
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
    connection = open_account_database()
    try:
        connection.execute("BEGIN IMMEDIATE")
        connection.execute(
            "INSERT INTO accounts (email, password_salt, password_hash, created_at) VALUES (?, ?, ?, ?)",
            (
                payload["email"].strip().lower(),
                salt,
                password_hash,
                datetime.now(timezone.utc).isoformat(timespec="seconds"),
            ),
        )
        append_registration(payload)
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()


def verify_account(email, password):
    connection = open_account_database()
    try:
        account = connection.execute(
            "SELECT password_salt, password_hash FROM accounts WHERE email = ?",
            (email,),
        ).fetchone()
    finally:
        connection.close()

    salt = account[0] if account else bytes(16)
    expected_hash = account[1] if account else bytes(32)
    actual_hash = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt,
        PASSWORD_HASH_ITERATIONS,
    )
    return account is not None and hmac.compare_digest(actual_hash, expected_hash)


def login_attempt_key(client_ip, email):
    now = time.time()
    key = (client_ip, email)
    attempts = LOGIN_ATTEMPTS.get(key)
    if attempts is None or now - attempts[0] >= LOGIN_WINDOW_SECONDS:
        LOGIN_ATTEMPTS[key] = (now, 0)
    return key


def record_login_failure(key):
    now = time.time()
    started_at, count = LOGIN_ATTEMPTS.get(key, (now, 0))
    if now - started_at >= LOGIN_WINDOW_SECONDS:
        started_at, count = now, 0
    LOGIN_ATTEMPTS[key] = (started_at, count + 1)


def create_session(response):
    session_id = secrets.token_urlsafe(32)
    AUTH_SESSIONS[session_id] = time.time() + AUTH_SESSION_TTL
    cookie = f"bizgrowth_auth={session_id}; HttpOnly; SameSite=Lax; Path=/; Max-Age={AUTH_SESSION_TTL}"
    if os.environ.get("COOKIE_SECURE", "").lower() in {"1", "true", "yes"}:
        cookie += "; Secure"
    response.send_json(200, {"signedIn": True}, (("Set-Cookie", cookie),))


class RequestHandler(BaseHTTPRequestHandler):
    def send_json(self, status, payload, extra_headers=()):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        for name, value in extra_headers:
            self.send_header(name, value)
        self.end_headers()
        self.wfile.write(body)

    def is_authenticated(self):
        cookie = SimpleCookie()
        try:
            cookie.load(self.headers.get("Cookie", ""))
        except CookieError:
            return False

        session_cookie = cookie.get("bizgrowth_auth")
        if session_cookie is None:
            return False

        session_id = session_cookie.value
        expires_at = AUTH_SESSIONS.get(session_id, 0)
        if expires_at <= time.time():
            AUTH_SESSIONS.pop(session_id, None)
            return False
        return True

    def do_GET(self):
        route = urlsplit(self.path).path
        if route == "/startup-chat.html" and not self.is_authenticated():
            self.send_response(303)
            self.send_header("Location", "/brand-login.html")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            return

        page = PAGES.get(route)
        if page is None:
            self.send_error(404, "Not found")
            return

        try:
            content = (ROOT / page).read_bytes()
        except OSError:
            self.send_error(500, "Could not load page")
            return

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(content)

    def do_POST(self):
        route = urlsplit(self.path).path
        if route not in {"/api/chat", "/api/register", "/api/auth/login"}:
            self.send_error(404, "Not found")
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self.send_json(400, {"error": "Invalid request length."})
            return

        if not 0 < content_length <= MAX_BODY_SIZE:
            self.send_json(413, {"error": "Request body is missing or too large."})
            return

        try:
            payload = json.loads(self.rfile.read(content_length))
        except (json.JSONDecodeError, UnicodeDecodeError):
            self.send_json(400, {"error": "Request body must be valid JSON."})
            return

        if not isinstance(payload, dict):
            self.send_json(400, {"error": "Request body must be a JSON object."})
            return

        if route == "/api/auth/login":
            email = payload.get("email")
            password = payload.get("password")
            if not isinstance(email, str) or not isinstance(password, str) or not 1 <= len(password) <= 128:
                self.send_json(400, {"error": "Enter your Gmail address and password."})
                return
            email = email.strip().lower()
            attempt_key = login_attempt_key(self.client_address[0], email)
            if LOGIN_ATTEMPTS.get(attempt_key, (0, 0))[1] >= MAX_LOGIN_ATTEMPTS:
                self.send_json(
                    429,
                    {"error": "Too many sign-in attempts. Try again in 15 minutes."},
                    (("Retry-After", str(LOGIN_WINDOW_SECONDS)),),
                )
                return
            if not re.fullmatch(r"[A-Za-z0-9._%+-]+@gmail\.com", email, re.IGNORECASE):
                record_login_failure(attempt_key)
                self.send_json(401, {"error": "The Gmail address or password is incorrect."})
                return
            try:
                valid_account = verify_account(email, password)
            except sqlite3.Error:
                self.send_json(500, {"error": "Could not access the account database. Please try again."})
                return
            if not valid_account:
                record_login_failure(attempt_key)
                self.send_json(401, {"error": "The Gmail address or password is incorrect."})
                return
            LOGIN_ATTEMPTS.pop(attempt_key, None)
            create_session(self)
            return

        if route == "/api/register":
            if not valid_registration_request(payload):
                self.send_json(400, {"error": "Please check the registration details and try again."})
                return
            try:
                create_account(payload)
            except sqlite3.IntegrityError:
                self.send_json(409, {"error": "An account with this Gmail already exists. Sign in with your password."})
                return
            except (OSError, sqlite3.Error, ValueError):
                self.send_json(500, {"error": "Could not save the registration workbook. Close it in Excel and try again."})
                return
            self.send_json(201, {"saved": True})
            return

        if not self.is_authenticated():
            self.send_json(401, {"error": "Sign in to your BizGrowth Hub account before using the assistant."})
            return

        if not valid_chat_request(payload):
            self.send_json(400, {"error": "The chat request is invalid or too long."})
            return

        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            self.send_json(
                503,
                {"error": "Chat is not configured yet. Set OPENAI_API_KEY on the server and restart it."},
            )
            return

        business_type = payload["businessType"].strip()
        assistant_mode = payload["assistantMode"]
        budget_details = payload.get("budgetDetails")
        business_context = f" The user's business type is: {business_type}." if business_type else ""
        if assistant_mode == "budget":
            business_context += (
                f" The user wants budget planning. Their budget is {budget_details['currency'].upper()} "
                f"{budget_details['amount']} for their {'monthly operating budget' if budget_details['period'] == 'monthly' else 'total startup budget'}."
            )
        api_payload = {
            "model": os.environ.get("OPENAI_MODEL", "gpt-4.1-mini"),
            "instructions": (
                "You are a practical small-business advisor. Give clear, actionable guidance. "
                "When the user asks for budget planning, separate one-time and recurring costs, state assumptions, and show arithmetic clearly. "
                "Ask a focused follow-up question when important information is missing. "
                "Do not promise specific financial or legal outcomes."
                f"{business_context}"
            ),
            "input": payload["messages"],
            "max_output_tokens": 700,
            "store": False,
        }
        request = Request(
            "https://api.openai.com/v1/responses",
            data=json.dumps(api_payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )

        try:
            with urlopen(request, timeout=45) as api_response:
                result = json.loads(api_response.read())
        except (HTTPError, URLError, TimeoutError, json.JSONDecodeError, OSError):
            self.send_json(
                502,
                {"error": "OpenAI could not complete the request. Check the server configuration and try again."},
            )
            return

        reply = "\n".join(
            item["text"]
            for output in result.get("output", [])
            for item in output.get("content", [])
            if item.get("type") == "output_text" and isinstance(item.get("text"), str)
        ).strip()
        if not reply:
            self.send_json(502, {"error": "The chat service returned an empty reply. Please try again."})
            return

        self.send_json(200, {"reply": reply})


if __name__ == "__main__":
    server = HTTPServer((HOST, PORT), RequestHandler)
    print(f"BizGrowth Hub is running at http://{HOST}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()