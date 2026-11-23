import hashlib
import hmac
import json
import os
import re
import secrets
import sqlite3
import time
from datetime import datetime, timezone
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.error import HTTPError
from urllib.parse import quote, urlsplit
from urllib.request import Request, urlopen


from openpyxl import Workbook, load_workbook

HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "3000"))
ROOT = Path(__file__).resolve().parent
DB_PATH = ROOT / "data" / "accounts.sqlite3"
XLSX_PATH = ROOT / "data" / "registrations.xlsx"
SESSIONS = {}

PAGES = {
    "/": "index.html",
    "/index.html": "index.html",
    "/assistant.html": "assistant.html",
    "/startup-help.html": "startup-help.html",
    "/brand-login.html": "brand-login.html",
    "/startup-chat.html": "startup-chat.html",
}


def ask_gemini(messages, mode, biz_type, budget):
    api_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if not api_key:
        return "GEMINI_API_KEY is not set in Render Environment variables."

    prompt = "You are a practical small-business advisor. Give clear, concise, actionable steps."
    if biz_type:
        prompt += f" Business type: {biz_type}."
    if mode == "budget" and isinstance(budget, dict):
        prompt += f" Budget: {budget.get('currency', 'INR')} {budget.get('amount', 0)} ({budget.get('period', 'startup')})."

    contents = []
    for m in messages:
        role = "model" if m.get("role") == "assistant" else "user"
        text = str(m.get("content", "")).strip()
        if text:
            contents.append({"role": role, "parts": [{"text": text}]})

    body = {
        "systemInstruction": {"parts": [{"text": prompt}]},
        "contents": contents,
        "generationConfig": {"maxOutputTokens": 700, "temperature": 0.7},
    }

        host = "generativelanguage." + "googleapis.com"
    url = f"https://{host}/v1beta/models/gemini-2.5-flash:generateContent"
    headers = {
        "Content-Type": "application/json",
        "x-goog-api-key": api_key,
    }
    req = Request(url, data=json.dumps(body).encode("utf-8"), headers=headers, method="POST")

    try:
        with urlopen(req, timeout=45) as resp:
            data = json.loads(resp.read())
        candidates = data.get("candidates") or [{}]
        first_candidate = candidates.pop(0) if candidates else {}
        parts = first_candidate.get("content", {}).get("parts", [])
        reply = "\n".join(p.get("text", "") for p in parts).strip()
        return reply or "No reply generated. Please try again."
    except HTTPError as err:
        detail = err.read().decode("utf-8", errors="ignore")
        print("Gemini HTTP Error:", err.code, detail)
        return f"Gemini Error {err.code}: Check your GEMINI_API_KEY in Render."
    except Exception as err:
        print("Gemini Error:", repr(err))
        return f"Error: {err}"



def get_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("CREATE TABLE IF NOT EXISTS users (email TEXT PRIMARY KEY, salt BLOB, hash BLOB)")
    return conn


def save_user(data):
    salt = secrets.token_bytes(16)
    pw_hash = hashlib.pbkdf2_hmac("sha256", data["password"].encode("utf-8"), salt, 200_000)
    conn = get_db()
    try:
        conn.execute("INSERT INTO users (email, salt, hash) VALUES (?, ?, ?)", (data["email"].strip().lower(), salt, pw_hash))
        conn.commit()
    finally:
        conn.close()

    XLSX_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(XLSX_PATH) if XLSX_PATH.exists() else Workbook()
    ws = wb.active
    if ws.max_row == 1 and ws["A1"].value is None:
        ws.append(["Date", "Name", "Surname", "Phone", "Gmail", "Business", "Mode", "Amount", "Currency", "Period"])
    b = data.get("budgetDetails") or {}
    ws.append([
        datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M"),
        data.get("firstName", ""),
        data.get("surname", ""),
        data.get("phone", ""),
        data.get("email", ""),
        data.get("businessType", ""),
        data.get("assistantMode", ""),
        b.get("amount", ""),
        b.get("currency", ""),
        b.get("period", ""),
    ])
    wb.save(XLSX_PATH)
    wb.close()


def check_user(email, password):
    conn = get_db()
    try:
        row = conn.execute("SELECT salt, hash FROM users WHERE email = ?", (email,)).fetchone()
    finally:
        conn.close()
    if not row:
        return False
    test_hash = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), row[0], 200_000)
    return hmac.compare_digest(test_hash, row[1])


class Handler(BaseHTTPRequestHandler):
    def send_json(self, status, data, cookie=None):
        raw = json.dumps(data).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        if cookie:
            self.send_header("Set-Cookie", cookie)
        self.end_headers()
        self.wfile.write(raw)

    def logged_in(self):
        c = SimpleCookie(self.headers.get("Cookie", ""))
        sid = c["bizgrowth_auth"].value if "bizgrowth_auth" in c else ""
        return SESSIONS.get(sid, 0) > time.time()

    def do_GET(self):
        path = urlsplit(self.path).path
        if path == "/startup-chat.html" and not self.logged_in():
            self.send_response(303)
            self.send_header("Location", "/brand-login.html")
            self.end_headers()
            return
        file_name = PAGES.get(path)
        if not file_name or not (ROOT / file_name).exists():
            self.send_error(404, "Not Found")
            return
        html = (ROOT / file_name).read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(html)))
        self.end_headers()
        self.wfile.write(html)

    def do_POST(self):
        path = urlsplit(self.path).path
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > 16000:
            self.send_json(400, {"error": "Invalid request size."})
            return
        try:
            payload = json.loads(self.rfile.read(length))
        except Exception:
            self.send_json(400, {"error": "Invalid JSON."})
            return

        if path == "/api/register":
            email = str(payload.get("email", "")).strip().lower()
            pw = str(payload.get("password", ""))
            if not re.fullmatch(r"[a-z0-9._%+-]+@gmail\.com", email) or len(pw) < 8:
                self.send_json(400, {"error": "Use a valid @gmail.com address and 8+ character password."})
                return
            try:
                save_user(payload)
                self.send_json(201, {"saved": True})
            except sqlite3.IntegrityError:
                self.send_json(409, {"error": "Account already exists. Please sign in."})
            except Exception as err:
                print("Register Error:", repr(err))
                self.send_json(500, {"error": "Could not save registration."})
            return

        if path == "/api/auth/login":
            email = str(payload.get("email", "")).strip().lower()
            pw = str(payload.get("password", ""))
            if not check_user(email, pw):
                self.send_json(401, {"error": "Wrong Gmail or password."})
                return
            sid = secrets.token_urlsafe(24)
            SESSIONS[sid] = time.time() + 28800
            self.send_json(200, {"signedIn": True}, f"bizgrowth_auth={sid}; HttpOnly; Path=/; Max-Age=28800")
            return

        if path == "/api/chat":
            if not self.logged_in():
                self.send_json(401, {"error": "Please sign in first."})
                return
            reply = ask_gemini(
                payload.get("messages", []),
                payload.get("assistantMode", "startup"),
                str(payload.get("businessType", "")).strip(),
                payload.get("budgetDetails"),
            )
            self.send_json(200, {"reply": reply})
            return

        self.send_error(404, "Not Found")


if __name__ == "__main__":
    print(f"Running on http://{HOST}:{PORT}")
    HTTPServer((HOST, PORT), Handler).serve_forever()
