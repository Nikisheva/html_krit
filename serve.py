#!/usr/bin/env python3
"""Сервер заявки на участие в AI Challenge.

Запуск из папки проекта: python3 serve.py
Форма: http://127.0.0.1:8765/apply.html
Список заявок (только с этого компьютера): http://127.0.0.1:8765/admin.html
"""

from __future__ import annotations

import csv
import io
import json
import os
import secrets
import socket
import threading
import uuid
from datetime import datetime, timedelta
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse
from zoneinfo import ZoneInfo

from ad_auth import authenticate
from entra_auth import authorize_url, enabled as entra_enabled, pkce_pair, profile_from_code

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
STORE = DATA / "applications.jsonl"
MSK = ZoneInfo("Europe/Moscow")
DEADLINE = datetime(2026, 10, 21, 23, 59, 59, tzinfo=MSK)
HOST = os.environ.get("APPLY_HOST", "0.0.0.0")
PORT = int(os.environ.get("APPLY_PORT", "8765"))
MAX_BODY = 32_768
LOCK = threading.Lock()
SESSIONS: dict[str, dict] = {}
OAUTH: dict[str, dict] = {}
LOGIN_HITS: dict[str, list[datetime]] = {}
SESSION_HOURS = 8
COOKIE = "apply_session"

TRAINING = {
    "passed": "Прошёл",
    "missed": "Не прошёл",
}
FIELDS = {
    "name": (2, 120),
    "position": (2, 160),
    "process": (0, 500),
}

STATIC_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
    ".gif": "image/gif",
    ".ttf": "font/ttf",
    ".woff": "font/woff",
    ".woff2": "font/woff2",
    ".pdf": "application/pdf",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}


def now_msk() -> datetime:
    return datetime.now(MSK)


def admin_token() -> str:
    token = os.environ.get("APPLY_ADMIN_TOKEN", "").strip()
    if token:
        return token
    path = ROOT / ".credentials.local"
    if not path.is_file():
        return ""
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        if key.strip() == "APPLY_ADMIN_TOKEN":
            return value.strip().strip('"').strip("'")
    return ""


def load_all() -> list[dict]:
    if not STORE.is_file():
        return []
    rows = []
    for line in STORE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        rows.append(json.loads(line))
    return rows


def save_all(rows: list[dict]) -> None:
    DATA.mkdir(exist_ok=True)
    payload = "".join(json.dumps(row, ensure_ascii=False) + "\n" for row in rows)
    tmp = STORE.with_suffix(".jsonl.tmp")
    tmp.write_text(payload, encoding="utf-8")
    tmp.replace(STORE)


def clean_text(value: object, limit: int) -> str:
    text = " ".join(str(value or "").split())
    return text[:limit]


def validate(body: dict) -> tuple[dict | None, str | None]:
    if now_msk() > DEADLINE:
        return None, "Приём заявок закрыт. Срок — 21 октября 2026, 23:59 мск."
    if str(body.get("company_website") or "").strip():
        return {"honeypot": True}, None

    item = {}
    for key, (low, high) in FIELDS.items():
        value = clean_text(body.get(key), high)
        if len(value) < low:
            return None, "Заполните все обязательные поля."
        item[key] = value

    training = str(body.get("training") or "")
    if training not in TRAINING:
        return None, "Укажите, прошли обучение или нет."
    item["training"] = training
    return item, None


def upsert(item: dict) -> tuple[dict, bool]:
    stamp = now_msk().isoformat(timespec="seconds")
    with LOCK:
        rows = load_all()
        for row in rows:
            same_login = item.get("ad_login") and row.get("ad_login") == item["ad_login"]
            same_name = item.get("name") and (row.get("name") or "").casefold() == item["name"].casefold()
            if same_login or same_name:
                row.update(item)
                row["updated_at"] = stamp
                save_all(rows)
                return row, True
        row = {
            "id": uuid.uuid4().hex[:8],
            "created_at": stamp,
            "updated_at": stamp,
            **item,
        }
        rows.append(row)
        save_all(rows)
        return row, False


def rows_csv(rows: list[dict]) -> bytes:
    buffer = io.StringIO()
    writer = csv.writer(buffer)
    writer.writerow(["id", "создана", "обновлена", "имя", "должность", "обучение", "процесс"])
    for row in rows:
        writer.writerow(
            [
                row.get("id", ""),
                row.get("created_at", ""),
                row.get("updated_at", ""),
                row.get("name", ""),
                row.get("position", ""),
                TRAINING.get(row.get("training", ""), ""),
                row.get("process", ""),
            ]
        )
    return buffer.getvalue().encode("utf-8-sig")


def safe_file(url_path: str) -> Path | None:
    rel = unquote(url_path).lstrip("/")
    if rel in ("",):
        rel = "apply.html"
    candidate = (ROOT / rel).resolve()
    if not str(candidate).startswith(str(ROOT)):
        return None
    if any(part.startswith(".") for part in candidate.relative_to(ROOT).parts):
        return None
    if candidate == DATA or DATA in candidate.parents:
        return None
    if not candidate.is_file():
        return None
    if candidate.suffix.lower() not in STATIC_TYPES:
        return None
    return candidate


class Handler(BaseHTTPRequestHandler):
    server_version = "AIChallengeApply/1.0"

    def log_message(self, fmt: str, *args) -> None:
        print("%s - %s" % (self.address_string(), fmt % args))

    def add_cors(self) -> None:
        origin = self.headers.get("Origin", "")
        if origin.startswith("https://nikisheva.github.io"):
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type")
            self.send_header("Vary", "Origin")

    def end_json(self, code: int, payload: dict) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        self.add_cors()
        self.end_headers()
        self.wfile.write(raw)

    def do_OPTIONS(self) -> None:
        path = urlparse(self.path).path
        if path not in ("/api/health", "/api/applications", "/api/login", "/api/me", "/api/logout"):
            self.end_json(404, {"error": "Страница не найдена."})
            return
        self.send_response(204)
        self.add_cors()
        self.send_header("Content-Length", "0")
        self.end_headers()

    def request_origin(self) -> str:
        host = self.headers.get("Host", f"127.0.0.1:{PORT}")
        proto = "https" if host.endswith("trycloudflare.com") or self.headers.get("X-Forwarded-Proto") == "https" else "http"
        return f"{proto}://{host}"

    def redirect(self, location: str) -> None:
        self.send_response(302)
        self.send_header("Location", location)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def is_local(self) -> bool:
        return self.client_address[0] in ("127.0.0.1", "::1")

    def cookie_token(self) -> str:
        jar = SimpleCookie()
        jar.load(self.headers.get("Cookie", ""))
        morsel = jar.get(COOKIE)
        return morsel.value if morsel else ""

    def current_user(self) -> dict | None:
        prune_sessions()
        item = SESSIONS.get(self.cookie_token())
        if not item:
            return None
        return item["user"]

    def set_session_cookie(self, token: str | None) -> None:
        secure = " Secure;" if self.headers.get("X-Forwarded-Proto") == "https" or self.headers.get("Host", "").endswith("trycloudflare.com") else ""
        if token:
            self.send_header(
                "Set-Cookie",
                f"{COOKIE}={token}; HttpOnly; Path=/; SameSite=Lax; Max-Age={SESSION_HOURS * 3600};{secure}",
            )
        else:
            self.send_header(
                "Set-Cookie",
                f"{COOKIE}=; HttpOnly; Path=/; SameSite=Lax; Max-Age=0;{secure}",
            )

    def is_admin(self) -> bool:
        if self.is_local():
            return True
        expected = admin_token()
        given = self.headers.get("X-Admin-Token", "")
        return bool(expected) and secrets.compare_digest(given, expected)

    def read_json(self) -> dict | None:
        length = int(self.headers.get("Content-Length", "0") or 0)
        if length <= 0 or length > MAX_BODY:
            return None
        raw = self.rfile.read(length)
        try:
            body = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            return None
        return body if isinstance(body, dict) else None

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/api/health":
            self.end_json(
                200,
                {
                    "ok": True,
                    "open": now_msk() <= DEADLINE,
                    "deadline": DEADLINE.isoformat(timespec="seconds"),
                    "auth": "entra" if entra_enabled() else "none",
                },
            )
            return
        if path == "/api/auth/microsoft":
            if not entra_enabled():
                self.redirect("/login.html?error=" + quote("Вход Microsoft ещё не настроен. Нужен Client ID приложения в Entra."))
                return
            state = secrets.token_urlsafe(24)
            verifier, challenge = pkce_pair()
            OAUTH[state] = {
                "verifier": verifier,
                "origin": self.request_origin(),
                "expires": now_msk() + timedelta(minutes=10),
            }
            self.redirect(authorize_url(self.request_origin(), state, challenge))
            return
        if path == "/api/auth/callback":
            query = parse_qs(urlparse(self.path).query)
            state = (query.get("state") or [""])[0]
            code = (query.get("code") or [""])[0]
            pending = OAUTH.pop(state, None)
            if not pending or pending["expires"] <= now_msk() or not code:
                self.redirect("/login.html?error=" + quote("Вход не завершён. Попробуйте ещё раз."))
                return
            profile, error = profile_from_code(pending["origin"], code, pending["verifier"])
            if error or profile is None:
                self.redirect("/login.html?error=" + quote(error or "Не удалось войти."))
                return
            token = create_session(profile)
            self.send_response(302)
            self.set_session_cookie(token)
            self.send_header("Location", "/apply.html")
            self.send_header("Content-Length", "0")
            self.end_headers()
            return
        if path == "/api/me":
            user = self.current_user()
            if not user:
                self.end_json(401, {"error": "Нужен вход через AD."})
                return
            self.end_json(200, user)
            return
        if path in ("/api/applications", "/api/applications.csv"):
            if not self.is_admin():
                self.end_json(403, {"error": "Список заявок доступен только организатору."})
                return
            rows = load_all()
            rows.sort(key=lambda row: row.get("created_at", ""))
            if path.endswith(".csv"):
                raw = rows_csv(rows)
                self.send_response(200)
                self.send_header("Content-Type", "text/csv; charset=utf-8")
                self.send_header(
                    "Content-Disposition",
                    "attachment; filename=ai-challenge-applications.csv",
                )
                self.send_header("Content-Length", str(len(raw)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(raw)
                return
            self.end_json(200, {"items": rows, "count": len(rows)})
            return

        file_path = safe_file(path)
        if file_path is None:
            self.end_json(404, {"error": "Страница не найдена."})
            return
        if file_path.name == "admin.html" and not self.is_local():
            self.end_json(403, {"error": "Список заявок доступен только организатору."})
            return
        raw = file_path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", STATIC_TYPES[file_path.suffix.lower()])
        if file_path.suffix.lower() == ".docx":
            self.send_header(
                "Content-Disposition",
                "attachment; filename=\"Cursor-backoffice-KRIT.docx\"",
            )
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path == "/api/logout":
            token = self.cookie_token()
            SESSIONS.pop(token, None)
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.set_session_cookie(None)
            raw = json.dumps({"ok": True}, ensure_ascii=False).encode("utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return
        if path == "/api/login":
            if not self.is_local():
                self.end_json(403, {"error": "Пароль AD на этой ссылке не принимается. Войдите через Microsoft."})
                return
            ip = self.client_address[0]
            if not login_allowed(ip):
                self.end_json(429, {"error": "Слишком много попыток входа. Подождите 15 минут."})
                return
            body = self.read_json()
            if body is None:
                self.end_json(400, {"error": "Не удалось прочитать данные входа."})
                return
            mark_login(ip)
            profile, error = authenticate(str(body.get("username") or ""), str(body.get("password") or ""))
            if error or profile is None:
                self.end_json(401, {"error": error or "Неверный логин или пароль."})
                return
            token = create_session(profile)
            raw = json.dumps({"ok": True, "user": public_user(profile)}, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.send_header("Cache-Control", "no-store")
            self.set_session_cookie(token)
            self.end_headers()
            self.wfile.write(raw)
            return
        if path != "/api/applications":
            self.end_json(404, {"error": "Страница не найдена."})
            return
        user = self.current_user()
        if entra_enabled() and not user:
            self.end_json(401, {"error": "Сначала войдите корпоративной учёткой."})
            return
        body = self.read_json()
        if body is None:
            self.end_json(400, {"error": "Не удалось прочитать заявку."})
            return
        item, error = validate(body)
        if error:
            code = 403 if "закрыт" in error else 400
            self.end_json(code, {"error": error})
            return
        if item.get("honeypot"):
            self.end_json(201, {"ok": True, "id": "00000000", "updated": False})
            return
        if user:
            item["ad_login"] = user["login"]
            item["name"] = user["name"] or item["name"]
            item["position"] = user["position"] or item["position"]
        row, updated = upsert(item)
        self.end_json(
            200 if updated else 201,
            {"ok": True, "id": row["id"], "updated": updated},
        )


def public_user(profile: dict) -> dict:
    return {
        "login": profile.get("login", ""),
        "name": profile.get("name", ""),
        "email": profile.get("email", ""),
        "unit": profile.get("unit", ""),
        "position": profile.get("position", ""),
    }


def prune_sessions() -> None:
    now = now_msk()
    expired = [key for key, item in SESSIONS.items() if item["expires"] <= now]
    for key in expired:
        SESSIONS.pop(key, None)


def create_session(profile: dict) -> str:
    prune_sessions()
    token = secrets.token_urlsafe(32)
    SESSIONS[token] = {
        "user": public_user(profile),
        "expires": now_msk() + timedelta(hours=SESSION_HOURS),
    }
    return token


def login_allowed(ip: str) -> bool:
    now = now_msk()
    window = now - timedelta(minutes=15)
    hits = [stamp for stamp in LOGIN_HITS.get(ip, []) if stamp > window]
    LOGIN_HITS[ip] = hits
    return len(hits) < 8


def mark_login(ip: str) -> None:
    LOGIN_HITS.setdefault(ip, []).append(now_msk())


def lan_ip() -> str:
    try:
        probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        probe.connect(("8.8.8.8", 80))
        ip = probe.getsockname()[0]
        probe.close()
        return ip
    except OSError:
        return ""


def main() -> None:
    DATA.mkdir(exist_ok=True)
    httpd = ThreadingHTTPServer((HOST, PORT), Handler)
    print("Форма для сотрудников: https://nikisheva.github.io/html_krit/apply.html", flush=True)
    print(f"Локально: http://127.0.0.1:{PORT}/apply.html", flush=True)
    ip = lan_ip()
    if ip and ip != "127.0.0.1":
        print(f"В этой сети: http://{ip}:{PORT}/apply.html", flush=True)
    print(f"Список заявок: http://127.0.0.1:{PORT}/admin.html", flush=True)
    httpd.serve_forever()


if __name__ == "__main__":
    main()
