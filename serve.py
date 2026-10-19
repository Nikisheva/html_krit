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
import re
import secrets
import socket
import threading
import uuid
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse
from zoneinfo import ZoneInfo

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
STORE = DATA / "applications.jsonl"
MSK = ZoneInfo("Europe/Moscow")
DEADLINE = datetime(2026, 10, 21, 23, 59, 59, tzinfo=MSK)
HOST = os.environ.get("APPLY_HOST", "0.0.0.0")
PORT = int(os.environ.get("APPLY_PORT", "8765"))
MAX_BODY = 32_768
LOCK = threading.Lock()

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
TRAINING = {
    "yes": "Буду на обучении 20 октября",
    "no": "Не смогу быть на обучении",
    "done": "Уже прошёл обучение",
}
FIELDS = {
    "name": (2, 120),
    "unit": (2, 160),
    "position": (2, 160),
    "email": (5, 160),
    "contact": (3, 160),
    "process": (0, 500),
}

STATIC_TYPES = {
    ".html": "text/html; charset=utf-8",
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

    item["email"] = item["email"].lower()
    if not EMAIL_RE.match(item["email"]):
        return None, "Укажите почту в формате имя@домен."

    training = str(body.get("training") or "")
    if training not in TRAINING:
        return None, "Укажите, будете ли на обучении 20 октября."
    item["training"] = training

    for key in ("solo", "rules", "data_rules"):
        if body.get(key) is not True:
            return None, "Нужно подтвердить правила участия."
        item[key] = True
    return item, None


def upsert(item: dict) -> tuple[dict, bool]:
    stamp = now_msk().isoformat(timespec="seconds")
    with LOCK:
        rows = load_all()
        for row in rows:
            if row.get("email") == item["email"]:
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
    writer.writerow(
        [
            "id",
            "создана",
            "обновлена",
            "имя",
            "подразделение",
            "должность",
            "почта",
            "контакт",
            "обучение",
            "процесс",
        ]
    )
    for row in rows:
        writer.writerow(
            [
                row.get("id", ""),
                row.get("created_at", ""),
                row.get("updated_at", ""),
                row.get("name", ""),
                row.get("unit", ""),
                row.get("position", ""),
                row.get("email", ""),
                row.get("contact", ""),
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

    def end_json(self, code: int, payload: dict) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(raw)

    def is_local(self) -> bool:
        return self.client_address[0] in ("127.0.0.1", "::1")

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
                },
            )
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
        raw = file_path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", STATIC_TYPES[file_path.suffix.lower()])
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/api/applications":
            self.end_json(404, {"error": "Страница не найдена."})
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
        row, updated = upsert(item)
        self.end_json(
            200 if updated else 201,
            {"ok": True, "id": row["id"], "updated": updated},
        )


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
    print(f"Форма для сотрудников: http://127.0.0.1:{PORT}/apply.html", flush=True)
    ip = lan_ip()
    if ip and ip != "127.0.0.1":
        print(f"Ссылка в сети: http://{ip}:{PORT}/apply.html", flush=True)
    print(f"Список заявок: http://127.0.0.1:{PORT}/admin.html", flush=True)
    httpd.serve_forever()


if __name__ == "__main__":
    main()
