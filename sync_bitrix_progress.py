#!/usr/bin/env python3
"""Сверка участников AI Challenge с Bitrix Stage.

Читает BITRIX_WEBHOOK_URL из .credentials.local. Значение вебхука в stdout не пишет.
"""

from __future__ import annotations

import json
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

import serve

ROOT = Path(__file__).resolve().parent
CHALLENGE_START = "2026-10-20"
REPORT = ROOT / "data" / "bitrix_progress.json"


def read_webhook() -> str:
    path = ROOT / ".credentials.local"
    if not path.is_file():
        return ""
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        if key.strip() == "BITRIX_WEBHOOK_URL":
            return value.strip().strip('"').strip("'").rstrip("/")
    return ""


def rest(base: str, method: str, payload: dict | None = None) -> dict:
    url = f"{base}/{method}.json"
    data = urllib.parse.urlencode(payload or {}, doseq=True).encode("utf-8")
    request = urllib.request.Request(url, data=data, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            body = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as err:
        raw = err.read().decode("utf-8", errors="replace")[:300]
        raise RuntimeError(f"{method} HTTP {err.code}") from None
    except urllib.error.URLError as err:
        raise RuntimeError(f"{method} network") from None
    if body.get("error"):
        raise RuntimeError(f"{method} {body.get('error')}")
    return body


def rest_all(base: str, method: str, payload: dict | None = None) -> list:
    start = 0
    rows: list = []
    while True:
        params = dict(payload or {})
        params["start"] = start
        body = rest(base, method, params)
        chunk = body.get("result") or []
        if isinstance(chunk, dict):
            chunk = chunk.get("types") or chunk.get("items") or list(chunk.values())
        if not isinstance(chunk, list):
            break
        rows.extend(chunk)
        next_start = body.get("next")
        if next_start is None:
            break
        start = int(next_start)
        if start <= 0:
            break
    return rows


def tokens(name: str) -> set[str]:
    return {part.casefold() for part in str(name or "").replace("ё", "е").split() if len(part) > 1}


def full_name(user: dict) -> str:
    parts = [user.get("LAST_NAME"), user.get("NAME"), user.get("SECOND_NAME")]
    return " ".join(part for part in parts if part).strip()


def user_active_since_start(user: dict) -> bool:
    for key in ("LAST_ACTIVITY_DATE", "LAST_LOGIN", "DATE_LAST_ACTIVITY"):
        value = str(user.get(key) or "")
        if value and value[:10] >= CHALLENGE_START:
            return True
    return str(user.get("IS_ONLINE") or "").upper() in {"Y", "1", "TRUE"}


def looks_admin(user: dict) -> bool | None:
    for key in ("IS_ADMIN", "ADMIN", "UF_ADMIN"):
        if key in user:
            return str(user.get(key) or "").upper() in {"Y", "1", "TRUE"}
    return None


def match_user(name: str, users: list[dict]) -> dict | None:
    need = tokens(name)
    if not need:
        return None
    scored = []
    for user in users:
        have = tokens(full_name(user))
        if not have:
            continue
        overlap = len(need & have)
        if overlap >= 2 or (overlap == 1 and len(need) == 1):
            scored.append((overlap, user))
    scored.sort(key=lambda item: item[0], reverse=True)
    return scored[0][1] if scored else None


def type_owner_id(spa: dict) -> str:
    for key in ("createdBy", "CREATED_BY", "created_by"):
        if spa.get(key) not in (None, "", 0, "0"):
            return str(spa.get(key))
    return ""


def type_title(spa: dict) -> str:
    return str(spa.get("title") or spa.get("TITLE") or spa.get("name") or "")


def type_done(spa: dict) -> bool:
    stages = spa.get("isStagesEnabled") or spa.get("isCategoriesEnabled")
    fields = spa.get("customFieldsCount") or spa.get("fieldsCount")
    if stages and int(fields or 0) >= 3:
        return True
    title = type_title(spa).casefold()
    return any(word in title for word in ("готов", "сделал", "готово", "заверш"))


def sync() -> dict:
    base = read_webhook()
    if not base:
        return {"ok": False, "error": "no-webhook"}
    user_payload = {"FILTER[ACTIVE]": "true"}
    for index, field in enumerate(
        ["ID", "NAME", "LAST_NAME", "SECOND_NAME", "LAST_ACTIVITY_DATE", "LAST_LOGIN", "IS_ONLINE", "ACTIVE"]
    ):
        user_payload[f"SELECT[{index}]"] = field
    users = rest_all(base, "user.get", user_payload)
    types = []
    try:
        types = rest_all(base, "crm.type.list", {})
    except RuntimeError:
        types = []
    rows = serve.load_all()
    report = []
    for row in rows:
        user = match_user(row.get("name") or "", users)
        owned = [spa for spa in types if user and type_owner_id(spa) == str(user.get("ID"))]
        entered = bool(user and user_active_since_start(user))
        started = bool(owned)
        done = any(type_done(spa) for spa in owned)
        admin = looks_admin(user) if user else None
        flags = {
            "bitrix_entered": entered,
            "project_started": started,
            "project_done": done,
        }
        if admin is not None:
            flags["bitrix_admin"] = admin
        if user:
            serve.update_progress(str(row.get("id")), flags)
        report.append(
            {
                "id": row.get("id"),
                "name": row.get("name"),
                "bitrix_user": full_name(user) if user else "",
                "bitrix_id": user.get("ID") if user else "",
                "entered": entered,
                "admin": admin,
                "started": started,
                "done": done,
                "processes": [type_title(spa) for spa in owned],
                "last_activity": (user or {}).get("LAST_ACTIVITY_DATE") or (user or {}).get("LAST_LOGIN") or "",
            }
        )
    REPORT.parent.mkdir(exist_ok=True)
    payload = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "users": len(users),
        "types": len(types),
        "items": report,
    }
    REPORT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"ok": True, "count": len(report), "users": len(users), "types": len(types), "items": report}


def print_table(result: dict) -> None:
    if not result.get("ok"):
        print("bitrix-sync", result.get("error"))
        return
    print(f"bitrix-sync users={result['users']} types={result['types']} participants={result['count']}")
    print("ФИО | Bitrix | Зашёл | Админ | Начал | Сделал | Процесс")
    for item in result["items"]:
        admin = "да" if item["admin"] else ("нет" if item["admin"] is False else "н/д")
        print(
            f"{item['name']} | {item['bitrix_user'] or 'нет'} | "
            f"{'да' if item['entered'] else 'нет'} | {admin} | "
            f"{'да' if item['started'] else 'нет'} | {'да' if item['done'] else 'нет'} | "
            f"{', '.join(item['processes']) or '—'}"
        )


if __name__ == "__main__":
    print_table(sync())
