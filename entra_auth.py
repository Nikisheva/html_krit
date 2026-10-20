"""Вход через корпоративный Microsoft Entra / AD krit.pro.

Пароль сотрудник вводит на странице Microsoft, не на форме заявки.
"""

from __future__ import annotations

import base64
import hashlib
import json
import os
import secrets
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TENANT_DEFAULT = "krit.pro"
GRAPH = "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle,department,mobilePhone,businessPhones"


def load_settings() -> dict[str, str]:
    values = {
        "AZURE_TENANT_ID": TENANT_DEFAULT,
        "AZURE_CLIENT_ID": "",
        "AZURE_CLIENT_SECRET": "",
    }
    env_path = ROOT / ".credentials.local"
    if env_path.is_file():
        for line in env_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, raw = line.split("=", 1)
            key = key.strip()
            if key in values:
                values[key] = raw.strip().strip('"').strip("'")
    for key in list(values):
        if os.environ.get(key):
            values[key] = os.environ[key].strip()
    return values


def enabled() -> bool:
    return bool(load_settings()["AZURE_CLIENT_ID"])


def pkce_pair() -> tuple[str, str]:
    verifier = secrets.token_urlsafe(64)
    digest = hashlib.sha256(verifier.encode("ascii")).digest()
    challenge = base64.urlsafe_b64encode(digest).rstrip(b"=").decode("ascii")
    return verifier, challenge


def authorize_url(origin: str, state: str, challenge: str) -> str:
    settings = load_settings()
    tenant = settings["AZURE_TENANT_ID"] or TENANT_DEFAULT
    query = urllib.parse.urlencode(
        {
            "client_id": settings["AZURE_CLIENT_ID"],
            "response_type": "code",
            "redirect_uri": f"{origin}/api/auth/callback",
            "response_mode": "query",
            "scope": "openid profile email User.Read",
            "state": state,
            "code_challenge": challenge,
            "code_challenge_method": "S256",
            "prompt": "select_account",
        }
    )
    return f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?{query}"


def _post(url: str, data: dict) -> dict:
    raw = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(url, data=raw, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    try:
        with urllib.request.urlopen(req, timeout=12) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        raise RuntimeError("Microsoft не подтвердил вход.") from exc


def _get(url: str, token: str) -> dict:
    req = urllib.request.Request(url)
    req.add_header("Authorization", "Bearer " + token)
    with urllib.request.urlopen(req, timeout=12) as resp:
        return json.loads(resp.read().decode("utf-8"))


def profile_from_code(origin: str, code: str, verifier: str) -> tuple[dict | None, str | None]:
    settings = load_settings()
    if not settings["AZURE_CLIENT_ID"]:
        return None, "Вход через Microsoft ещё не настроен."
    tenant = settings["AZURE_TENANT_ID"] or TENANT_DEFAULT
    payload = {
        "client_id": settings["AZURE_CLIENT_ID"],
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": f"{origin}/api/auth/callback",
        "code_verifier": verifier,
        "scope": "openid profile email User.Read",
    }
    if settings["AZURE_CLIENT_SECRET"]:
        payload["client_secret"] = settings["AZURE_CLIENT_SECRET"]
    try:
        tokens = _post(f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token", payload)
        graph = _get(GRAPH, tokens.get("access_token") or "")
    except Exception:
        return None, "Не удалось получить профиль из корпоративной учётки."

    email = str(graph.get("mail") or graph.get("userPrincipalName") or "").lower()
    if not email.endswith("@krit.pro"):
        return None, "Нужна корпоративная учётка @krit.pro."
    phones = graph.get("businessPhones") or []
    login = email.split("@", 1)[0]
    return {
        "login": login,
        "name": str(graph.get("displayName") or login),
        "email": email,
        "unit": str(graph.get("department") or ""),
        "position": str(graph.get("jobTitle") or ""),
        "contact": str(graph.get("mobilePhone") or (phones[0] if phones else "")),
    }, None
