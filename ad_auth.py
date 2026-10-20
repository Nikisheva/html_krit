"""Вход сотрудников через корпоративный AD krit.pro.

Пароль используется только для LDAP bind и нигде не сохраняется.
"""

from __future__ import annotations

import os
import re
import ssl
from pathlib import Path

ROOT = Path(__file__).resolve().parent
LOGIN_RE = re.compile(r"^[A-Za-z0-9._\\-]{2,64}$")


def load_settings() -> dict[str, str]:
    values = {
        "AD_DOMAIN": "krit.pro",
        "AD_BASE_DN": "DC=krit,DC=pro",
        "AD_SERVER": "",
        "AD_PORT": "636",
        "AD_USE_SSL": "1",
        "AD_STARTTLS": "1",
        "AD_TLS_VALIDATE": "0",
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


def normalize_login(raw: str, domain: str) -> tuple[str, str] | None:
    text = " ".join(str(raw or "").split())
    if "\\" in text:
        text = text.split("\\", 1)[1]
    if "@" in text:
        account, upn_domain = text.split("@", 1)
        if upn_domain.lower() != domain.lower():
            return None
        text = account
    if not LOGIN_RE.match(text) or "\\" in text:
        return None
    account = text
    return account, f"{account}@{domain}"


def discover_hosts(settings: dict[str, str]) -> list[str]:
    hosts: list[str] = []
    if settings["AD_SERVER"]:
        hosts.append(settings["AD_SERVER"])
    domain = settings["AD_DOMAIN"]
    try:
        import dns.resolver  # type: ignore
        answers = dns.resolver.resolve(f"_ldap._tcp.dc._msdcs.{domain}", "SRV")
        ranked = sorted(answers, key=lambda item: (item.priority, item.weight))
        for item in ranked:
            host = str(item.target).rstrip(".")
            if host not in hosts:
                hosts.append(host)
    except Exception:
        pass
    if domain not in hosts:
        hosts.append(domain)
    return hosts


def _tls(settings: dict[str, str]):
    from ldap3 import Tls

    validate = ssl.CERT_REQUIRED if settings["AD_TLS_VALIDATE"] == "1" else ssl.CERT_NONE
    return Tls(validate=validate)


def _first(attrs: dict, *names: str) -> str:
    for name in names:
        values = attrs.get(name) or []
        if values:
            text = " ".join(str(values[0]).split())
            if text:
                return text
    return ""


def authenticate(username: str, password: str) -> tuple[dict | None, str | None]:
    if not str(password or ""):
        return None, "Введите пароль домена."

    settings = load_settings()
    identity = normalize_login(username, settings["AD_DOMAIN"])
    if identity is None:
        return None, "Укажите корпоративный логин, например ivanov.ivan."
    account, upn = identity

    try:
        from ldap3 import ALL, Connection, Server, SIMPLE
        from ldap3.core.exceptions import LDAPBindError, LDAPException
        from ldap3.utils.conv import escape_filter_chars
    except ImportError:
        return None, "На сервере формы не установлена библиотека ldap3."

    query = (
        "(&(objectClass=user)(!(objectClass=computer))"
        f"(|(sAMAccountName={escape_filter_chars(account)})"
        f"(userPrincipalName={escape_filter_chars(upn)})))"
    )
    attributes = [
        "displayName",
        "cn",
        "mail",
        "userPrincipalName",
        "department",
        "title",
        "telephoneNumber",
        "mobile",
        "sAMAccountName",
    ]
    last_error = "Контроллер домена недоступен. Нужна корпоративная сеть или VPN."
    modes = []
    if settings["AD_USE_SSL"] == "1":
        modes.append((int(settings["AD_PORT"] or "636"), True, False))
    if settings["AD_STARTTLS"] == "1":
        modes.append((389, False, True))

    for host in discover_hosts(settings):
        for port, use_ssl, start_tls in modes:
            try:
                server = Server(
                    host,
                    port=port,
                    use_ssl=use_ssl,
                    tls=_tls(settings),
                    connect_timeout=5,
                    get_info=ALL,
                )
                conn = Connection(
                    server,
                    user=upn,
                    password=password,
                    authentication=SIMPLE,
                    auto_bind=False,
                    raise_exceptions=True,
                    receive_timeout=8,
                )
                conn.open()
                if start_tls:
                    conn.start_tls()
                conn.bind()
            except LDAPBindError:
                return None, "Неверный логин или пароль."
            except LDAPException:
                last_error = "Контроллер домена недоступен. Нужна корпоративная сеть или VPN."
                continue
            except Exception:
                last_error = "Контроллер домена недоступен. Нужна корпоративная сеть или VPN."
                continue

            try:
                conn.search(settings["AD_BASE_DN"], query, attributes=attributes, size_limit=1)
                if not conn.entries:
                    return None, "Учётная запись в AD найдена, но без карточки сотрудника."
                entry = conn.entries[0]
                attrs = {key: list(entry[key].values) for key in entry.entry_attributes}
                profile = {
                    "login": _first(attrs, "sAMAccountName") or account,
                    "name": _first(attrs, "displayName", "cn"),
                    "email": (_first(attrs, "mail", "userPrincipalName") or upn).lower(),
                    "unit": _first(attrs, "department"),
                    "position": _first(attrs, "title"),
                    "contact": _first(attrs, "telephoneNumber", "mobile"),
                }
                if not profile["name"]:
                    profile["name"] = profile["login"]
                return profile, None
            finally:
                try:
                    conn.unbind()
                except Exception:
                    pass

    return None, last_error
