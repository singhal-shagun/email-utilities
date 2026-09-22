"""Shared helpers for send-email-via-smtp.

Non-secret defaults, .env loading, env parsing, and validation.
Kept separate so ``main.py`` only contains the email workflow
(``build_message`` / ``send_email`` / ``main``).
"""

from __future__ import annotations

import os
import re
import warnings
from pathlib import Path
from typing import Iterable


# Non-secret defaults only. Secrets/hosts/recipients must come from
# environment or .env / .env.local files (see .env.example).
# This module intentionally contains no real credentials so the repo
# can stay public.
DEFAULT_SMTP_PORT = 587
DEFAULT_USE_STARTTLS = True
DEFAULT_USE_SSL = False
DEFAULT_SMTP_TIMEOUT = 30
DEFAULT_SUBJECT = "Test email with attachments"
DEFAULT_BODY = "Hello,\n\nThis is a test email sent via SMTP with multiple attachments."

# Warn (rather than fail) when attachments exceed typical SMTP limits.
MAX_TOTAL_ATTACHMENT_BYTES = 25 * 1024 * 1024


def _as_list(value: Iterable[str] | str | os.PathLike[str] | None) -> list:
    """Normalize a single string/Path or an iterable into a list."""
    if value is None:
        return []
    if isinstance(value, (str, os.PathLike)):
        return [value]
    return list(value)


def _normalize_addresses(
    addresses: Iterable[str] | str | None, field_name: str
) -> list[str]:
    """Strip, drop empties, and validate email addresses."""
    raw_list = _as_list(addresses)
    normalized = [
        str(addr).strip() for addr in raw_list if str(addr).strip()
    ]
    for addr in normalized:
        if "@" not in addr or " " in addr:
            raise ValueError(f"Invalid email address in {field_name}: {addr!r}")
    return normalized


def _parse_dotenv_file(path: Path) -> dict[str, str]:
    """Minimal .env parser (no external dependency).

    Supports `KEY=value`, `export KEY=value`, single/double quotes,
    and `#` comments. Returns {} if the file does not exist.
    """
    values: dict[str, str] = {}
    if not path.is_file():
        return values
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export ") :].strip()
        if "=" not in line:
            continue
        key, _, raw = line.partition("=")
        key = key.strip()
        raw = raw.strip()
        if len(raw) >= 2 and raw[0] == raw[-1] and raw[0] in ("'", '"'):
            raw = raw[1:-1]
        # Inline comments only when value is unquoted.
        if raw and raw[0] not in ("'", '"') and "#" in raw:
            raw = raw.split("#", 1)[0].strip()
        if key:
            values[key] = raw
    return values


def _load_dotenv_files() -> list[str]:
    """Load `.env` then `.env.local` without overriding real env vars.

    Searches the package directory first, then the current working
    directory. `.env.local` wins over `.env`. Existing
    `os.environ` entries always win over files.
    """
    loaded: list[str] = []
    candidates: list[Path] = []
    for base in (Path(__file__).resolve().parent, Path.cwd()):
        candidates.append(base / ".env")
        candidates.append(base / ".env.local")
    for path in candidates:
        for key, value in _parse_dotenv_file(path).items():
            os.environ.setdefault(key, value)
        if path.is_file():
            loaded.append(str(path))
    return loaded


def _split_list(raw_value: str) -> list[str]:
    """Split on ';', ',', or newlines (portable across Windows/POSIX).

    Strips whitespace and surrounding quotes so
    'a.pdf; "b.pdf", c.pdf' works. ':' is intentionally not a
    separator to avoid breaking Windows drive letters (C:\\...).
    """
    parts = re.split(r"[;,\n]+", raw_value)
    cleaned = []
    for item in parts:
        item = item.strip().strip("\"'")
        if item:
            cleaned.append(item)
    return cleaned


def _read_attachments_from_env() -> list[str]:
    raw_value = os.getenv("EMAIL_ATTACHMENTS", "")
    if not raw_value.strip():
        return []
    return _split_list(raw_value)


def _read_address_list_from_env(name: str, default: list[str]) -> list[str]:
    raw_value = os.getenv(name)
    if raw_value is None:
        return list(default)
    return _split_list(raw_value)


def _read_recipients_from_env() -> list[str]:
    return _read_address_list_from_env("TO_ADDRESSES", [])


def _env_flag(name: str, default: bool) -> bool:
    raw_value = os.getenv(name)
    if raw_value is None:
        return default
    return raw_value.strip().lower() in {"1", "true", "yes", "on"}


def _parse_port(raw_value: str) -> int:
    try:
        port = int(raw_value.strip())
    except (AttributeError, ValueError):
        raise ValueError(f"Invalid SMTP_PORT: {raw_value!r}. Must be 1-65535.")
    if not 1 <= port <= 65535:
        raise ValueError(f"Invalid SMTP_PORT: {raw_value!r}. Must be 1-65535.")
    return port


def _parse_timeout(raw_value: str) -> int:
    try:
        timeout = int(raw_value.strip())
    except (AttributeError, ValueError):
        raise ValueError(f"Invalid SMTP_TIMEOUT: {raw_value!r}. Must be >= 1.")
    if timeout < 1:
        raise ValueError(f"Invalid SMTP_TIMEOUT: {raw_value!r}. Must be >= 1.")
    return timeout


def _env_tristate(name: str) -> bool | None:
    """Return True/False if set, None if unset or empty.

    Empty values (e.g. `EMAIL_USE_SSL=` from a template) are treated
    as unset so port-based auto-detection still works.
    """
    raw_value = os.getenv(name)
    if raw_value is None or not raw_value.strip():
        return None
    return raw_value.strip().lower() in {"1", "true", "yes", "on"}


def _resolve_security(smtp_port: int) -> tuple[bool, bool]:
    """Resolve (use_ssl, use_starttls) without surprising conflicts.

    Rules:
    - Explicit `true` + explicit `true` -> error with fix instructions.
    - Port 465 with no explicit SSL choice -> SSL on (standard SMTPS).
    - Port 465 + explicit STARTTLS=true but SSL unset -> prefer SSL,
      disable STARTTLS with a warning (the common copy-paste mistake
      from the port-587 example).
    - Otherwise STARTTLS defaults to on, SSL defaults to off.
    """
    ssl_opt = _env_tristate("EMAIL_USE_SSL")
    tls_opt = _env_tristate("EMAIL_USE_STARTTLS")

    if ssl_opt is True and tls_opt is True:
        raise ValueError(
            "EMAIL_USE_SSL and EMAIL_USE_STARTTLS are both enabled — "
            "use one, not both. For port 465 set EMAIL_USE_STARTTLS=false "
            "(implicit SSL). For port 587 set EMAIL_USE_SSL=false and "
            "EMAIL_USE_STARTTLS=true."
        )

    if ssl_opt is None and tls_opt is None:
        use_ssl = bool(DEFAULT_USE_SSL or smtp_port == 465)
        use_starttls = bool(DEFAULT_USE_STARTTLS and not use_ssl)
        return use_ssl, use_starttls

    if ssl_opt is True:
        # SSL explicitly on; STARTTLS unset means off.
        return True, False

    if ssl_opt is False:
        use_starttls = tls_opt if tls_opt is not None else DEFAULT_USE_STARTTLS
        return False, bool(use_starttls)

    # ssl_opt is None from here.
    if tls_opt is True:
        if smtp_port == 465:
            warnings.warn(
                "SMTP_PORT=465 implies implicit SSL, but "
                "EMAIL_USE_STARTTLS=true is set. Disabling STARTTLS and "
                "using SSL. Set EMAIL_USE_STARTTLS=false to silence this.",
                stacklevel=2,
            )
            return True, False
        return False, True

    # tls_opt is False here.
    use_ssl = bool(DEFAULT_USE_SSL or smtp_port == 465)
    return use_ssl, False


def _resolve_use_ssl(smtp_port: int) -> bool:
    use_ssl, _ = _resolve_security(smtp_port)
    return use_ssl


def _resolve_use_starttls(use_ssl: bool) -> bool:
    # Kept for backward compatibility; prefer _resolve_security().
    tls_opt = _env_tristate("EMAIL_USE_STARTTLS")
    if tls_opt is not None:
        return tls_opt
    # STARTTLS is the default for port 587, but must be off for SMTPS
    # unless explicitly requested (which then fails fast as exclusive).
    return bool(DEFAULT_USE_STARTTLS and not use_ssl)


def _require_configured(value: str | None, name: str, placeholders: set[str]) -> str:
    cleaned = (value or "").strip()
    if not cleaned or cleaned in placeholders or cleaned.startswith("your-"):
        raise ValueError(
            f"{name} is not configured (got {value!r}). "
            f"Set {name} in .env/.env.local or environment."
        )
    return cleaned


def render_template(template: str, mapping: dict[str, str]) -> str:
    """Render `[KEY]` and `{KEY}` placeholders from mapping.

    Example: with ``{"RECIPIENT_NAME": "Asha"}``, both
    ``"Dear [RECIPIENT_NAME],"`` and ``"Dear {RECIPIENT_NAME},"``
    become ``"Dear Asha,"``. Unknown placeholders are left as-is.
    """
    rendered = template
    for key, value in mapping.items():
        rendered = rendered.replace(f"[{key}]", str(value))
        rendered = rendered.replace("{" + key + "}", str(value))
    return rendered
