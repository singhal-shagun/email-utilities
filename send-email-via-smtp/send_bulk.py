"""Send personalized bulk emails driven by an Excel file.

Each row = one recipient. Expected columns (case-insensitive):

    RECIPIENT_NAME, RECIPIENT_EMAIL_ADDRESS,
    ATTACHMENT1_FILE_PATH, ATTACHMENT2_FILE_PATH, ... (any number of
    columns whose header starts with "ATTACHMENT")

Attachment cells hold file names; several names per cell may be
separated with ``;`` or ``,``. Relative names are resolved against
``--attachments-dir`` (default: the Excel file's folder).

Usage:
    pip install -r requirements.txt
    python send_bulk.py --excel guests.xlsx --dry-run
    python send_bulk.py --excel guests.xlsx

The body/subject are templates where ``[RECIPIENT_NAME]`` (or
``{RECIPIENT_NAME}``) is replaced per row, e.g.::

    Dear [RECIPIENT_NAME],

    We are glad to invite you to our family dinner.

    Regards,
    XYZ

Config (flags win over env, env wins over .env files):
    SMTP_* / FROM_ADDRESS like main.py, plus
    BULK_EXCEL, BULK_SUBJECT (or EMAIL_SUBJECT), BULK_BODY,
    BULK_BODY_FILE, BULK_BODY_HTML_FILE, BULK_ATTACHMENTS_DIR,
    BULK_DELAY_SECONDS, BULK_DRY_RUN, BULK_CC (or EMAIL_CC),
    BULK_BCC (or EMAIL_BCC).

    CC/BCC addresses are added to every email. Note: a CC'd address
    will therefore receive one copy per recipient.
"""

from __future__ import annotations

import argparse
import os
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

from helpers_smtp import (
    DEFAULT_SMTP_PORT,
    DEFAULT_SMTP_TIMEOUT,
    _load_dotenv_files,
    _normalize_addresses,
    _parse_port,
    _parse_timeout,
    _read_address_list_from_env,
    _require_configured,
    _resolve_security,
    _split_list,
    render_template,
)
from main import send_email

DEFAULT_BODY_TEMPLATE = """Dear [RECIPIENT_NAME],

We are glad to invite you to our family dinner.

Regards,
XYZ"""

REQUIRED_COLUMNS = ("RECIPIENT_NAME", "RECIPIENT_EMAIL_ADDRESS")


@dataclass
class Recipient:
    name: str
    email: str
    attachments: list[str] = field(default_factory=list)
    row: int = 0


def _require_openpyxl():
    try:
        import openpyxl  # type: ignore
    except ImportError:
        raise RuntimeError(
            "openpyxl is required for Excel bulk sends. "
            "Run: pip install -r requirements.txt"
        )
    return openpyxl


def load_recipients(excel_path: Path) -> list[Recipient]:
    """Read and validate the Excel roster. Fails fast with row numbers."""
    openpyxl = _require_openpyxl()
    if not excel_path.is_file():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    sheet = workbook.active
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f"Excel file is empty: {excel_path}")

    headers = [
        (str(cell).strip().upper() if cell is not None else "")
        for cell in rows[0]
    ]
    col_index = {header: i for i, header in enumerate(headers) if header}
    missing = [c for c in REQUIRED_COLUMNS if c not in col_index]
    if missing:
        raise ValueError(
            f"Missing required column(s) {missing} in {excel_path}. "
            f"Found headers: {headers}"
        )
    attachment_cols = [
        i for i, header in enumerate(headers) if header.startswith("ATTACHMENT")
    ]
    if not attachment_cols:
        print("Warning: no ATTACHMENT* columns found; sending without attachments.")

    recipients: list[Recipient] = []
    errors: list[str] = []
    for lineno, row in enumerate(rows[1:], start=2):
        if all(cell is None or str(cell).strip() == "" for cell in row):
            continue  # skip blank rows
        name = str(row[col_index["RECIPIENT_NAME"]]).strip() if row[col_index["RECIPIENT_NAME"]] is not None else ""
        email = str(row[col_index["RECIPIENT_EMAIL_ADDRESS"]]).strip() if row[col_index["RECIPIENT_EMAIL_ADDRESS"]] is not None else ""
        if not name:
            errors.append(f"row {lineno}: RECIPIENT_NAME is empty")
            continue
        try:
            _normalize_addresses([email], f"row {lineno}")
        except ValueError:
            errors.append(f"row {lineno}: invalid RECIPIENT_EMAIL_ADDRESS: {email!r}")
            continue
        attachments: list[str] = []
        for col in attachment_cols:
            if col < len(row) and row[col] is not None and str(row[col]).strip():
                attachments.extend(_split_list(str(row[col])))
        recipients.append(Recipient(name=name, email=email, attachments=attachments, row=lineno))

    if errors:
        raise ValueError("Invalid rows in Excel file:\n- " + "\n- ".join(errors))
    if not recipients:
        raise ValueError(f"No recipients found in {excel_path}.")
    return recipients


def resolve_attachments(
    recipients: list[Recipient], attachments_dir: Path
) -> None:
    """Resolve attachment names to existing files. Fails fast if missing."""
    missing: list[str] = []
    for recipient in recipients:
        resolved: list[str] = []
        for raw in recipient.attachments:
            name = raw.strip().strip("\"'")
            if not name:
                continue
            candidate = Path(name)
            if not candidate.is_absolute() and not candidate.exists():
                candidate = attachments_dir / name
            if not candidate.is_file():
                missing.append(f"row {recipient.row} ({recipient.email}): {name!r}")
            else:
                resolved.append(str(candidate))
        recipient.attachments = resolved
    if missing:
        raise FileNotFoundError(
            "Attachment(s) not found (set --attachments-dir if needed):\n- "
            + "\n- ".join(missing)
        )


def _read_text_file(path: str | None) -> str | None:
    if not path:
        return None
    return Path(path).read_text(encoding="utf-8")


def send_bulk(
    excel: Path,
    subject_template: str,
    body_template: str,
    smtp_host: str,
    smtp_port: int,
    username: str,
    password: str,
    from_address: str,
    attachments_dir: Path | None = None,
    body_html_template: str | None = None,
    dry_run: bool = False,
    delay_seconds: float = 0.0,
    use_starttls: bool = True,
    use_ssl: bool = False,
    timeout: int = DEFAULT_SMTP_TIMEOUT,
    cc_addresses: list[str] | None = None,
    bcc_addresses: list[str] | None = None,
) -> tuple[int, int]:
    recipients = load_recipients(excel)
    base_dir = attachments_dir or excel.parent
    resolve_attachments(recipients, base_dir)
    cc_list = _normalize_addresses(cc_addresses or [], "cc_addresses")
    bcc_list = _normalize_addresses(bcc_addresses or [], "bcc_addresses")

    sent = failed = 0
    for i, recipient in enumerate(recipients):
        mapping = {
            "RECIPIENT_NAME": recipient.name,
            "RECIPIENT_EMAIL_ADDRESS": recipient.email,
        }
        subject = render_template(subject_template, mapping)
        body = render_template(body_template, mapping)
        body_html = (
            render_template(body_html_template, mapping)
            if body_html_template
            else None
        )
        label = f"[{i + 1}/{len(recipients)}] row {recipient.row}: {recipient.email}"
        if dry_run:
            print(f"DRY-RUN {label} name={recipient.name!r} "
                  f"attachments={len(recipient.attachments)} "
                  f"cc={cc_list} bcc={bcc_list}")
            sent += 1
            continue
        try:
            send_email(
                smtp_host=smtp_host,
                smtp_port=smtp_port,
                username=username,
                password=password,
                from_address=from_address,
                to_addresses=[recipient.email],
                subject=subject,
                body=body,
                attachments=recipient.attachments,
                cc_addresses=cc_list,
                bcc_addresses=bcc_list,
                use_starttls=use_starttls,
                use_ssl=use_ssl,
                timeout=timeout,
                body_html=body_html,
            )
            print(f"SENT {label}")
            sent += 1
        except Exception as exc:  # noqa: BLE001 — report per-row, keep going
            print(f"FAILED {label}: {exc}", file=sys.stderr)
            failed += 1
        if delay_seconds and i < len(recipients) - 1:
            time.sleep(delay_seconds)
    print(f"Done: {sent} sent, {failed} failed, {len(recipients)} total.")
    return sent, failed


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--excel", default=os.getenv("BULK_EXCEL", ""),
                        help="Excel roster path (or BULK_EXCEL).")
    parser.add_argument("--subject", default=os.getenv("BULK_SUBJECT", os.getenv("EMAIL_SUBJECT", "You are invited!")),
                        help="Subject template (supports [RECIPIENT_NAME]).")
    parser.add_argument("--body", default=os.getenv("BULK_BODY", ""),
                        help="Body template string (or use --body-file).")
    parser.add_argument("--body-file", default=os.getenv("BULK_BODY_FILE", ""),
                        help="File with the plain-text body template.")
    parser.add_argument("--body-html-file", default=os.getenv("BULK_BODY_HTML_FILE", ""),
                        help="Optional HTML body template file.")
    parser.add_argument("--attachments-dir", default=os.getenv("BULK_ATTACHMENTS_DIR", ""),
                        help="Base dir for relative attachment names.")
    parser.add_argument("--cc", default=os.getenv("BULK_CC", os.getenv("EMAIL_CC", "")),
                        help="CC address(es) added to every email (; or , separated).")
    parser.add_argument("--bcc", default=os.getenv("BULK_BCC", os.getenv("EMAIL_BCC", "")),
                        help="BCC address(es) added to every email (; or , separated).")
    parser.add_argument("--dry-run", action="store_true",
                        default=os.getenv("BULK_DRY_RUN", "").strip().lower() in {"1", "true", "yes", "on"},
                        help="Validate + print plan without sending.")
    parser.add_argument("--delay", type=float,
                        default=float(os.getenv("BULK_DELAY_SECONDS", "0") or 0),
                        help="Seconds to wait between sends.")
    return parser.parse_args(argv)


def main_bulk(argv: list[str] | None = None) -> None:
    _load_dotenv_files()
    args = parse_args(argv)

    if not args.excel:
        raise ValueError("Excel file required: --excel guests.xlsx or BULK_EXCEL.")
    excel = Path(args.excel)

    if args.body_file:
        body_template = Path(args.body_file).read_text(encoding="utf-8")
    elif args.body:
        # Allow literal \n in env-provided bodies.
        body_template = args.body.replace("\\n", "\n")
    else:
        body_template = os.getenv("EMAIL_BODY", DEFAULT_BODY_TEMPLATE)

    body_html_template = _read_text_file(args.body_html_file or None)

    smtp_port = _parse_port(os.getenv("SMTP_PORT", str(DEFAULT_SMTP_PORT)))
    use_ssl, use_starttls = _resolve_security(smtp_port)
    timeout = _parse_timeout(os.getenv("SMTP_TIMEOUT", str(DEFAULT_SMTP_TIMEOUT)))
    smtp_host = _require_configured(
        os.getenv("SMTP_HOST", ""), "SMTP_HOST", {"smtp.example.com"}
    )
    from_address = _require_configured(
        os.getenv("FROM_ADDRESS", ""), "FROM_ADDRESS", {"your-email@example.com"}
    )
    username = os.getenv("SMTP_USERNAME", "")
    password = os.getenv("SMTP_PASSWORD", "")
    # CC/BCC come from --cc/--bcc flags, which default to BULK_CC/BULK_BCC
    # or EMAIL_CC/EMAIL_BCC — so an EMAIL_CC already in .env.local is
    # picked up automatically, just like in main.py.
    cc_addresses = _read_address_list_from_env("EMAIL_CC", []) if not args.cc else _split_list(args.cc)
    bcc_addresses = _read_address_list_from_env("EMAIL_BCC", []) if not args.bcc else _split_list(args.bcc)

    sent, failed = send_bulk(
        excel=excel,
        subject_template=args.subject,
        body_template=body_template,
        smtp_host=smtp_host,
        smtp_port=smtp_port,
        username=username,
        password=password,
        from_address=from_address,
        attachments_dir=Path(args.attachments_dir) if args.attachments_dir else None,
        body_html_template=body_html_template,
        dry_run=args.dry_run,
        delay_seconds=args.delay,
        use_starttls=use_starttls,
        use_ssl=use_ssl,
        timeout=timeout,
        cc_addresses=cc_addresses,
        bcc_addresses=bcc_addresses,
    )
    if failed and not args.dry_run:
        sys.exit(1)


if __name__ == "__main__":
    try:
        main_bulk()
    except (ValueError, FileNotFoundError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
