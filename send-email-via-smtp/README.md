# Send Email via SMTP

Send single emails with attachments (`main.py`) or personalized bulk
invitations driven by an Excel roster (`send_bulk.py`), using any SMTP
server (Gmail, Outlook, your company server, or a local test relay).
No credentials live in the code — everything is configured in
gitignored `.env` / `.env.local` files, so this repo stays safe to publish.

## What you can do

- Send one email with multiple attachments of any type (PDF, PNG, …).
- Send personalized bulk emails: each recipient gets their own name in
  the body and only their own attachments, listed in an Excel sheet.
- Plain-text bodies, with optional HTML version.
- Works with STARTTLS (port 587) and implicit SSL (port 465).

## Project files

| File              | Purpose                                                      |
| ----------------- | ------------------------------------------------------------ |
| `main.py`         | Send a single email. Configuration from env / `.env` files.  |
| `send_bulk.py`    | Send personalized bulk emails from an Excel roster.          |
| `helpers_smtp.py` | Shared helpers (config loading, validation). You can ignore it. |
| `.env.example`    | Template for all settings. Copy it to `.env` / `.env.local`. |
| `requirements.txt`| Only needed for Excel bulk sends (`openpyxl`).              |

## Prerequisites

- Python 3.11 or newer. Check with:
  ```powershell
  python --version
  ```
- An SMTP account (see provider settings below), or a local test relay.

## Installation

1. Open a terminal in this folder (`send-email-via-smtp`).
2. (Recommended) Create and activate a virtual environment:
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```
3. Install the Excel dependency (only needed for `send_bulk.py`):
   ```powershell
   pip install -r requirements.txt
   ```

## Configuration (do this first)

1. Copy the template:
   ```powershell
   Copy-Item .env.example .env.local
   ```
   Use `.env.local` for your real, personal settings (it is gitignored
   and never committed). Real environment variables always override
   whatever is in the files.
2. Open `.env.local` in a text editor and fill in at least:
   ```ini
   SMTP_HOST=smtp.gmail.com
   SMTP_PORT=587
   SMTP_USERNAME=you@gmail.com
   SMTP_PASSWORD=your-app-password
   FROM_ADDRESS=you@gmail.com
   ```
3. Pick the right port/security combo:

   | Port | Setting               | When to use              |
   | ---- | --------------------- | ------------------------ |
   | 587  | `EMAIL_USE_STARTTLS=true`  | Standard sending (Gmail default) |
   | 465  | `EMAIL_USE_SSL=true`, `EMAIL_USE_STARTTLS=false` | Implicit SSL servers |

   If you leave both security settings empty, port 465 auto-selects SSL
   and everything else defaults to STARTTLS.

### Provider examples

- **Gmail**: `SMTP_HOST=smtp.gmail.com`, `SMTP_PORT=587`,
  `EMAIL_USE_STARTTLS=true`. You must create an *App Password*
  (Google Account → Security → 2-Step Verification → App passwords) and
  use it as `SMTP_PASSWORD` — your normal password will not work.
- **Outlook / Microsoft 365**: `SMTP_HOST=smtp.office365.com`,
  `SMTP_PORT=587`, `EMAIL_USE_STARTTLS=true`.
- **Local testing** (no real emails): run a relay like MailHog and leave
  `SMTP_USERNAME` / `SMTP_PASSWORD` empty.

## Usage A — single email (`main.py`)

1. Set the recipients and content in `.env.local`:
   ```ini
   TO_ADDRESSES=friend@example.com
   EMAIL_SUBJECT=Hello!
   EMAIL_BODY=Hi there, see attached.
   EMAIL_ATTACHMENTS=./file1.pdf; ./file2.png
   ```
   Multiple addresses/attachments are separated with `;` or `,`.
2. Send:
   ```powershell
   python main.py
   ```
3. Optional extras in `.env.local`: `EMAIL_CC`, `EMAIL_BCC`,
   `EMAIL_USE_STARTTLS`, `EMAIL_USE_SSL`, `SMTP_TIMEOUT`.

## Usage B — personalized bulk emails (`send_bulk.py`)

Each row of an Excel file becomes one email, with the recipient's name
filled into the body and only their attachments attached.

### Step 1 — build the Excel roster

Create e.g. `guests.xlsx` with this header row (headers are
case-insensitive; attachment columns may be as many as you like):

| RECIPIENT_NAME | RECIPIENT_EMAIL_ADDRESS | ATTACHMENT1_FILE_PATH | ATTACHMENT2_FILE_PATH | ATTACHMENT3_FILE_PATH |
| -------------- | ----------------------- | ---------------- | ---------------- | ---------------- |
| Asha           | asha@example.com        | invite_asha.pdf  |                  |                  |
| Ben            | ben@example.com         | invite_ben.pdf   | menu.pdf         |                  |

Rules:

- `RECIPIENT_NAME` → inserted into the body wherever you write
  `[RECIPIENT_NAME]` (or `{RECIPIENT_NAME}`).
- `RECIPIENT_EMAIL_ADDRESS` → where that row's email goes. Invalid
  addresses are reported with row numbers before anything is sent.
- Every column whose header starts with `ATTACHMENT` is treated as an
  attachment column. A cell may hold several files separated with `;`.
- Blank rows are skipped.
- Relative file names are looked up in `--attachments-dir` (default:
  the folder containing the Excel file).

### Step 2 — write the body template

Example (`body_template.txt`):

```text
Dear [RECIPIENT_NAME],

We are glad to invite you to our family dinner.

Regards,
XYZ
```

For an HTML version as well, add e.g. `body_template.html` with the
same `[RECIPIENT_NAME]` placeholders.

### Step 3 — dry-run first (validates everything, sends nothing)

```powershell
python send_bulk.py --excel guests.xlsx --body-file body_template.txt --dry-run
```

Fix any reported errors (bad emails, missing attachments) and repeat
until clean.

### Step 4 — send

```powershell
python send_bulk.py --excel guests.xlsx --subject "You are invited, [RECIPIENT_NAME]!" --body-file body_template.txt --attachments-dir ./invites --delay 1
```

With HTML:

```powershell
python send_bulk.py --excel guests.xlsx --body-file body_template.txt --body-html-file body_template.html --dry-run
```

All flags can alternatively live in `.env.local`
(`BULK_EXCEL`, `BULK_SUBJECT`, `BULK_BODY`, `BULK_BODY_FILE`,
`BULK_BODY_HTML_FILE`, `BULK_ATTACHMENTS_DIR`, `BULK_DELAY_SECONDS`,
`BULK_DRY_RUN`, `BULK_CC`, `BULK_BCC`); command-line flags win when both are set.

### CC / BCC in bulk mode

`send_bulk.py` honors the same `EMAIL_CC` / `EMAIL_BCC` from `.env.local`
as `main.py` — no extra setup needed. You can also override per run:

```powershell
python send_bulk.py --excel guests.xlsx --cc boss@example.com --bcc archive@example.com --dry-run
```

Or set `BULK_CC` / `BULK_BCC` in `.env.local` to keep bulk CCs separate
from single-send CCs. Note: CC/BCC addresses are added to **every**
email, so a CC'd person receives one copy per recipient.

## Body formats supported

| Format     | How                                                     |
| ---------- | ------------------------------------------------------- |
| Plain text | Default. `EMAIL_BODY` / `--body` / `--body-file`.       |
| HTML       | `--body-html-file` (plus a plain-text fallback).        |
| Template   | `[RECIPIENT_NAME]` or `{RECIPIENT_NAME}` in subject and body (bulk mode). |

## Troubleshooting

| Error | Fix |
| ----- | --- |
| `SMTP_HOST is not configured` / `No recipients configured` | You skipped the Configuration step — fill in `.env.local`. |
| `EMAIL_USE_SSL and EMAIL_USE_STARTTLS are both enabled` | Use one only: port 465 → `EMAIL_USE_STARTTLS=false`; port 587 → `EMAIL_USE_SSL=false`, `EMAIL_USE_STARTTLS=true`. |
| `STARTTLS failed … port 465` | Same as above — port 465 needs SSL, not STARTTLS. |
| `Attachment(s) not found` | Check file names and `--attachments-dir`. Run `--dry-run` first. |
| Gmail `Authentication failed` | Use an App Password, not your login password. |
| `openpyxl is required` | Run `pip install -r requirements.txt`. |
| Total attachments rejected | Providers typically cap at ~25 MB per email; the script warns above that size. |

## Security

- Never put real passwords in code, chat logs, or screenshots.
- `.env` and `.env.local` are gitignored — only `.env.example`
  (placeholders) is committed. Verify with
  `git status` before committing.
- Prefer an App Password over your main account password wherever
  the provider supports it.
