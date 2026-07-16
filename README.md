# Outlook Web Outreach (Selenium)

Excel-driven email outreach for **Outlook on the web**. The script opens Chrome, lets you sign in to Outlook yourself, personalises each message from a template, and writes the result back into the same workbook.

It never asks for or stores your password. Login stays in a local Chrome profile folder that is gitignored.

## Features

- Read recipients from an Excel `Recipients` sheet
- Choose subject/body from a `Templates` sheet via `Template Key`
- Personalise with placeholders:
  - `{{First Name}}` — first word of the Name column
  - `{{name}}` — same as first name
  - `{{Company Name}}` / `{{company}}` — Company column
- Optional `from_email` to pick a sender already available in Outlook’s **From** menu
- Dry-run mode creates drafts without sending
- Updates `Status`, `Sent At`, and `Notes` after every row
- Skips `SENT` rows on later runs

## Requirements

- Windows, macOS, or Linux
- Python 3.10+
- Google Chrome
- An Outlook / Microsoft 365 account you can open in the browser

## Quick start

### 1. Clone and install

```powershell
git clone <your-repo-url>
cd build-a-python-automation-tool-using

py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

On macOS/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Create your Excel workbook

Create `outreach.xlsx` (or any name you set in `config.json`) with two sheets.

**Recipients**

| Name | Company | Email | Template Key | Status | Sent At | Notes |
|------|---------|-------|--------------|--------|---------|-------|
| Priya Sharma | Acme Corp | you@example.com | default | READY | | |

**Templates**

| Template Key | Subject | Body |
|--------------|---------|------|
| default | Exploring KYC for {{Company Name}} | Hi {{First Name}}, … |

Rules:

- Headers must match exactly (including `Company`).
- `Status` blank, `READY`, or `FAILED` = eligible. `SENT` is skipped.
- Close the workbook in Excel before running so the script can save updates.

### 3. Configure `config.json`

```json
{
  "workbook_path": "outreach.xlsx",
  "outlook_url": "https://outlook.office.com/mail/",
  "from_email": "",
  "chrome_profile_path": "chrome_profile",
  "dry_run": true,
  "max_emails_per_run": 1,
  "delay_between_emails_seconds": 4,
  "wait_timeout_seconds": 45
}
```

| Setting | Meaning |
|---------|---------|
| `workbook_path` | Path to your `.xlsx` (relative to this folder or absolute) |
| `outlook_url` | Outlook Web URL |
| `from_email` | Optional sender from Outlook’s From menu; `""` = default account |
| `chrome_profile_path` | Local folder for the automation Chrome profile |
| `dry_run` | `true` = draft only; `false` = actually send |
| `max_emails_per_run` | Cap per run (start with `1`) |
| `delay_between_emails_seconds` | Pause after each email |
| `wait_timeout_seconds` | How long to wait for Outlook UI controls |

### 4. Run

```powershell
python outreach.py
```

1. Chrome opens with the project profile.
2. Sign in to Outlook in that window (first run only, then session is reused).
3. Wait until your inbox is visible.
4. Press **Enter** in the terminal.
5. Review drafts (or sent mail), then press **Enter** again to close Chrome.

## Recommended first test

1. Keep `"dry_run": true`.
2. Set `"max_emails_per_run": 1`.
3. Put **your own email** as the only recipient.
4. Run, open Drafts in Outlook, confirm subject/body look right.
5. Only then set `"dry_run": false` and send to people you are allowed to contact.

## Placeholders

| Placeholder | Source |
|-------------|--------|
| `{{First Name}}` | First word of `Name` |
| `{{name}}` | Same as first name |
| `{{Company Name}}` | `Company` column |
| `{{company}}` | Same as company |

Example subject:

`Exploring KYC, CKYC & DPDP Solutions for {{Company Name}}`

## Changing Outlook account

- **Different login:** delete the `chrome_profile` folder, run again, sign in with the new account.
- **Different From address on the same login:** set `from_email` in `config.json` to an address shown in Outlook’s From menu.

## Safeguards

- Do not commit `chrome_profile/` (already gitignored).
- Do not commit real contact workbooks if they contain personal data (`*.xlsx` is gitignored).
- Start small; respect your organisation’s policies and anti-spam / consent rules.
- Failed rows are marked `FAILED` with a short note; fix and set Status back to `READY` to retry.

## Troubleshooting

| Problem | What to try |
|---------|-------------|
| Chrome fails to start / DevToolsActivePort | Close leftover automation Chrome windows and run again |
| Timed out waiting for a control | Outlook UI may differ; check `outlook_failure_*.png` and adjust selectors in `OUTLOOK` inside `outreach.py` |
| Body missing / emoji error | Body is inserted via JavaScript so emoji is supported; confirm the Templates `Body` cell is not empty |
| Workbook not updating | Close the file in Excel, then rerun |
| Unknown template key | Every recipient `Template Key` must exist on the Templates sheet |

## Project files

| File | Purpose |
|------|---------|
| `outreach.py` | Main automation script |
| `config.json` | Runtime settings |
| `requirements.txt` | Python dependencies |
| `README.md` | This guide |

Create your own `.xlsx` locally; it is not included in the repo on purpose.

## License / use

Use only for legitimate outreach to contacts you are permitted to email.
