# Bills Monthly

Automated monthly bill generator that creates mobile and landline bills from Excel templates, converts them to PDF, and emails them.

## How It Works

1. Opens Excel templates (`Mobile_Bill_Template.xlsx`, `Landline_Bill_Template.xlsx`)
2. Stamps them with current billing dates and a random bill number
3. Converts Excel → PDF
4. Emails both PDFs to a recipient

## Entry Points

| Script | Platform | PDF Conversion | Config |
|---|---|---|---|
| `main.py` | Linux / Mac / Docker | LibreOffice headless | Environment variables |
| `generate_bills.py` | Windows (GUI) | VBScript + Excel COM | `config.ini` |
| `run_bills_console.py` | Windows (console) | `win32com` + Excel COM | `config.ini` |

## Setup

### Linux / Mac / Docker (Production)

1. Set environment variables:
   - `SENDER_EMAIL` — Gmail address
   - `SENDER_PASSWORD` — Gmail app password
   - `RECIPIENT_EMAIL` — Where to send the bills
   - `SMTP_SERVER` — `smtp.gmail.com`
   - `SMTP_PORT` — `587` (STARTTLS) or `465` (SSL)
2. Run `python main.py`

### Windows (Local)

1. Install Python 3.8+ and MS Excel
2. `pip install -r requirements.txt`
3. For console script: `pip install pywin32`
4. Copy `config.ini.example` → `config.ini` and fill in your credentials
5. Run `python run_bills_console.py` or `python generate_bills.py`

## Files

```
├── bill_utils.py              # Shared utility functions
├── main.py                    # Railway/Docker entry point
├── generate_bills.py          # Windows GUI (tkinter) entry point
├── run_bills_console.py       # Windows console entry point
├── test_script.py             # Cross-platform smoke test
├── Mobile_Bill_Template.xlsx  # Excel template
├── Landline_Bill_Template.xlsx
├── config.ini                 # Local credentials (gitignored)
├── requirements.txt           # Python dependencies
└── Run_Bill_Generator.bat     # Windows shortcut
```

## Testing

```bash
python test_script.py
```
