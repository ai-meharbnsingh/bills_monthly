#!/usr/bin/env python3
"""
Railway Cron Job: Monthly Bill Generator

Runs on the 2nd of each month. Generates mobile and landline bills
from Excel templates, converts to PDF via LibreOffice, and emails them.

All credentials are read from environment variables (set in Railway dashboard).
"""

import os
import sys
import smtplib
import random
import string
import subprocess
import tempfile
from datetime import datetime
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Templates live next to this script
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MOBILE_TEMPLATE = os.path.join(SCRIPT_DIR, 'Mobile_Bill_Template.xlsx')
LANDLINE_TEMPLATE = os.path.join(SCRIPT_DIR, 'Landline_Bill_Template.xlsx')


def get_env(name):
    """Get a required environment variable or exit with an error."""
    value = os.environ.get(name)
    if not value:
        print(f"ERROR: Missing required environment variable: {name}")
        sys.exit(1)
    return value


def generate_random_bill_no():
    letters = ''.join(random.choices(string.ascii_uppercase, k=2))
    digits = ''.join(random.choices(string.digits, k=14))
    return f"{letters}{digits}"


def update_excel_file(template_path, temp_dir, is_mobile_bill):
    """Update an Excel template with current billing dates and a random bill number."""
    print(f"  Loading template: {os.path.basename(template_path)}")
    wb = load_workbook(filename=template_path)
    ws = wb.active

    today = datetime.now()
    statement_date_dt = today.replace(day=23) - relativedelta(months=1)
    period_start_dt = today.replace(day=23) - relativedelta(months=2)
    period_end_dt = today.replace(day=22) - relativedelta(months=1)
    due_date_dt = today.replace(day=12)

    statement_date_str = f"Statement Date:{statement_date_dt.strftime('%d %b %Y')}"
    period_start_str = period_start_dt.strftime('%d %b %Y')
    period_end_str = period_end_dt.strftime('%d %b %Y')
    statement_period_str = f"Statement Period:{period_start_str}-{period_end_str}"
    due_date_q7_str = due_date_dt.strftime('%d-%b-%Y')
    due_date_s12_str = f"Amount after due date ({due_date_dt.strftime('%d %B')})"
    new_bill_no = f"Bill No. {generate_random_bill_no()}"

    if is_mobile_bill:
        ws['J5'] = statement_date_str
        ws['J6'] = statement_period_str
        ws['Q7'] = due_date_q7_str
        ws['S12'] = due_date_s12_str
        ws['H82'] = new_bill_no
        temp_filename = "temp_mobile_bill.xlsx"
    else:
        ws['J7'] = statement_date_str
        ws['J8'] = statement_period_str
        ws['Q7'] = due_date_q7_str
        ws['S12'] = due_date_s12_str
        ws['H82'] = new_bill_no
        temp_filename = "temp_landline_bill.xlsx"

    temp_excel_path = os.path.join(temp_dir, temp_filename)
    wb.save(temp_excel_path)
    print(f"  Generated: {temp_filename}")
    return temp_excel_path


def convert_excel_to_pdf(excel_path, pdf_path):
    """Convert Excel to PDF using LibreOffice headless (Linux)."""
    print(f"  Converting {os.path.basename(excel_path)} to PDF...")
    outdir = os.path.dirname(pdf_path)

    result = subprocess.run(
        [
            'libreoffice', '--headless', '--calc',
            '--convert-to', 'pdf',
            '--outdir', outdir,
            excel_path,
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )

    if result.returncode != 0:
        print(f"  LibreOffice stderr: {result.stderr}")
        raise RuntimeError(f"LibreOffice conversion failed (exit {result.returncode})")

    # LibreOffice outputs <basename>.pdf in outdir
    lo_output = os.path.join(outdir, os.path.splitext(os.path.basename(excel_path))[0] + '.pdf')
    if not os.path.exists(lo_output):
        raise RuntimeError(f"PDF not found after conversion: {lo_output}")

    # Rename to the desired filename (e.g. "Mobile Bill February-26.pdf")
    if lo_output != pdf_path:
        os.rename(lo_output, pdf_path)

    print(f"  PDF created: {os.path.basename(pdf_path)}")


def send_email(sender_email, sender_password, recipient_email, smtp_server, smtp_port, attachments):
    """Send an email with the generated PDF bills attached."""
    print("\n3. Sending email...")

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f"Your Bills for {datetime.now().strftime('%B %Y')}"
    msg.attach(MIMEText("Please find your monthly bills attached.\n\nThank you.", 'plain'))

    for file_path in attachments:
        with open(file_path, 'rb') as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
            msg.attach(part)
        print(f"  Attached: {os.path.basename(file_path)}")

    print(f"  Connecting to {smtp_server}:{smtp_port}...")
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)
    server.send_message(msg)
    server.quit()
    print(f"  Email sent successfully to {recipient_email}")


def main():
    print("=" * 60)
    print("MONTHLY BILL GENERATOR (Railway Cron)")
    print(f"Run time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # Read credentials from environment
    sender_email = get_env('SENDER_EMAIL')
    sender_password = get_env('SENDER_PASSWORD')
    recipient_email = get_env('RECIPIENT_EMAIL')
    smtp_server = get_env('SMTP_SERVER')
    smtp_port = int(get_env('SMTP_PORT'))

    temp_dir = tempfile.mkdtemp(prefix='bills_')
    files_to_cleanup = []

    try:
        # --- Mobile Bill ---
        print("\n1. Generating Mobile Bill...")
        mobile_xlsx = update_excel_file(MOBILE_TEMPLATE, temp_dir, is_mobile_bill=True)
        files_to_cleanup.append(mobile_xlsx)

        mobile_pdf = os.path.join(temp_dir, f"Mobile Bill {datetime.now().strftime('%B-%y')}.pdf")
        convert_excel_to_pdf(mobile_xlsx, mobile_pdf)
        files_to_cleanup.append(mobile_pdf)

        # --- Landline Bill ---
        print("\n2. Generating Landline Bill...")
        landline_xlsx = update_excel_file(LANDLINE_TEMPLATE, temp_dir, is_mobile_bill=False)
        files_to_cleanup.append(landline_xlsx)

        landline_pdf = os.path.join(temp_dir, f"Landline Bill {datetime.now().strftime('%B-%y')}.pdf")
        convert_excel_to_pdf(landline_xlsx, landline_pdf)
        files_to_cleanup.append(landline_pdf)

        # --- Send Email ---
        send_email(sender_email, sender_password, recipient_email,
                   smtp_server, smtp_port, [mobile_pdf, landline_pdf])

        print("\n" + "=" * 60)
        print("SUCCESS! Bills generated and emailed.")
        print("=" * 60)

    except Exception as e:
        print("\n" + "=" * 60)
        print(f"ERROR: {e}")
        print("=" * 60)
        sys.exit(1)

    finally:
        print("\nCleaning up temporary files...")
        for f in files_to_cleanup:
            if os.path.exists(f):
                try:
                    os.remove(f)
                    print(f"  Deleted: {os.path.basename(f)}")
                except Exception:
                    print(f"  Warning: Could not delete {os.path.basename(f)}")
        try:
            os.rmdir(temp_dir)
        except Exception:
            pass


if __name__ == "__main__":
    main()
