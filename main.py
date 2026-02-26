#!/usr/bin/env python3
"""
Railway Cron Job: Monthly Bill Generator

Runs on the 2nd of each month. Generates mobile and landline bills
from Excel templates, converts to PDF via LibreOffice, and emails them.

All credentials are read from environment variables (set in Railway dashboard).
"""

import os
import sys
import subprocess
import tempfile
from datetime import datetime

from bill_utils import update_excel_file, send_email_smtp, pdf_filename

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
    smtp_port_raw = get_env('SMTP_PORT')
    if not smtp_port_raw.isdigit() or not 1 <= int(smtp_port_raw) <= 65535:
        print(f"ERROR: Invalid SMTP_PORT value: '{smtp_port_raw}'. Must be a number between 1 and 65535.")
        sys.exit(1)
    smtp_port = int(smtp_port_raw)

    temp_dir = tempfile.mkdtemp(prefix='bills_')
    files_to_cleanup = []

    try:
        # --- Mobile Bill ---
        print("\n1. Generating Mobile Bill...")
        mobile_xlsx, error = update_excel_file(MOBILE_TEMPLATE, temp_dir, is_mobile_bill=True)
        if error:
            raise RuntimeError(error)
        files_to_cleanup.append(mobile_xlsx)
        print(f"  Generated: {os.path.basename(mobile_xlsx)}")

        mobile_pdf = os.path.join(temp_dir, pdf_filename("Mobile"))
        convert_excel_to_pdf(mobile_xlsx, mobile_pdf)
        files_to_cleanup.append(mobile_pdf)

        # --- Landline Bill ---
        print("\n2. Generating Landline Bill...")
        landline_xlsx, error = update_excel_file(LANDLINE_TEMPLATE, temp_dir, is_mobile_bill=False)
        if error:
            raise RuntimeError(error)
        files_to_cleanup.append(landline_xlsx)
        print(f"  Generated: {os.path.basename(landline_xlsx)}")

        landline_pdf = os.path.join(temp_dir, pdf_filename("Landline"))
        convert_excel_to_pdf(landline_xlsx, landline_pdf)
        files_to_cleanup.append(landline_pdf)

        # --- Send Email ---
        print("\n3. Sending email...")
        attachments = [mobile_pdf, landline_pdf]
        for a in attachments:
            print(f"  Attached: {os.path.basename(a)}")
        print(f"  Connecting to {smtp_server}:{smtp_port}...")

        success, error, _ = send_email_smtp(
            sender_email, sender_password, recipient_email,
            smtp_server, smtp_port, attachments
        )
        if not success:
            raise RuntimeError(error)
        print(f"  Email sent successfully to {recipient_email}")

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
