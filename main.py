#!/usr/bin/env python3
"""
Monthly Bill Generator (GitHub Actions)

Runs on the 3rd of each month. Generates mobile and landline bills
by overlaying dates/bill numbers on template PDFs, and emails them.

All credentials are read from environment variables.
"""

import os
import sys
import tempfile
from datetime import datetime

from bill_utils import compute_billing_dates, send_email_smtp, pdf_filename
from pdf_generator import generate_mobile_bill_pdf, generate_landline_bill_pdf


def get_env(name):
    """Get a required environment variable or exit with an error."""
    value = os.environ.get(name)
    if not value:
        print(f"ERROR: Missing required environment variable: {name}")
        sys.exit(1)
    return value


def main():
    print("=" * 60)
    print("MONTHLY BILL GENERATOR")
    print(f"Run time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    # Read credentials from environment
    sender_email = get_env('SENDER_EMAIL')
    sender_password = get_env('SENDER_PASSWORD')
    recipient_email = get_env('RECIPIENT_EMAIL')
    smtp_server = get_env('SMTP_SERVER')
    smtp_port_raw = get_env('SMTP_PORT')
    if not smtp_port_raw.isdigit() or not 1 <= int(smtp_port_raw) <= 65535:
        print(f"ERROR: Invalid SMTP_PORT value: '{smtp_port_raw}'")
        sys.exit(1)
    smtp_port = int(smtp_port_raw)

    temp_dir = tempfile.mkdtemp(prefix='bills_')
    files_to_cleanup = []

    # Compute billing dates once (same for both bills)
    dates = compute_billing_dates()
    print(f"\nBilling dates: {dates['statement_date_str']}")

    try:
        # --- Mobile Bill ---
        print("\n1. Generating Mobile Bill...")
        mobile_pdf = os.path.join(temp_dir, pdf_filename("Mobile"))
        generate_mobile_bill_pdf(mobile_pdf, dates)
        files_to_cleanup.append(mobile_pdf)
        print(f"  Generated: {os.path.basename(mobile_pdf)}")

        # --- Landline Bill ---
        print("\n2. Generating Landline Bill...")
        landline_pdf = os.path.join(temp_dir, pdf_filename("Landline"))
        generate_landline_bill_pdf(landline_pdf, dates)
        files_to_cleanup.append(landline_pdf)
        print(f"  Generated: {os.path.basename(landline_pdf)}")

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
