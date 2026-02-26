"""
Windows Console Bill Generator

Generates mobile and landline bills from Excel templates, converts to PDF
via win32com (Excel COM automation), and emails them.

Reads credentials from config.ini. Requires pywin32 and MS Excel.
"""

import os
import sys
import configparser
import tempfile
from datetime import datetime

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 is required. Install with: pip install pywin32")
    print("       This script only works on Windows with MS Excel installed.")
    sys.exit(1)

from bill_utils import update_excel_file, send_email_smtp, pdf_filename

# --- Configuration ---
CONFIG_FILE = 'config.ini'
MOBILE_TEMPLATE = 'Mobile_Bill_Template.xlsx'
LANDLINE_TEMPLATE = 'Landline_Bill_Template.xlsx'


def convert_excel_to_pdf(excel_path, pdf_path):
    """Convert Excel to PDF using win32com (Excel COM automation)."""
    print("  Converting to PDF...")
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(os.path.abspath(excel_path))
        workbook.ActiveSheet.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
        workbook.Close(False)
        print(f"  PDF created: {os.path.basename(pdf_path)}")
        return True, None
    except Exception as e:
        return False, f"Could not convert to PDF: {e}"
    finally:
        if excel:
            excel.Quit()


def main():
    print("=" * 60)
    print("BILL GENERATION SCRIPT")
    print("=" * 60)

    temp_dir = tempfile.mkdtemp(prefix='bills_')
    files_to_cleanup = []

    try:
        # Load config
        print("\nLoading configuration...")
        config = configparser.ConfigParser()
        config.read(CONFIG_FILE)
        print(f"  Config loaded from {CONFIG_FILE}")

        # --- Process Mobile Bill ---
        print("\n1. Generating Mobile Bill...")
        mobile_temp_xlsx, error = update_excel_file(
            MOBILE_TEMPLATE, temp_dir, is_mobile_bill=True
        )
        if error:
            raise Exception(error)
        files_to_cleanup.append(mobile_temp_xlsx)
        print(f"  Generated: {os.path.basename(mobile_temp_xlsx)}")

        mobile_pdf_path = os.path.join(temp_dir, pdf_filename("Mobile"))
        files_to_cleanup.append(mobile_pdf_path)
        success, error = convert_excel_to_pdf(mobile_temp_xlsx, mobile_pdf_path)
        if not success:
            raise Exception(error)

        # --- Process Landline Bill ---
        print("\n2. Generating Landline Bill...")
        landline_temp_xlsx, error = update_excel_file(
            LANDLINE_TEMPLATE, temp_dir, is_mobile_bill=False
        )
        if error:
            raise Exception(error)
        files_to_cleanup.append(landline_temp_xlsx)
        print(f"  Generated: {os.path.basename(landline_temp_xlsx)}")

        landline_pdf_path = os.path.join(temp_dir, pdf_filename("Landline"))
        files_to_cleanup.append(landline_pdf_path)
        success, error = convert_excel_to_pdf(landline_temp_xlsx, landline_pdf_path)
        if not success:
            raise Exception(error)

        # --- Send Email ---
        print("\n3. Sending email...")
        try:
            sender_email = config.get('Email', 'SENDER_EMAIL')
            sender_password = config.get('Email', 'SENDER_PASSWORD')
            recipient_email = config.get('Email', 'RECIPIENT_EMAIL')
            smtp_server = config.get('Email', 'SMTP_SERVER')
            smtp_port = int(config.get('Email', 'SMTP_PORT'))
        except (configparser.NoSectionError, configparser.NoOptionError) as e:
            raise Exception(f"Could not read setting '{e.option}' from config.ini.")

        pdfs_to_send = [mobile_pdf_path, landline_pdf_path]
        for p in pdfs_to_send:
            print(f"  Attached: {os.path.basename(p)}")
        print(f"  Connecting to {smtp_server}:{smtp_port}...")

        success, error, recipient = send_email_smtp(
            sender_email, sender_password, recipient_email,
            smtp_server, smtp_port, pdfs_to_send
        )
        if not success:
            raise Exception(error)

        print(f"  Email sent successfully to {recipient}")
        print("\n" + "=" * 60)
        print("SUCCESS!")
        print(f"Bills sent to: {recipient}")
        print("=" * 60)

    except Exception as e:
        print("\n" + "=" * 60)
        print("ERROR OCCURRED")
        print(f"{e}")
        print("=" * 60)
        sys.exit(1)

    finally:
        print("\nCleaning up temporary files...")
        for f in files_to_cleanup:
            if os.path.exists(f):
                try:
                    os.remove(f)
                    print(f"  Deleted: {os.path.basename(f)}")
                except OSError:
                    print(f"  Warning: Could not delete {os.path.basename(f)}")
        try:
            os.rmdir(temp_dir)
        except OSError:
            pass


if __name__ == "__main__":
    main()
