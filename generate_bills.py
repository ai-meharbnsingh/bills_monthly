"""
Windows GUI Bill Generator (tkinter)

Generates mobile and landline bills from Excel templates, converts to PDF
via VBScript/Excel COM, and emails them. Shows a status window during processing.

Reads credentials from config.ini.
"""

import os
import sys
import subprocess
import tempfile
import threading
import configparser
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

from bill_utils import update_excel_file, send_email_smtp, pdf_filename

# --- Configuration Constants ---
CONFIG_FILE = 'config.ini'
MOBILE_TEMPLATE = 'Mobile_Bill_Template.xlsx'
LANDLINE_TEMPLATE = 'Landline_Bill_Template.xlsx'


# --- Status Window Class ---
class StatusWindow:
    def __init__(self, parent):
        self._root = parent
        self.top = tk.Toplevel(parent)
        self.top.title("Processing...")

        parent.update_idletasks()
        width, height = 350, 100
        x = (parent.winfo_screenwidth() // 2) - (width // 2)
        y = (parent.winfo_screenheight() // 2) - (height // 2)
        self.top.geometry(f'{width}x{height}+{x}+{y}')

        self.label = tk.Label(self.top, text="Initializing...",
                              padx=20, pady=20, font=("Helvetica", 10))
        self.label.pack()
        self.top.lift()

    def update_status(self, text):
        """Thread-safe status update using root.after()."""
        self._root.after(0, self._do_update, text)

    def _do_update(self, text):
        self.label.config(text=text)

    def close(self):
        """Thread-safe close."""
        self._root.after(0, self.top.destroy)


# --- Helper Functions ---
def resource_path(relative_path):
    """Resolve path for PyInstaller bundles or normal execution."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def convert_excel_to_pdf(excel_path, pdf_path):
    """Convert Excel to PDF using VBScript — Windows only."""
    try:
        abs_excel_path = os.path.abspath(excel_path)
        abs_pdf_path = os.path.abspath(pdf_path)

        vbs_script = f'''
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open("{abs_excel_path}")
objWorkbook.ExportAsFixedFormat 0, "{abs_pdf_path}"
objWorkbook.Close False

objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing
'''
        vbs_path = os.path.join(tempfile.gettempdir(), "excel_to_pdf.vbs")
        with open(vbs_path, 'w') as f:
            f.write(vbs_script)

        result = subprocess.run(
            ['cscript', '//nologo', vbs_path],
            capture_output=True, text=True, timeout=30
        )

        try:
            os.remove(vbs_path)
        except OSError:
            pass

        if os.path.exists(abs_pdf_path):
            return True, None
        else:
            error_msg = result.stderr if result.stderr else "Unknown error"
            return False, (
                f"Could not convert to PDF.\n"
                f"Please ensure MS Excel is installed.\n\nError: {error_msg}"
            )

    except subprocess.TimeoutExpired:
        return False, "Excel conversion timed out. Please close any Excel windows and try again."
    except Exception as e:
        return False, (
            f"Could not convert to PDF.\n"
            f"Please ensure MS Excel is installed.\n\nError: {e}"
        )


def main_process(root, status_window, config_path, mobile_template_path,
                  landline_template_path):
    """Worker thread: generate bills, convert to PDF, and email."""
    temp_dir = tempfile.mkdtemp(prefix='bills_')
    files_to_cleanup = []

    try:
        status_window.update_status("Loading configuration...")
        config = configparser.ConfigParser()
        config.read(config_path)

        # --- Process Mobile Bill ---
        status_window.update_status("Generating Mobile Bill...")
        mobile_temp_xlsx, error = update_excel_file(
            mobile_template_path, temp_dir, is_mobile_bill=True
        )
        if error:
            raise Exception(error)
        files_to_cleanup.append(mobile_temp_xlsx)

        status_window.update_status("Converting Mobile Bill to PDF...")
        mobile_pdf_path = os.path.join(temp_dir, pdf_filename("Mobile"))
        files_to_cleanup.append(mobile_pdf_path)
        success, error = convert_excel_to_pdf(mobile_temp_xlsx, mobile_pdf_path)
        if not success:
            raise Exception(error)

        # --- Process Landline Bill ---
        status_window.update_status("Generating Landline Bill...")
        landline_temp_xlsx, error = update_excel_file(
            landline_template_path, temp_dir, is_mobile_bill=False
        )
        if error:
            raise Exception(error)
        files_to_cleanup.append(landline_temp_xlsx)

        status_window.update_status("Converting Landline Bill to PDF...")
        landline_pdf_path = os.path.join(temp_dir, pdf_filename("Landline"))
        files_to_cleanup.append(landline_pdf_path)
        success, error = convert_excel_to_pdf(landline_temp_xlsx, landline_pdf_path)
        if not success:
            raise Exception(error)

        # --- Send Email ---
        status_window.update_status("Connecting to email server...")
        try:
            sender_email = config.get('Email', 'SENDER_EMAIL')
            sender_password = config.get('Email', 'SENDER_PASSWORD')
            recipient_email = config.get('Email', 'RECIPIENT_EMAIL')
            smtp_server = config.get('Email', 'SMTP_SERVER')
            smtp_port = int(config.get('Email', 'SMTP_PORT'))
        except (configparser.NoSectionError, configparser.NoOptionError) as e:
            raise Exception(f"Could not read setting from config.ini: {e}")

        pdfs_to_send = [mobile_pdf_path, landline_pdf_path]
        success, error, recipient = send_email_smtp(
            sender_email, sender_password, recipient_email,
            smtp_server, smtp_port, pdfs_to_send
        )
        if not success:
            raise Exception(error)

        status_window.close()
        root.after(0, messagebox.showinfo, "Success!",
                   f"Bills have been successfully generated and sent to {recipient}")

    except Exception as e:
        status_window.close()
        root.after(0, messagebox.showerror, "An Error Occurred", str(e))

    finally:
        for f in files_to_cleanup:
            if os.path.exists(f):
                try:
                    os.remove(f)
                except OSError:
                    pass
        try:
            os.rmdir(temp_dir)
        except OSError:
            pass
        root.after(0, root.destroy)


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    status_win = StatusWindow(root)

    config_file_path = resource_path(CONFIG_FILE)
    mobile_template_file_path = resource_path(MOBILE_TEMPLATE)
    landline_template_file_path = resource_path(LANDLINE_TEMPLATE)

    main_thread = threading.Thread(
        target=main_process,
        args=(root, status_win, config_file_path,
              mobile_template_file_path, landline_template_file_path)
    )
    main_thread.start()

    root.mainloop()