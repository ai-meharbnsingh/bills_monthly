"""
Shared utilities for bill generation.
"""

import os
import random
import string
import smtplib
from datetime import datetime
from dateutil.relativedelta import relativedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


def generate_random_bill_no():
    """Generate a random bill number: 2 uppercase letters + 14 digits."""
    letters = ''.join(random.choices(string.ascii_uppercase, k=2))
    digits = ''.join(random.choices(string.digits, k=14))
    return f"{letters}{digits}"


def compute_billing_dates():
    """Compute all billing dates relative to today."""
    today = datetime.now()
    statement_date_dt = today.replace(day=23) - relativedelta(months=1)
    period_start_dt = today.replace(day=23) - relativedelta(months=2)
    period_end_dt = today.replace(day=22) - relativedelta(months=1)
    due_date_dt = today.replace(day=12)

    return {
        'statement_date_str': f"Statement Date:{statement_date_dt.strftime('%d %b %Y')}",
        'statement_period_str': (
            f"Statement Period:{period_start_dt.strftime('%d %b %Y')}"
            f"-{period_end_dt.strftime('%d %b %Y')}"
        ),
        'due_date_q7_str': due_date_dt.strftime('%d-%b-%Y'),
        'due_date_s12_str': f"Amount after due date ({due_date_dt.strftime('%d %B')})",
        'bill_no_str': f"Bill No. {generate_random_bill_no()}",
    }


def send_email_smtp(sender_email, sender_password, recipient_email,
                    smtp_server, smtp_port, attachments):
    """Send an email with PDF bill attachments."""
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f"Your Bills for {datetime.now().strftime('%B %Y')}"
    msg.attach(MIMEText(
        "Please find your monthly bills attached.\n\nThank you.", 'plain'
    ))

    for file_path in attachments:
        with open(file_path, 'rb') as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
            part['Content-Disposition'] = (
                f'attachment; filename="{os.path.basename(file_path)}"'
            )
            msg.attach(part)

    try:
        if smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, sender_password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)
        return True, None, recipient_email
    except Exception as e:
        return False, f"Failed to send email: {e}", recipient_email


def pdf_filename(bill_type):
    """Generate a PDF filename like 'Mobile Bill February-26.pdf'."""
    return f"{bill_type} Bill {datetime.now().strftime('%B-%y')}.pdf"
