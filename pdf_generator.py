"""
PDF Bill Generator using ReportLab overlay on template PDFs.

This approach:
1. Uses the reference PDFs as background templates (preserves all images/formatting)
2. Overlays text fields (dates, bill numbers) on top
3. Keeps the same date/bill number generation logic
"""

import os
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from pypdf import PdfReader, PdfWriter
from bill_utils import compute_billing_dates, generate_random_bill_no

# Template paths
MOBILE_TEMPLATE_PDF = "refernc_bills/Mobile Bill.pdf"
LANDLINE_TEMPLATE_PDF = "refernc_bills/Landline Bill.pdf"


def number_to_words(num):
    """Convert a number to words (Indian format)."""
    ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine']
    teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 
             'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen']
    tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']
    
    def convert_less_than_thousand(n):
        if n == 0:
            return ""
        elif n < 10:
            return ones[n]
        elif n < 20:
            return teens[n - 10]
        elif n < 100:
            return tens[n // 10] + (" " + ones[n % 10] if n % 10 != 0 else "")
        else:
            return ones[n // 100] + " Hundred" + (" " + convert_less_than_thousand(n % 100) if n % 100 != 0 else "")
    
    if num == 0:
        return "Zero"
    
    # Handle decimal (paisa)
    rupees = int(num)
    paisa = round((num - rupees) * 100)
    
    result = ""
    
    # Crores
    crores = rupees // 10000000
    if crores > 0:
        result += convert_less_than_thousand(crores) + " Crore "
        rupees %= 10000000
    
    # Lakhs
    lakhs = rupees // 100000
    if lakhs > 0:
        result += convert_less_than_thousand(lakhs) + " Lakh "
        rupees %= 100000
    
    # Thousands
    thousands = rupees // 1000
    if thousands > 0:
        result += convert_less_than_thousand(thousands) + " Thousand "
        rupees %= 1000
    
    # Remaining
    if rupees > 0:
        result += convert_less_than_thousand(rupees)
    
    result = result.strip()
    
    # Add paisa if any
    if paisa > 0:
        result += " and " + convert_less_than_thousand(paisa) + " Paisa"
    
    return result + " only"


def create_overlay_mobile(dates, bill_no):
    """Create a PDF overlay for mobile bill with text fields."""
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)
    
    # Set font
    c.setFont("Helvetica", 9)
    c.setFillColorRGB(0, 0, 0)
    
    # Statement Date (J5 area) - top right
    c.drawString(420, 692, dates['statement_date_str'])
    
    # Statement Period (J6 area)
    c.drawString(420, 678, dates['statement_period_str'])
    
    # Due Date (Q7 area) - mid right
    c.drawString(435, 625, dates['due_date_q7_str'])
    
    # Amount after due date text (S12 area)
    c.drawString(480, 120, dates['due_date_s12_str'])
    
    # Bill Number (H82 area) - bottom
    c.setFont("Helvetica", 8)
    c.drawString(280, 58, f"Bill No. {bill_no}")
    
    c.save()
    packet.seek(0)
    return packet


def create_overlay_landline(dates, bill_no):
    """Create a PDF overlay for landline bill with text fields."""
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)
    
    # Set font
    c.setFont("Helvetica", 9)
    c.setFillColorRGB(0, 0, 0)
    
    # Statement Date (J7 area) - top right
    c.drawString(420, 685, dates['statement_date_str'])
    
    # Statement Period (J8 area)
    c.drawString(420, 671, dates['statement_period_str'])
    
    # Due Date (Q7 area) - mid right
    c.drawString(435, 618, dates['due_date_q7_str'])
    
    # Amount after due date text (S12 area)
    c.drawString(480, 115, dates['due_date_s12_str'])
    
    # Bill Number (H82 area) - bottom
    c.setFont("Helvetica", 8)
    c.drawString(280, 53, f"Bill No. {bill_no}")
    
    c.save()
    packet.seek(0)
    return packet


def generate_mobile_bill_pdf(output_path, dates=None):
    """Generate mobile bill PDF by overlaying text on template."""
    if dates is None:
        dates = compute_billing_dates()
    
    bill_no = generate_random_bill_no()
    
    # Read template PDF
    template_reader = PdfReader(MOBILE_TEMPLATE_PDF)
    template_page = template_reader.pages[0]
    
    # Create overlay
    overlay_pdf = PdfReader(create_overlay_mobile(dates, bill_no))
    overlay_page = overlay_pdf.pages[0]
    
    # Merge overlay onto template
    template_page.merge_page(overlay_page)
    
    # Write output
    writer = PdfWriter()
    writer.add_page(template_page)
    
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    return output_path


def generate_landline_bill_pdf(output_path, dates=None):
    """Generate landline bill PDF by overlaying text on template."""
    if dates is None:
        dates = compute_billing_dates()
    
    bill_no = generate_random_bill_no()
    
    # Read template PDF
    template_reader = PdfReader(LANDLINE_TEMPLATE_PDF)
    template_page = template_reader.pages[0]
    
    # Create overlay
    overlay_pdf = PdfReader(create_overlay_landline(dates, bill_no))
    overlay_page = overlay_pdf.pages[0]
    
    # Merge overlay onto template
    template_page.merge_page(overlay_page)
    
    # Write output
    writer = PdfWriter()
    writer.add_page(template_page)
    
    with open(output_path, 'wb') as f:
        writer.write(f)
    
    return output_path


if __name__ == "__main__":
    # Test generation
    import tempfile
    
    temp_dir = tempfile.mkdtemp()
    
    print("Testing PDF generation...")
    
    mobile_pdf = os.path.join(temp_dir, "Mobile Bill Test.pdf")
    generate_mobile_bill_pdf(mobile_pdf)
    print(f"Generated: {mobile_pdf}")
    
    landline_pdf = os.path.join(temp_dir, "Landline Bill Test.pdf")
    generate_landline_bill_pdf(landline_pdf)
    print(f"Generated: {landline_pdf}")
    
    print("\nDates used:")
    dates = compute_billing_dates()
    for k, v in dates.items():
        print(f"  {k}: {v}")
