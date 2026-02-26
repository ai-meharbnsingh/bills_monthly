FROM python:3.11-slim

# Install LibreOffice Calc (headless) for Excel-to-PDF conversion
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice-calc && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy shared module, templates, and entry point
COPY bill_utils.py .
COPY Mobile_Bill_Template.xlsx .
COPY Landline_Bill_Template.xlsx .
COPY main.py .

CMD ["python", "main.py"]
