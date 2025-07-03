!pip install PyPDF2 openpyxl

from google.colab import files
import re
import PyPDF2
import openpyxl

# Upload PDF file
uploaded = files.upload()
pdf_path = list(uploaded.keys())[0]
excel_path = 'invoice_Breakdown.xlsx'

# Create Excel workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Invoices'
ws.append(['Invoice Date', 'BPAY Biller Code', 'BPAY Reference', 'Invoice #', 'HAWB / Entry No', 'Total Outstanding'])

# Function to extract the largest $ amount from a page
def extract_max_amount(text):
    amounts = re.findall(r'\$([0-9,]+\.\d{2})', text)
    if amounts:
        max_amount = max(float(a.replace(',', '')) for a in amounts)
        return f"${max_amount:,.2f}"
    return 'NOT FOUND'

# Process each page independently
with open(pdf_path, 'rb') as file:
    reader = PyPDF2.PdfReader(file)
    for page in reader.pages:
        text = page.extract_text() or ""

        # Extract data per page
        invoice_date = re.search(r'Issue Date[:\s]*([0-9]{1,2}[/-][A-Za-z]{3}[/-][0-9]{2,4})', text)
        biller_code = re.search(r'Biller Code[:\s]*(\d+)', text)
        bpay_ref = re.search(r'Reference[:\s]*(\d+)', text)
        invoice_no = re.search(r'(100\d{7,})', text)  # Finds any 10+ digit number starting with 100
        tracking = re.search(r'(?:Flight:\s*)?((?:1Z|W)[0-9A-Z]{10,})', text)
        entry = re.search(r'\b(Q[0-9A-Z]+)\b', text)
        total_outstanding = extract_max_amount(text)

        # Use tracking number or entry number
        hawb_or_entry = tracking.group(1) if tracking else (entry.group(1) if entry else 'NOT FOUND')

        # Append row
        ws.append([
            invoice_date.group(1) if invoice_date else 'NOT FOUND',
            biller_code.group(1) if biller_code else 'NOT FOUND',
            bpay_ref.group(1) if bpay_ref else 'NOT FOUND',
            invoice_no.group(1) if invoice_no else 'NOT FOUND',
            hawb_or_entry,
            total_outstanding
        ])

# Save and download
wb.save(excel_path)
files.download(excel_path)
