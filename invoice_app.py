import streamlit as st
import re
import PyPDF2
import openpyxl
from io import BytesIO

st.title("📄 DAFF Invoices PDF to Excel Converter")

uploaded_file = st.file_uploader("Upload Invoice PDF", type=["pdf"])

if uploaded_file is not None:
    # Create Excel workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Invoices'
    ws.append(['Invoice Date', 'BPAY Biller Code', 'BPAY Reference', 'Invoice #', 'HAWB / Entry No', 'Total Outstanding'])

    pdf_reader = PyPDF2.PdfReader(uploaded_file)

    def extract_max_amount(text):
        amounts = re.findall(r'\$([0-9,]+\.\d{2})', text)
        if amounts:
            max_amount = max(float(a.replace(',', '')) for a in amounts)
            return f"${max_amount:,.2f}"
        return 'NOT FOUND'

    for page in pdf_reader.pages:
        text = page.extract_text() or ""

        invoice_date = re.search(r'Issue Date[:\s]*([0-9]{1,2}[/-][A-Za-z]{3}[/-][0-9]{2,4})', text)
        biller_code = re.search(r'Biller Code[:\s]*(\d+)', text)
        bpay_ref = re.search(r'Reference[:\s]*(\d+)', text)
        invoice_no = re.search(r'(100\d{7,})', text)
        tracking = re.search(r'(?:Flight:\s*)?((?:1Z|W)[0-9A-Z]{10,})', text)
        entry = re.search(r'\b(Q[0-9A-Z]+)\b', text)
        total_outstanding = extract_max_amount(text)

        hawb_or_entry = tracking.group(1) if tracking else (entry.group(1) if entry else 'NOT FOUND')

        ws.append([
            invoice_date.group(1) if invoice_date else 'NOT FOUND',
            biller_code.group(1) if biller_code else 'NOT FOUND',
            bpay_ref.group(1) if bpay_ref else 'NOT FOUND',
            invoice_no.group(1) if invoice_no else 'NOT FOUND',
            hawb_or_entry,
            total_outstanding
        ])

    # Save Excel to memory
    excel_data = BytesIO()
    wb.save(excel_data)
    excel_data.seek(0)

    st.success("✅ Processing complete!")

    st.download_button(
        label="📥 Download Excel File",
        data=excel_data,
        file_name="invoice_breakdown.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
