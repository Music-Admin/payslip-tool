import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

# Function to extract pay period from the first row
def extract_pay_period(file):
    df = pd.read_csv(file, nrows=1, header=None)
    return df.iloc[0, 1] if df.shape[1] > 1 else "Not Specified"

# Function to find the header row by searching for key columns
def find_header_row(file):
    df_preview = pd.read_csv(file, header=None, nrows=10)
    file.seek(0)
    for i, row in df_preview.iterrows():
        if {"Employee", "Rate", "Net Pay"}.issubset(set(row.dropna().astype(str))):
            return i
    return None

# Generates a single payslip PDF
def generate_payslip(employee_name, details, pay_period, logo_path):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.utils import ImageReader

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=30, bottomMargin=50)
    elements = []

    styles = getSampleStyleSheet()
    header_style = ParagraphStyle(name="HeaderStyle", fontSize=12, alignment=0, spaceAfter=2)
    right_aligned_style = ParagraphStyle(name="RightAlign", fontSize=12, alignment=2, spaceAfter=2)
    footer_style = ParagraphStyle(name="FooterStyle", fontSize=9, alignment=1, textColor=colors.black)

    # Logo
    try:
        logo = ImageReader(logo_path)
        iw, ih = logo.getSize()
        aspect_ratio = iw / ih
        logo_image = Image(logo_path, width=aspect_ratio * 60, height=60)
        elements.append(logo_image)
    except:
        elements.append(Paragraph("Company Logo Not Found", header_style))

    elements.append(Spacer(1, 40))

    # Header Table
    employee_text = Paragraph(f"<b>Employee:</b> {employee_name}", header_style)
    rate_text = Paragraph(f"<b>Rate:</b> {details.get('Rate', 'N/A')}", header_style)
    pay_period_text = Paragraph(f"<b>Period:</b> {pay_period}", right_aligned_style)
    payslip_title = Paragraph("<b>PAYSLIP</b>", right_aligned_style)
    elements.append(Table([[employee_text, payslip_title], [rate_text, pay_period_text]], colWidths=[260, 260]))
    elements.append(Spacer(1, 60))

    # Payroll Details Table
    data = [["Category", "Amount"]]
    for key, value in details.items():
        if key not in ["Employee", "Rate", "Net Pay"] and not pd.isna(value) and value != 0:
            data.append([key, f"${value:.2f}"])
    data += [["", ""], ["", ""], ["Net Pay", f"${details.get('Net Pay', 0):.2f}"]]

    table = Table(data, colWidths=[370, 150])
    table_styles = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4167B1")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTNAME", (-2, -1), (-1, -1), "Helvetica-Bold"),
    ]
    for i in range(1, len(data) - 1):
        bg = colors.whitesmoke if i % 2 == 0 else colors.lightgrey
        table_styles.append(("BACKGROUND", (0, i), (-1, i), bg))
    table.setStyle(TableStyle(table_styles))
    elements.append(table)

    # Footer spacing
    elements.append(Spacer(1, 380 - (len(data) * 10)))

    # Footer Line
    footer_line = Table([[Paragraph("", footer_style)]], colWidths=[600])
    footer_line.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, -1), 1.5, colors.HexColor("#4167B1"))]))
    elements.append(footer_line)
    elements.append(Spacer(1, 5))

    # Footer Contact
    footer_data = [[
        Paragraph("https://musicadmin.com/", footer_style),
        Paragraph("hello@musicadmin.com", footer_style),
        Paragraph("615-200-0122", footer_style)
    ]]
    footer_table = Table(footer_data, colWidths=[173, 173, 173], rowHeights=15)
    footer_table.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("TEXTCOLOR", (0, 0), (-1, -1), colors.black)]))
    elements.append(footer_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# Generates ZIP of all payslips
def generate_zip(df, pay_period, logo_url):
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for _, row in df.iterrows():
            emp_name = row["Employee"]
            pdf = generate_payslip(emp_name, row.to_dict(), pay_period, logo_url)
            zip_file.writestr(f"Payslip_{emp_name}.pdf", pdf.read())
    zip_buffer.seek(0)
    return zip_buffer

# --- Streamlit UI ---
st.title("Employee Payslip Generator")
uploaded_file = st.file_uploader("Upload Payroll CSV", type=["csv"])

if uploaded_file:
    logo_url = "https://raw.githubusercontent.com/Music-Admin/mini-tools/refs/heads/main/streamlit-apps/logo-large.png"
    buffer = BytesIO(uploaded_file.read())

    try:
        pay_period = extract_pay_period(BytesIO(buffer.getvalue()))
        header_row = find_header_row(BytesIO(buffer.getvalue()))
        df = pd.read_csv(BytesIO(buffer.getvalue()), header=header_row)
    except Exception as e:
        st.error(f"Error parsing file: {e}")
        st.stop()

    if not {"Employee", "Rate", "Net Pay"}.issubset(df.columns):
        st.error("CSV must contain Employee, Rate, and Net Pay columns.")
    else:
        st.success(f"Found {len(df)} employees. Pay Period: {pay_period}")
        if st.button("Generate Payslips ZIP"):
            zip_file = generate_zip(df, pay_period, logo_url)
            st.download_button("Download All Payslips (ZIP)", zip_file, "Payslips.zip", "application/zip")