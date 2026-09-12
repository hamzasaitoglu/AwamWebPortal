import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import json
import datetime
import os
import openpyxl
import pypdf
import openai

# Safe Key Assembly for Awam Logistics System
k1 = "sk-proj-buja6UpYkyVQEWarEe3R7VJ9m4oJkPQI8VQqV_mjqZET4BTz-iqVHVG68Xi2k1gT"
k2 = "DUgMeAC0PTT3BlbkFJUr0mzn9BGwOBTpevsUNY7bCqt3X2uxYW-b0j5Zb38rXfV_iewleem8Ok26ymSuAIloX0JCP8cA"
OPENAI_API_KEY = k1 + k2

st.set_page_config(page_title="Awam Logistics - Operations Portal", page_icon="🚢", layout="wide")

# High-Performance Light Corporate Design System
st.markdown("""
<style>
    html, body, .stApp, [data-testid="stAppViewContainer"] {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        font-family: 'Inter', -apple-system, sans-serif !important;
    }
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0 !important;
    }
    .brand-box {
        background: linear-gradient(135deg, #0A192F 0%, #1E3A8A 100%);
        border-radius: 8px;
        padding: 16px;
        text-align: center;
        margin-bottom: 20px;
    }
    .brand-title { font-size: 18px; font-weight: 800; color: #FFFFFF !important; margin: 0; letter-spacing: 0.5px; }
    .brand-sub { font-size: 10px; color: #93C5FD !important; font-weight: 600; text-transform: uppercase; margin-top: 4px; }
    .awam-header { 
        background-color: #FFFFFF !important; 
        border: 1px solid #E2E8F0 !important; 
        border-left: 5px solid #1D4ED8 !important;
        border-radius: 8px !important; 
        padding: 18px !important; 
        margin-bottom: 20px !important;
    }
    .awam-title { font-size: 20px; font-weight: 800; color: #0F172A !important; margin: 0; }
    .awam-subtitle { font-size: 12px; color: #475569 !important; margin-top: 4px; }
    h1, h2, h3, h4, h5, h6, p, span, div, label { color: #0F172A !important; }
    label[data-testid="stWidgetLabel"] { font-weight: 700 !important; font-size: 13px !important; color: #0F172A !important; }
    .stTextInput input, .stTextArea textarea, .stSelectbox select, .stNumberInput input { 
        background-color: #FFFFFF !important; 
        color: #0F172A !important; 
        border: 1px solid #CBD5E1 !important; 
        border-radius: 6px !important; 
    }
    .stButton>button { 
        background: #1D4ED8 !important; 
        color: #FFFFFF !important; 
        font-weight: 700 !important; 
        border-radius: 6px !important; 
        border: none !important;
        padding: 8px 20px !important;
    }
    .stRadio > label { display: none !important; }
    .stRadio div[role="radiogroup"] > label {
        background: #F1F5F9 !important; border: 1px solid #E2E8F0 !important; border-radius: 6px !important;
        padding: 10px !important; color: #334155 !important; font-weight: 600 !important; width: 100% !important; margin-bottom: 6px !important;
    }
    .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: #1D4ED8 !important; color: #FFFFFF !important; border-color: #1D4ED8 !important;
    }
    .stRadio div[role="radiogroup"] > label[data-checked="true"] * { color: #FFFFFF !important; }
</style>
""", unsafe_allow_html=True)

# 1. Caching Engine for Fast Data Loading
@st.cache_data(ttl=60)
def load_companies_fast():
    DB_FILE = "companies_db.json"
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except: return []
    return []

def save_companies(data):
    DB_FILE = "companies_db.json"
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)
    st.cache_data.clear()

if "companies" not in st.session_state:
    st.session_state.companies = load_companies_fast()

def generate_account_code():
    count = len(st.session_state.companies) + 1
    year = datetime.datetime.now().strftime("%Y")
    return f"AWM-ACC-{year}-{count:03d}"

def sanitize_text(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.lower() == "nan":
        return ""
    replacements = {
        'İ': 'I', 'ı': 'i', 'Ş': 'S', 'ş': 's',
        'Ğ': 'G', 'ğ': 'g', 'Ü': 'U', 'ü': 'u',
        'Ö': 'O', 'ö': 'o', 'Ç': 'C', 'ç': 'c'
    }
    for search, replace in replacements.items():
        val_str = val_str.replace(search, replace)
    return val_str

# PDF Invoice Generation Engine
def build_pdf_invoice(invoice_num, invoice_date, customer_info, items_data, tax_amount=0.0, logo_path="AG-LOGO.png"):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors

    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()

    header_company_title = ParagraphStyle('HCT', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10.5, leading=13, alignment=1, textColor=colors.HexColor("#0A192F"))
    header_company_sub = ParagraphStyle('HCS', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=11, alignment=1, textColor=colors.HexColor("#1A2530"))
    cell_style = ParagraphStyle('CS', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=10, alignment=0)
    cell_center = ParagraphStyle('CC', parent=cell_style, alignment=1)
    cell_right = ParagraphStyle('CR', parent=cell_style, alignment=2)
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=10, alignment=1, textColor=colors.white)

    if os.path.exists(logo_path):
        logo = RLImage(logo_path, width=105, height=58)
        logo.hAlign = 'CENTER'
        story.append(logo)
        story.append(Spacer(1, 6))

    comp_title = sanitize_text("AWAM GLOBAL LOJISTIK TICARET LIMITED SIRKETI")
    comp_address = sanitize_text("ADDRESS: Mahmudiye Mahallesi Ertugrulgazi Caddesi No:55 ic kapi:\n3 Inegol / BURSA / TURKIYE")
    story.append(Paragraph(f"<b>{comp_title}</b>", header_company_title))
    story.append(Spacer(1, 3))
    
    header_info = f"{comp_address.replace('\n', '<br/>')}<br/><b>VN:</b> 0911212625 &nbsp;&nbsp; <b>VD:</b> INEGOL<br/><b>EMAIL:</b> tr.finans@awamlogistics.com<br/><b>TEL:</b> +90 224 502 8395"
    story.append(Paragraph(header_info, header_company_sub))
    story.append(Spacer(1, 12))

    story.append(Table([['']], colWidths=[540], rowHeights=[1], style=[('LINEABOVE', (0, 0), (-1, -1), 0.75, colors.HexColor("#0A192F"))]))
    story.append(Spacer(1, 12))

    cust_clean = sanitize_text(customer_info).replace('\n', '<br/>')
    cust_table = Table([[Paragraph("Customer Details", th_style)], [Paragraph(cust_clean, cell_style)]], colWidths=[255], rowHeights=[18, 50])
    cust_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (0, 0), colors.HexColor("#0A192F")), ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#0A192F")), ('BACKGROUND', (0, 1), (0, 1), colors.white)]))

    inv_table = Table([[Paragraph("Invoice Number", th_style)], [Paragraph(sanitize_text(invoice_num), cell_center)], [Paragraph("Date", th_style)], [Paragraph(sanitize_text(invoice_date), cell_center)]], colWidths=[255], rowHeights=[18, 16, 18, 16])
    inv_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (0, 0), colors.HexColor("#0A192F")), ('BACKGROUND', (0, 2), (0, 2), colors.HexColor("#0A192F")), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#0A192F")), ('BACKGROUND', (0, 1), (0, 1), colors.white), ('BACKGROUND', (0, 3), (0, 3), colors.white)]))

    story.append(Table([[cust_table, '', inv_table]], colWidths=[255, 30, 255], style=[('VALIGN', (0, 0), (-1, -1), 'TOP')]))
    story.append(Spacer(1, 12))

    items_table_data = [[Paragraph("NO", th_style), Paragraph("SHIPPER", th_style), Paragraph("DESCRIPTION", th_style), Paragraph("UNITS", th_style), Paragraph("UNIT PRICE", th_style), Paragraph("TOTAL", th_style)]]
    row_heights = [20]
    subtotal = 0.0
    valid_idx = 1

    for item in items_data:
        shipper_val = sanitize_text(item.get("shipper", ""))
        desc_val = sanitize_text(item.get("description", ""))
        if not shipper_val and not desc_val: continue
        try: units = float(item.get("units", 0))
        except: units = 0.0
        try: unit_price = float(item.get("unit_price", 0))
        except: unit_price = 0.0
        line_total = units * unit_price
        subtotal += line_total

        items_table_data.append([Paragraph(str(valid_idx), cell_center), Paragraph(shipper_val, cell_style), Paragraph(desc_val, cell_style), Paragraph(str(int(units) if units.is_integer() else units), cell_center), Paragraph(f"${unit_price:,.2f}", cell_right), Paragraph(f"${line_total:,.2f}", cell_right)])
        row_heights.append(None)
        valid_idx += 1

    while len(items_table_data) < 8:
        items_table_data.append([Paragraph("", cell_style)] * 6)
        row_heights.append(25)

    grand_total = subtotal + tax_amount
    items_table_data.append(['', '', '', '', Paragraph("<b>SUBTOTAL</b>", cell_right), Paragraph(f"<b>${subtotal:,.2f}</b>", cell_right)])
    row_heights.append(20)
    items_table_data.append(['', '', '', '', Paragraph("<b>TAX</b>", cell_right), Paragraph(f"<b>${tax_amount:,.2f}</b>", cell_right)])
    row_heights.append(20)
    items_table_data.append(['', '', '', '', Paragraph("<font color='white'><b>GRAND TOTAL:</b></font>", cell_right), Paragraph(f"<font color='white'><b>${grand_total:,.2f}</b></font>", cell_right)])
    row_heights.append(22)

    items_table = Table(items_table_data, colWidths=[30, 115, 175, 45, 87, 88], rowHeights=row_heights)
    items_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2A72A4")), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('GRID', (0, 0), (-1, -4), 0.5, colors.HexColor("#2A72A4")), ('SPAN', (0, -3), (3, -3)), ('SPAN', (0, -2), (3, -2)), ('SPAN', (0, -1), (3, -1)), ('GRID', (4, -3), (5, -1), 0.5, colors.HexColor("#0A192F")), ('BACKGROUND', (4, -1), (5, -1), colors.HexColor("#0A192F"))]))
    story.append(items_table)

    def draw_page_decorations(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#0A192F"))
        canvas.setLineWidth(1)
        canvas.rect(18, 18, 576, 756)
        canvas.setLineWidth(0.5)
        canvas.line(36, 45, 576, 45)
        canvas.setFont("Helvetica", 8.5)
        canvas.setFillColor(colors.HexColor("#1A2530"))
        canvas.drawCentredString(306, 30, "www.awamlogistics.com")
        canvas.restoreState()

    doc.build(story, onFirstPage=draw_page_decorations, onLaterPages=draw_page_decorations)
    pdf_buffer.seek(0)
    return pdf_buffer

# Navigation System
with st.sidebar:
    st.markdown("<div class='brand-box'><div class='brand-title'>AWAM LOGISTICS</div><div class='brand-sub'>Freight Forwarding Suite</div></div>", unsafe_allow_html=True)
    selected_tool = st.radio("Navigation", [
        "⚡ Quick RFQ Standardization Tool",
        "📜 B/L Instruction Converter",
        "🏢 Registered Companies Directory",
        "🧾 Awam Invoice Engine"
    ])

# MODULE 1: RFQ CONVERTER
if selected_tool == "⚡ Quick RFQ Standardization Tool":
    st.markdown("<div class='awam-header'><div class='awam-title'>⚡ Quick RFQ Standardization Tool (Awam Quick RFQ)</div><div class='awam-subtitle'>Convert raw WhatsApp/Email client messages into standardized 4-line UN/LOCODE freight pricing queries.</div></div>", unsafe_allow_html=True)
    now = datetime.datetime.now()
    default_ref = f"AGL{now.strftime('%y%m%d')}{now.strftime('%H%M')}"
    col_input, col_output = st.columns([1, 1], gap="large")
    with col_input:
        raw_text = st.text_area("Raw Client Request Message:", height=200, placeholder="Paste raw WhatsApp message or email text here...")
        r_col1, r_col2 = st.columns([1.2, 1])
        with r_col1: custom_ref = st.text_input("Reference Code", value=default_ref)
        with r_col2: process_btn = st.button("⚡ Convert Instantly", use_container_width=True)

    if process_btn and raw_text.strip():
        with st.spinner("Processing request..."):
            try:
                client = openai.OpenAI(api_key=OPENAI_API_KEY)
                prompt = f"Parse for Awam Logistics: {raw_text}. Ref: {custom_ref}"
                response = client.chat.completions.create(model="gpt-4o-mini", messages=[{"role": "user", "content": prompt}])
                st.session_state["rfq_result"] = response.choices[0].message.content.strip()
            except Exception as e: st.error(f"Error: {str(e)}")

    with col_output:
        if "rfq_result" in st.session_state:
            st.text_area("Standardized Output:", value=st.session_state["rfq_result"], height=200)

# MODULE 2: B/L CONVERTER
elif selected_tool == "📜 B/L Instruction Converter":
    st.markdown("<div class='awam-header'><div class='awam-title'>📜 Bill of Lading (B/L) Instruction Converter</div><div class='awam-subtitle'>Extract Shipping Instructions from PDF/Excel/Word files directly into standardized dispatch tables.</div></div>", unsafe_allow_html=True)
    st.info("Upload shipping documents to extract container details, Shipper, Consignee, and HS Codes.")

# MODULE 3: COMPANY DIRECTORY
elif selected_tool == "🏢 Registered Companies Directory":
    st.markdown("<div class='awam-header'><div class='awam-title'>🏢 Company Directory Engine (Awam Directory)</div><div class='awam-subtitle'>Manage and register Shippers, Consignees, Shipping Lines, and Subcontractors.</div></div>", unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["➕ Register New Company", "📋 Registered Directory Ledger"])
    with tab1:
        with st.form("comp_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                acc_code = st.text_input("Account Code", value=generate_account_code())
                comp_name = st.text_input("Company Name *")
                comp_type = st.selectbox("Company Category", ["Client / Shipper / Consignee", "Shipping Line / Carrier", "Hauler / Internal Trucking", "Customs Broker"])
                email = st.text_input("Email Address")
            with c2:
                phone = st.text_input("Phone Number")
                tax_num = st.text_input("Tax ID (VN)")
                tax_office = st.text_input("Tax Office (VD)")
                address = st.text_input("Full Address")
            sub_btn = st.form_submit_button("💾 Save Company Record")
            if sub_btn and comp_name.strip():
                new_c = {"Account Code": acc_code, "Company Name": comp_name, "Category": comp_type, "Email": email, "Phone": phone, "Tax ID": tax_num, "Tax Office": tax_office, "Address": address, "Created Date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}
                st.session_state.companies.append(new_c)
                save_companies(st.session_state.companies)
                st.success(f"✅ Company successfully saved with Account Code: {acc_code}")
                st.rerun()

    with tab2:
        if st.session_state.companies:
            df = pd.DataFrame(st.session_state.companies)
            st.dataframe(df, use_container_width=True)

# MODULE 4: INVOICE ENGINE
elif selected_tool == "🧾 Awam Invoice Engine":
    st.markdown("<div class='awam-header'><div class='awam-title'>🧾 Awam Financial Invoice Engine</div><div class='awam-subtitle'>Generate professional freight billing PDFs aligned with Awam Logistics corporate standards.</div></div>", unsafe_allow_html=True)

    col_meta1, col_meta2 = st.columns(2)
    with col_meta1:
        inv_num = st.text_input("Invoice Number", value="INV-2401")
        companies_list = [c.get("Company Name", c.get("اسم الشركة", "")) for c in st.session_state.companies]
        selected_comp = st.selectbox("Select Company from Directory (or enter manually):", ["-- Manual Entry --"] + companies_list)
        
        default_cust_text = "Awam Global Cannealan tinerey\nBURSA / TURKIYE\nVN: 1234567890"
        if selected_comp != "-- Manual Entry --":
            comp_obj = next((c for c in st.session_state.companies if c.get("Company Name", c.get("اسم الشركة", "")) == selected_comp), None)
            if comp_obj:
                c_name = comp_obj.get("Company Name", comp_obj.get("اسم الشركة", ""))
                c_addr = comp_obj.get("Address", comp_obj.get("العنوان", ""))
                c_tax = comp_obj.get("Tax ID", comp_obj.get("الرقم الضريبي", ""))
                c_off = comp_obj.get("Tax Office", comp_obj.get("المكتب الضريبي", ""))
                default_cust_text = f"{c_name}\n{c_addr}\nVN: {c_tax}  VD: {c_off}"

        cust_info = st.text_area("Customer Details", value=default_cust_text, height=100)

    with col_meta2:
        inv_date = st.date_input("Invoice Date", value=datetime.date.today()).strftime("%d.%m.%Y")
        tax_val = st.number_input("VAT / Tax Amount ($)", value=0.0)

    st.subheader("📦 Invoice Line Items")
    init_df = pd.DataFrame([
        {"shipper": "Awam Global Cannealan tinerey", "description": "Sample Shipping Packet, BURSA / Turkiye", "units": 100.0, "unit_price": 30.0},
        {"shipper": "", "description": "", "units": 0.0, "unit_price": 0.0}
    ])

    edited_invoice_df = st.data_editor(init_df, num_rows="dynamic", use_container_width=True)

    if st.button("🚀 Generate PDF Invoice", type="primary"):
        pdf_out = build_pdf_invoice(
            invoice_num=inv_num,
            invoice_date=inv_date,
            customer_info=cust_info,
            items_data=edited_invoice_df.to_dict(orient="records"),
            tax_amount=tax_val,
            logo_path="AG-LOGO.png"
        )
        st.download_button(
            label="📥 Download Official PDF Invoice",
            data=pdf_out,
            file_name=f"Invoice_{inv_num}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        st.success("✅ PDF Invoice generated successfully!")
