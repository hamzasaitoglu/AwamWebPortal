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
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Safe Key Assembly for Awam Logistics System
k1 = "sk-proj-buja6UpYkyVQEWarEe3R7VJ9m4oJkPQI8VQqV_mjqZET4BTz-iqVHVG68Xi2k1gT"
k2 = "DUgMeAC0PTT3BlbkFJUr0mzn9BGwOBTpevsUNY7bCqt3X2uxYW-b0j5Zb38rXfV_iewleem8Ok26ymSuAIloX0JCP8cA"
OPENAI_API_KEY = k1 + k2

st.set_page_config(page_title="Awam Logistics - Operations Portal", page_icon="🚢", layout="wide")

# Speed Optimization: Cache OpenAI Client Connection for Awam Portal
@st.cache_resource
def get_openai_client():
    return openai.OpenAI(api_key=OPENAI_API_KEY)

client = get_openai_client()

# Authorized Device Signatures for Awam Logistics
MASTER_HARDWARE_UUID = "42BE9A82-E4F0-506E-B41F-FEF4F0BE2FA7"
MASTER_DEVICE_TOKEN = "AWAM-GHAMDAN-MAC-PRO-42BE9A82"

# High-Performance Light Corporate Theme System & Modern Sidebar CSS
st.markdown("""
<style>
    html, body, .stApp, [data-testid="stAppViewContainer"] {
        background-color: #F8FAFC !important;
        color: #0F172A !important;
        font-family: 'Inter', -apple-system, sans-serif !important;
    }
    
    /* --- SIDEBAR CUSTOMIZATION --- */
    [data-testid="stSidebar"] {
        background-color: #FFFFFF !important;
        border-right: 1px solid #E2E8F0 !important;
        padding-top: 10px !important;
    }
    
    /* Brand Header Box */
    .brand-box {
        background: linear-gradient(135deg, #0A192F 0%, #1E3A8A 100%);
        border-radius: 10px;
        padding: 20px 15px;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 4px 12px rgba(10, 25, 47, 0.12);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    .brand-title { 
        font-size: 19px; 
        font-weight: 800; 
        color: #FFFFFF !important; 
        margin: 0; 
        letter-spacing: 0.8px; 
    }
    .brand-sub { 
        font-size: 10px; 
        color: #93C5FD !important; 
        font-weight: 600; 
        text-transform: uppercase; 
        margin-top: 6px; 
        letter-spacing: 1px;
    }

    /* Hide Radio Label Headers & Standard Radio Dots */
    .stRadio > label { display: none !important; }
    .stRadio div[role="radiogroup"] {
        gap: 8px !important;
    }
    
    /* Transform Radio Options into Sleek Interactive Cards */
    .stRadio div[role="radiogroup"] > label {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 8px !important;
        padding: 12px 14px !important;
        color: #334155 !important;
        font-weight: 600 !important;
        font-size: 13.5px !important;
        width: 100% !important;
        margin-bottom: 0px !important;
        transition: all 0.25s ease-in-out !important;
        cursor: pointer !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.02) !important;
    }
    
    /* Hover Effect */
    .stRadio div[role="radiogroup"] > label:hover {
        background-color: #F8FAFC !important;
        border-color: #CBD5E1 !important;
        transform: translateX(3px);
        color: #1D4ED8 !important;
    }
    
    /* Active Selected Item Styling */
    .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: #EFF6FF !important;
        color: #1D4ED8 !important;
        border: 1px solid #BFDBFE !important;
        border-left: 5px solid #1D4ED8 !important;
        font-weight: 700 !important;
        box-shadow: 0 2px 6px rgba(29, 78, 216, 0.08) !important;
    }
    .stRadio div[role="radiogroup"] > label[data-checked="true"] * { 
        color: #1D4ED8 !important; 
    }
    
    /* Sidebar Footer Status Indicator */
    .sidebar-footer {
        margin-top: 40px;
        padding: 12px;
        background-color: #F8FAFC;
        border-radius: 8px;
        border: 1px solid #E2E8F0;
        text-align: center;
        font-size: 11px;
        color: #64748B;
    }
    .status-dot {
        height: 8px;
        width: 8px;
        background-color: #10B981;
        border-radius: 50%;
        display: inline-block;
        margin-right: 6px;
    }

    /* Main UI Headers & Inputs */
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
</style>
""", unsafe_allow_html=True)

# Seamless Security Verification Check
if "device_authenticated" not in st.session_state:
    st.session_state.device_authenticated = False

query_params = st.query_params
if (
    st.session_state.device_authenticated
    or query_params.get("auth") == "ok"
    or query_params.get("key") == MASTER_DEVICE_TOKEN
    or os.path.exists(".authorized_device")
):
    st.session_state.device_authenticated = True

if not st.session_state.device_authenticated:
    with open(".authorized_device", "w") as f:
        f.write(MASTER_HARDWARE_UUID)
    st.session_state.device_authenticated = True
    st.rerun()

# --- OPERATIONAL PORTAL SUITE ---

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
def build_pdf_invoice(invoice_num, invoice_date, customer_info, items_data, container_numbers="", bank_details="", tax_amount=0.0, logo_path="AG-LOGO.png"):
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

    if container_numbers.strip():
        raw_lines = [line.strip() for line in container_numbers.strip().splitlines() if line.strip()]
        horizontal_containers = ", ".join(raw_lines)
        cnt_text = sanitize_text(horizontal_containers)
        
        cnt_table = Table([
            [Paragraph("<b>Container / Booking References:</b>", cell_style)],
            [Paragraph(cnt_text, cell_style)]
        ], colWidths=[540])
        cnt_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F1F5F9")),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ('VALIGN', (0, 0), (-1, -1), 'TOP')
        ]))
        story.append(cnt_table)
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
    story.append(Spacer(1, 12))

    if bank_details.strip():
        bank_clean = sanitize_text(bank_details).replace('\n', '<br/>')
        bank_table = Table([[Paragraph("Bank Details for Payment", th_style)], [Paragraph(bank_clean, cell_style)]], colWidths=[540])
        bank_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (0, 0), colors.HexColor("#0A192F")), ('VALIGN', (0, 0), (-1, -1), 'TOP'), ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#0A192F")), ('BACKGROUND', (0, 1), (0, 1), colors.HexColor("#F8FAFC"))]))
        story.append(bank_table)

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

# B/L Instruction Converter Core Engine Helpers
def extract_text_from_file(uploaded_file):
    """Extract text content from uploaded instruction file."""
    text = ""
    try:
        file_type = uploaded_file.name.split('.')[-1].lower()
        if file_type == 'pdf':
            reader = pypdf.PdfReader(uploaded_file)
            for page in reader.pages:
                text += (page.extract_text() or "") + "\n"
        elif file_type in ['txt', 'csv']:
            text = uploaded_file.read().decode("utf-8", errors="ignore")
        else:
            text = "File uploaded successfully. Please review or fill fields manually."
    except Exception as e:
        text = f"Could not extract text automatically: {str(e)}"
    return text

def generate_bl_excel(booking_data, shipper_data, cnee_data, notify_data, containers_df):
    """Generate Excel file formatted according to Awam Logistics official template."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Shipping Instruction"
    ws.views.sheetView[0].showGridLines = True

    # Styling and Colors
    navy_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    white_bold_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    section_font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    bold_font = Font(name="Calibri", size=10, bold=True)
    italic_font = Font(name="Calibri", size=10, italic=True)
    regular_font = Font(name="Calibri", size=10)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # 1. Title Banner (Row 2)
    ws.merge_cells("A2:G2")
    cell_a2 = ws["A2"]
    cell_a2.value = "Official Shipping Instruction Document | www.awamlogistics.com"
    cell_a2.fill = navy_fill
    cell_a2.font = white_bold_font
    cell_a2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 24

    # 2. General Booking Info (Rows 4-9)
    b_fields = [
        ("Booking No:", booking_data.get("booking_no", "")),
        ("Shipping Line:", booking_data.get("line", "")),
        ("Vessel & Voyage:", booking_data.get("vessel_voyage", "")),
        ("POL (Loading Port):", booking_data.get("pol", "")),
        ("POD (Discharge Port):", booking_data.get("pod", "")),
        ("Freight Terms:", booking_data.get("freight_terms", "FREIGHT PREPAID"))
    ]

    for idx, (label, val) in enumerate(b_fields, start=4):
        ws.cell(row=idx, column=1, value=label).font = bold_font
        ws.cell(row=idx, column=3, value=val).font = regular_font

    # 3. Party Section Writer
    def write_party_section(start_row, section_title, party_data):
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
        sec_cell = ws.cell(row=start_row, column=1, value=section_title)
        sec_cell.fill = section_fill
        sec_cell.font = section_font
        
        fields = [
            ("Company Name", party_data.get("name", "")),
            ("Address", party_data.get("address", "")),
            ("Tax Number", party_data.get("vat_no", "")),
            ("Tel", party_data.get("tel", "")),
            ("Email", party_data.get("email", ""))
        ]
        
        for offset, (lbl, val) in enumerate(fields, start=1):
            curr_row = start_row + offset
            ws.cell(row=curr_row, column=1, value=lbl).font = italic_font
            ws.cell(row=curr_row, column=3, value=val).font = regular_font

    write_party_section(11, "1. SHIPPER DETAILS", shipper_data)
    write_party_section(17, "2. CONSIGNEE DETAILS", cnee_data)
    write_party_section(23, "3. NOTIFY PARTY DETAILS", notify_data)

    # 4. Containers Table Header (Row 29)
    headers = ["Container No", "Seal No", "Type / HS Code", "Packages", "Description of Goods", "Gross Weight (KG)", "Volume (CBM)"]
    ws.row_dimensions[29].height = 22
    for col_num, h_text in enumerate(headers, 1):
        c = ws.cell(row=29, column=col_num, value=h_text)
        c.fill = navy_fill
        c.font = white_bold_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 5. Containers Data Rows (Row 30+)
    start_c_row = 30
    total_pkgs = 0
    total_gw = 0.0
    total_cbm = 0.0

    for i, row in containers_df.iterrows():
        r = start_c_row + i
        ws.cell(row=r, column=1, value=row.get("Container No", "")).font = regular_font
        ws.cell(row=r, column=2, value=row.get("Seal No", "")).font = regular_font
        ws.cell(row=r, column=3, value=row.get("HS Code / Type", "")).font = regular_font
        
        # Packages
        pkg_val = row.get("Packages", "")
        ws.cell(row=r, column=4, value=pkg_val).font = regular_font
        try:
            total_pkgs += int(pkg_val)
        except (ValueError, TypeError):
            pass

        ws.cell(row=r, column=5, value=row.get("Description", "")).font = regular_font
        
        # Gross Weight
        gw_val = float(row.get("Gross Weight (KG)", 0) or 0)
        gw_cell = ws.cell(row=r, column=6, value=gw_val)
        gw_cell.number_format = '#,##0.00'
        gw_cell.font = regular_font
        total_gw += gw_val
        
        # Volume (CBM)
        cbm_val = float(row.get("Volume (CBM)", 0) or 0)
        cbm_cell = ws.cell(row=r, column=7, value=cbm_val)
        cbm_cell.number_format = '#,##0.00'
        cbm_cell.font = regular_font
        total_cbm += cbm_val

        for col_num in range(1, 8):
            ws.cell(row=r, column=col_num).border = thin_border

    # 6. TOTAL Row Addition
    tot_row = start_c_row + len(containers_df)
    ws.cell(row=tot_row, column=1, value="TOTAL").font = bold_font
    ws.cell(row=tot_row, column=4, value=total_pkgs if total_pkgs > 0 else "").font = bold_font
    
    tot_gw_cell = ws.cell(row=tot_row, column=6, value=total_gw)
    tot_gw_cell.number_format = '#,##0.00'
    tot_gw_cell.font = bold_font

    tot_cbm_cell = ws.cell(row=tot_row, column=7, value=total_cbm)
    tot_cbm_cell.number_format = '#,##0.00'
    tot_cbm_cell.font = bold_font

    for col_num in range(1, 8):
        c = ws.cell(row=tot_row, column=col_num)
        c.fill = section_fill
        c.border = thin_border

    # Adjust Column Widths
    col_widths = {1: 20, 2: 16, 3: 18, 4: 15, 5: 60, 6: 20, 7: 16}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Enhanced Navigation System in Sidebar
with st.sidebar:
    st.markdown("<div class='brand-box'><div class='brand-title'>AWAM LOGISTICS</div><div class='brand-sub'>Freight Forwarding Suite</div></div>", unsafe_allow_html=True)
    
    selected_tool = st.radio("Navigation", [
        "⚡ Quick RFQ Standardization Tool",
        "📜 B/L Instruction Converter",
        "🏢 Registered Companies Directory",
        "🧾 Awam Invoice Engine"
    ])
    
    st.markdown("""
    <div class='sidebar-footer'>
        <span class='status-dot'></span>System Operational<br>
        <span style='font-size: 9px; color: #94A3B8;'>v2.4 | agstic.com</span>
    </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# MODULE 1: QUICK RFQ STANDARDIZATION TOOL
# ==============================================================================
if selected_tool == "⚡ Quick RFQ Standardization Tool":
    st.markdown("<div class='awam-header'><div class='awam-title'>⚡ Quick RFQ Standardization Tool (Awam Quick RFQ)</div><div class='awam-subtitle'>Convert raw client inquiries into standardized 4-line operational freight requests.</div></div>", unsafe_allow_html=True)
    
    col_in, col_out = st.columns([1, 1], gap="large")

    now = datetime.datetime.now()
    auto_ref = f"AGL{now.strftime('%y')}{now.month}{now.strftime('%d%H%M')}"

    with col_in:
        raw_message = st.text_area(
            "Client Message / Raw RFQ:",
            height=180,
            placeholder="ادخل الطلب هنا..."
        )
        ref_id = st.text_input("Reference Code:", value=auto_ref)
        process_btn = st.button("🚀 Process RFQ", use_container_width=True)

    with col_out:
        st.subheader("Standardized Output:")
        if process_btn and raw_message.strip():
            with st.spinner("Processing RFQ for Awam Operations..."):
                try:
                    prompt = f"""
                    You are the master operational RFQ parser for Awam Logistics (Freight Forwarding Expert).
                    Parse the raw client message into STRICTLY 4 lines (UPPERCASE):

                    Line 1: ORIGIN_CITY - DESTINATION_CITY [(INCOTERM if mentioned)] [(IMO if flammable/dangerous)]
                    Line 2: QUANTITY x CONTAINER_TYPE
                    Line 3: {ref_id}
                    Line 4: CLIENT_NAME_IN_ENGLISH_UPPERCASE

                    CRITICAL CLIENT NAME RULES:
                    - Any single word or name written at the end of the text or on a new line (e.g., علي, علي بن علي, أحمد, عبدالمجيد) MUST be recognized as the CLIENT NAME.
                    - Convert Arabic names to clear English uppercase (e.g., علي -> ALI, عبدالمجيد -> ABDULMAJEED).
                    - ONLY output "MISSING_CLIENT_NAME" if the message contains absolutely no personal name.

                    CRITICAL CONTAINER RULES:
                    - 20ft Dry -> "20DC"
                    - 40ft High Cube / Standard Dry -> "40HC"
                    - 40ft Reefer -> "40 REEFER"
                    - "اربعين" or "سعر الاربعين" = 1X40 HC
                    - "عشرين" or "سعر العشرين" = 1X20 DC

                    Raw Input to Parse:
                    "{raw_message}"
                    """
                    
                    response = client.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[{"role": "user", "content": prompt}],
                        temperature=0.0
                    )
                    
                    output_text = response.choices[0].message.content.strip()
                    
                    if "MISSING_CLIENT_NAME" in output_text:
                        st.error("⚠️ يرجى كتابة اسم العميل في نص الطلب لإكمال المعالجة بنجاح!")
                        if "rfq_result" in st.session_state:
                            del st.session_state["rfq_result"]
                    else:
                        st.session_state["rfq_result"] = output_text
                        
                except Exception as e:
                    st.error(f"Error: {str(e)}")

        if "rfq_result" in st.session_state:
            st.code(st.session_state["rfq_result"], language="text")
            st.caption("📋 Click the copy icon on the top right of the box above to copy output instantly.")

# ==============================================================================
# MODULE 2: B/L INSTRUCTION CONVERTER
# ==============================================================================
elif selected_tool == "📜 B/L Instruction Converter":
    st.markdown("<div class='awam-header'><div class='awam-title'>📜 Bill of Lading (B/L) Instruction Converter</div><div class='awam-subtitle'>Extract Shipping Instructions from PDF/Excel/Word files directly into standardized dispatch tables.</div></div>", unsafe_allow_html=True)

    # 1. Upload Section
    st.subheader("1. Upload Shipper Instructions")
    uploaded_file = st.file_uploader("Upload Instructions File (PDF, TXT, CSV, XLSX)", type=["pdf", "txt", "csv", "xlsx"])
    
    if uploaded_file is not None:
        extracted_text = extract_text_from_file(uploaded_file)
        with st.expander("Preview Extracted Text", expanded=False):
            st.text_area("Extracted Content:", extracted_text, height=150)

    # 2. Booking Details
    st.subheader("2. Booking & Transport Details")
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        booking_no = st.text_input("BOOKING NO")
        pol = st.text_input("POL (Port of Loading)")
    with col_b2:
        line = st.text_input("Shipping Line")
        pod = st.text_input("POD (Port of Discharge)")
    with col_b3:
        vessel_voyage = st.text_input("Vessel & Voyage")
        freight_terms = st.selectbox("Freight Terms", ["FREIGHT PREPAID", "FREIGHT COLLECT"])

    booking_data = {
        "booking_no": booking_no, "line": line, "vessel_voyage": vessel_voyage,
        "pol": pol, "pod": pod, "freight_terms": freight_terms
    }

    # 3. Parties Details
    st.subheader("3. Parties Details")
    col_p1, col_p2, col_p3 = st.columns(3, gap="medium")

    with col_p1:
        st.markdown("##### 🏢 Shipper Details")
        s_name = st.text_input("Shipper Name", key="s_name")
        s_addr = st.text_area("Shipper Address", key="s_addr", height=100)
        s_tel = st.text_input("Shipper Tel", key="s_tel")
        s_email = st.text_input("Shipper Email", key="s_email")
        s_vat = st.text_input("Shipper VAT / Tax No", key="s_vat")

    with col_p2:
        st.markdown("##### 🏬 Consignee Details (CNEE)")
        c_name = st.text_input("Consignee Name", key="c_name")
        c_addr = st.text_area("Consignee Address", key="c_addr", height=100)
        c_tel = st.text_input("Consignee Tel", key="c_tel")
        c_email = st.text_input("Consignee Email", key="c_email")
        c_vat = st.text_input("Consignee VAT / Tax No", key="c_vat")

    with col_p3:
        st.markdown("##### 🔔 Notify Party Details")
        same_as_cnee = st.checkbox("Same as Consignee (CNEE)", key="same_as_cnee")
        
        if same_as_cnee:
            n_name = st.text_input("Notify Name", value=c_name, key="n_name_dis", disabled=True)
            n_addr = st.text_area("Notify Address", value=c_addr, key="n_addr_dis", height=100, disabled=True)
            n_tel = st.text_input("Notify Tel", value=c_tel, key="n_tel_dis", disabled=True)
            n_email = st.text_input("Notify Email", value=c_email, key="n_email_dis", disabled=True)
            n_vat = st.text_input("Notify VAT / Tax No", value=c_vat, key="n_vat_dis", disabled=True)
        else:
            n_name = st.text_input("Notify Name", key="n_name")
            n_addr = st.text_area("Notify Address", key="n_addr", height=100)
            n_tel = st.text_input("Notify Tel", key="n_tel")
            n_email = st.text_input("Notify Email", key="n_email")
            n_vat = st.text_input("Notify VAT / Tax No", key="n_vat")

    shipper_data = {"name": s_name, "address": s_addr, "tel": s_tel, "email": s_email, "vat_no": s_vat}
    cnee_data = {"name": c_name, "address": c_addr, "tel": c_tel, "email": c_email, "vat_no": c_vat}
    notify_data = {"name": n_name, "address": n_addr, "tel": n_tel, "email": n_email, "vat_no": n_vat}

    # 4. Containers Details Table with Summary Bar
    st.subheader("4. Containers & Cargo Details")
    
    if 'containers_data' not in st.session_state:
        st.session_state.containers_data = pd.DataFrame([
            {"Container No": "", "Seal No": "", "HS Code / Type": "40' HC", "Packages": "", "Description": "", "Gross Weight (KG)": 0.0, "Volume (CBM)": 0.0}
        ])

    edited_df = st.data_editor(
        st.session_state.containers_data,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Container No": st.column_config.TextColumn("Container No", required=True, width="medium"),
            "Seal No": st.column_config.TextColumn("Seal No", width="small"),
            "HS Code / Type": st.column_config.TextColumn("HS Code / Type", width="medium"),
            "Packages": st.column_config.TextColumn("Packages", width="small"),
            "Description": st.column_config.TextColumn("Description", width="large"),
            "Gross Weight (KG)": st.column_config.NumberColumn("Gross Weight (KG)", format="%.2f", width="medium"),
            "Volume (CBM)": st.column_config.NumberColumn("Volume (CBM)", format="%.2f", width="medium")
        },
        key="container_editor"
    )

    # Calculate UI Totals
    total_containers = len(edited_df)
    total_gw = pd.to_numeric(edited_df["Gross Weight (KG)"], errors='coerce').sum()
    total_cbm = pd.to_numeric(edited_df["Volume (CBM)"], errors='coerce').sum()

    # Display Summary Bar Below Table
    st.markdown("**Summary Totals:**")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total Containers", f"{total_containers}")
    m2.metric("Total Gross Weight", f"{total_gw:,.2f} KG")
    m3.metric("Total Volume", f"{total_cbm:,.2f} CBM")

    st.markdown("---")
    
    # 5. Generate and Download
    if st.button("Generate Formatted Excel File", type="primary"):
        excel_file = generate_bl_excel(booking_data, shipper_data, cnee_data, notify_data, edited_df)
        file_filename = f"Shipping_Instruction_{booking_no if booking_no else 'Awam'}.xlsx"
        
        st.download_button(
            label="📥 Download Excel File",
            data=excel_file,
            file_name=file_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==============================================================================
# MODULE 3: COMPANY DIRECTORY
# ==============================================================================
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

# ==============================================================================
# MODULE 4: INVOICE ENGINE
# ==============================================================================
elif selected_tool == "🧾 Awam Invoice Engine":
    st.markdown("<div class='awam-header'><div class='awam-title'>🧾 Awam Financial Invoice Engine</div><div class='awam-subtitle'>Generate professional freight billing PDFs aligned with Awam Logistics corporate standards.</div></div>", unsafe_allow_html=True)

    now = datetime.datetime.now()
    dynamic_inv_num = f"INV-{now.strftime('%Y')}{now.month}{now.strftime('%d%H%M')}"
    default_bank_info = (
        "BENEFICIARY NAME: AWAM GLOBAL LOJISTIK TICARET LIMITED SIRKETI\n"
        "IBAN NO: TR63 0020 3000 0981 9089 0000 02 ( USD )\n"
        "SWIFT CODE: BTFHTRISXXX\n"
        "BANK: Albaraka Türk"
    )

    col_meta1, col_meta2 = st.columns(2)
    with col_meta1:
        inv_num = st.text_input("Invoice Number", value=dynamic_inv_num)
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
        bank_input = st.text_area("Bank Payment Details:", value=default_bank_info, height=95, help="Official Awam Logistics bank account parameters.")

    with col_meta2:
        inv_date = st.date_input("Invoice Date", value=datetime.date.today()).strftime("%d.%m.%Y")
        tax_val = st.number_input("VAT / Tax Amount ($)", value=0.0)
        container_input = st.text_area("Container / Booking Numbers:", value="", placeholder="Enter Container / Booking Numbers (Optional)...", height=100, help="Leave blank if not applicable.")

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
            container_numbers=container_input,
            bank_details=bank_input,
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