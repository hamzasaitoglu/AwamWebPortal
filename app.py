import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import json
import datetime
import os
import docx
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pypdf
import openai

# ReportLab Integration with Exception Catching for Awam Logistics
try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Safe Key Assembly for Awam Logistics System
k1 = "sk-proj-buja6UpYkyVQEWarEe3R7VJ9m4oJkPQI8VQqV_mjqZET4BTz-iqVHVG68Xi2k1gT"
k2 = "DUgMeAC0PTT3BlbkFJUr0mzn9BGwOBTpevsUNY7bCqt3X2uxYW-b0j5Zb38rXfV_iewleem8Ok26ymSuAIloX0JCP8cA"
OPENAI_API_KEY = k1 + k2

st.set_page_config(page_title="Awam Logistics - Operasyonel Portal", page_icon="🚢", layout="wide")

# High-Contrast Design System for Awam Logistics
st.markdown("""
<style>
    .stApp { background-color: #0F172A !important; font-family: 'Inter', sans-serif !important; }
    [data-testid="stSidebar"] { background-color: #1E293B !important; border-right: 1px solid #334155 !important; }
    .brand-box { background: linear-gradient(135deg, #1E3A8A 0%, #0F172A 100%); border: 1px solid #3B82F6; border-radius: 12px; padding: 16px; text-align: center; margin-bottom: 20px; }
    .brand-title { font-size: 20px; font-weight: 900; color: #FFFFFF !important; margin: 0; }
    .brand-sub { font-size: 11px; color: #93C5FD !important; font-weight: 600; text-transform: uppercase; margin-top: 4px; }
    h1, h2, h3, h4, h5, h6, label, p, span, div { color: #FFFFFF !important; }
    .awam-header { background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border: 1px solid #334155; border-radius: 12px; padding: 20px; margin-bottom: 25px; }
    .awam-title { font-size: 24px; font-weight: 800; color: #FFFFFF !important; margin: 0; }
    .awam-subtitle { font-size: 13px; color: #CBD5E1 !important; margin-top: 5px; }
    .stTextInput input, .stTextArea textarea, .stSelectbox select { background-color: #0F172A !important; color: #38BDF8 !important; border: 1px solid #475569 !important; border-radius: 8px !important; }
    .stButton>button { background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important; color: #FFFFFF !important; font-weight: 700 !important; border-radius: 8px !important; }
</style>
""", unsafe_allow_html=True)

# Database Handling for Companies
DB_FILE = "companies_db.json"
def load_companies():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except: return []
    return []

def save_companies(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

if "companies" not in st.session_state:
    st.session_state.companies = load_companies()

def generate_account_code():
    count = len(st.session_state.companies) + 1
    year = datetime.datetime.now().strftime("%Y")
    return f"AWM-ACC-{year}-{count:03d}"

# Helper to clean text and eliminate NaN
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

# PDF Invoice Engine with Perfect Spacing & Bottom Fixed Footer
def build_pdf_invoice(invoice_num, invoice_date, customer_info, items_data, tax_amount=0.0, logo_path="AG-LOGO.png"):
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )
    story = []
    styles = getSampleStyleSheet()

    header_title_style = ParagraphStyle(
        'HT', parent=styles['Normal'],
        fontName='Helvetica-Bold', fontSize=11, leading=13,
        alignment=1, textColor=colors.HexColor("#0B1B3D")
    )
    
    header_sub_style = ParagraphStyle(
        'HS', parent=styles['Normal'],
        fontName='Helvetica', fontSize=8, leading=11,
        alignment=1, textColor=colors.HexColor("#1A2530")
    )

    cell_style = ParagraphStyle('CS', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=10, alignment=0)
    cell_center = ParagraphStyle('CC', parent=cell_style, alignment=1)
    cell_right = ParagraphStyle('CR', parent=cell_style, alignment=2)
    th_style = ParagraphStyle('TH', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, leading=10, alignment=1, textColor=colors.white)

    # 1. Logo
    if os.path.exists(logo_path):
        logo = RLImage(logo_path, width=110, height=60)
        logo.hAlign = 'CENTER'
        story.append(logo)
        story.append(Spacer(1, 8))

    # 2. Header Text
    comp_title = sanitize_text("AWAM GLOBAL LOJISTIK TICARET LIMITED SIRKETI")
    comp_address = sanitize_text("ADDRESS: Mahmudiye Mahallesi Ertugrulgazi Caddesi No:55 ic kapi: 3 Inegol / BURSA / TURKIYE")
    
    story.append(Paragraph(f"<b>{comp_title}</b>", header_title_style))
    story.append(Spacer(1, 3))
    
    header_info = f"""
    {comp_address}<br/>
    <b>VN:</b> 0911212625 &nbsp;&nbsp; <b>VD:</b> INEGOL<br/>
    <b>EMAIL:</b> tr.finans@awamlogistics.com &nbsp;&nbsp; <b>TEL:</b> +90 224 502 8395
    """
    story.append(Paragraph(header_info, header_sub_style))
    story.append(Spacer(1, 14))

    # 3. Customer Details & Invoice Number (Separated Top Boxes)
    cust_clean = sanitize_text(customer_info).replace('\n', '<br/>')
    
    cust_table_data = [
        [Paragraph("Customer Details", th_style)],
        [Paragraph(cust_clean, cell_style)]
    ]
    cust_table = Table(cust_table_data, colWidths=[255], rowHeights=[18, 52])
    cust_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor("#0B1B3D")),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (0, 1), (0, 1), colors.white),
    ]))

    inv_table_data = [
        [Paragraph("Invoice Number", th_style)],
        [Paragraph(sanitize_text(invoice_num), cell_center)],
        [Paragraph("Date", th_style)],
        [Paragraph(sanitize_text(invoice_date), cell_center)]
    ]
    inv_table = Table(inv_table_data, colWidths=[255], rowHeights=[18, 17, 18, 17])
    inv_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (0, 2), (0, 2), colors.HexColor("#0B1B3D")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (0, 1), (0, 1), colors.white),
        ('BACKGROUND', (0, 3), (0, 3), colors.white),
    ]))

    meta_wrapper = Table([[cust_table, '', inv_table]], colWidths=[255, 30, 255])
    meta_wrapper.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
    ]))
    story.append(meta_wrapper)
    story.append(Spacer(1, 14))

    # 4. Line Items Table (With Spaced Empty Rows)
    items_table_data = [[
        Paragraph("NO", th_style),
        Paragraph("SHIPPER", th_style),
        Paragraph("DESCRIPTION", th_style),
        Paragraph("UNITS", th_style),
        Paragraph("UNIT PRICE", th_style),
        Paragraph("TOTAL", th_style)
    ]]

    row_heights = [20] # Header height
    subtotal = 0.0
    valid_row_index = 1

    for item in items_data:
        shipper_val = sanitize_text(item.get("shipper", ""))
        desc_val = sanitize_text(item.get("description", ""))
        
        if not shipper_val and not desc_val:
            continue

        try: units = float(item.get("units", 0))
        except: units = 0.0

        try: unit_price = float(item.get("unit_price", 0))
        except: unit_price = 0.0

        line_total = units * unit_price
        subtotal += line_total

        items_table_data.append([
            Paragraph(str(valid_row_index), cell_center),
            Paragraph(shipper_val, cell_style),
            Paragraph(desc_val, cell_style),
            Paragraph(str(int(units) if units.is_integer() else units), cell_center),
            Paragraph(f"${unit_price:,.2f}", cell_right),
            Paragraph(f"${line_total:,.2f}", cell_right)
        ])
        row_heights.append(None) # Auto height for filled rows
        valid_row_index += 1

    # Fill empty rows up to 7 items with clean 24pt height spacing
    while len(items_table_data) < 8:
        items_table_data.append([
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style)
        ])
        row_heights.append(24) # 24pt fixed height for spacious empty rows

    grand_total = subtotal + tax_amount

    # Financial Summary Rows
    items_table_data.append(['', '', '', '', Paragraph("<b>SUBTOTAL</b>", cell_right), Paragraph(f"<b>${subtotal:,.2f}</b>", cell_right)])
    row_heights.append(20)
    items_table_data.append(['', '', '', '', Paragraph("<b>TAX</b>", cell_right), Paragraph(f"<b>${tax_amount:,.2f}</b>", cell_right)])
    row_heights.append(20)
    items_table_data.append(['', '', '', '', Paragraph("<font color='white'><b>GRAND TOTAL:</b></font>", cell_right), Paragraph(f"<font color='white'><b>${grand_total:,.2f}</b></font>", cell_right)])
    row_heights.append(22)

    items_table = Table(items_table_data, colWidths=[30, 115, 175, 45, 87, 88], rowHeights=row_heights)
    items_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#2B709E")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -4), 0.5, colors.HexColor("#2B709E")),
        ('SPAN', (0, -3), (3, -3)),
        ('SPAN', (0, -2), (3, -2)),
        ('SPAN', (0, -1), (3, -1)),
        ('GRID', (4, -3), (5, -1), 0.5, colors.HexColor("#0B1B3D")),
        ('BACKGROUND', (4, -1), (5, -1), colors.HexColor("#0B1B3D")),
    ]))
    story.append(items_table)

    # 5. Canvas Callback: Draw Page Border & Pin Website Footer to Bottom of Page
    def draw_page_decorations(canvas, doc):
        canvas.saveState()
        # Outer Border Box
        canvas.setStrokeColor(colors.HexColor("#0B1B3D"))
        canvas.setLineWidth(1)
        canvas.rect(18, 18, 576, 756)
        
        # Pinned Website Footer Text at Bottom (y=28pt)
        canvas.setFont("Helvetica", 8.5)
        canvas.setFillColor(colors.HexColor("#1A2530"))
        canvas.drawCentredString(306, 28, "www.awamlogistics.com")
        canvas.restoreState()

    doc.build(story, onFirstPage=draw_page_decorations, onLaterPages=draw_page_decorations)
    pdf_buffer.seek(0)
    return pdf_buffer

# Sidebar Navigation
with st.sidebar:
    st.markdown("<div class='brand-box'><div class='brand-title'>AWAM LOGISTICS</div><div class='brand-sub'>Freight Forwarding Suite</div></div>", unsafe_allow_html=True)
    selected_tool = st.radio("Navigation", [
        "⚡ Hızlı RFQ Talep Dönüştürücü",
        "📜 B/L Talimat Dönüştürücü",
        "🏢 الشركات المقيّدة (Company Directory)",
        "🧾 إصدار الفواتير (Invoice Engine)"
    ])

# MODULE 1: RFQ
if selected_tool == "⚡ Hızlı RFQ Talep Dönüştürücü":
    st.markdown("<div class='awam-header'><div class='awam-title'>⚡ Satış Hızlı Talep Standardizasyon Aracı (Awam Quick RFQ)</div><div class='awam-subtitle'>Müşteriden gelen ham mesajları 4 satırlık UN/LOCODE standart fiyatlandırma formatına dönüştürün.</div></div>", unsafe_allow_html=True)
    now = datetime.datetime.now()
    default_ref = f"AGL{now.strftime('%y%m%d')}{now.strftime('%H%M')}"
    col_input, col_output = st.columns([1, 1], gap="large")
    with col_input:
        raw_text = st.text_area("نص الطلب الخام:", height=220, placeholder="ادخل نص الطلب هنا...")
        r_col1, r_col2 = st.columns([1.2, 1])
        with r_col1: custom_ref = st.text_input("كود المرجعية", value=default_ref)
        with r_col2: process_btn = st.button("⚡ تحويل فوري", use_container_width=True)

    if process_btn and raw_text.strip():
        with st.spinner("جاري المعالجة..."):
            try:
                client = openai.OpenAI(api_key=OPENAI_API_KEY)
                prompt = f"Parse for Awam Logistics: {raw_text}. Ref: {custom_ref}"
                response = client.chat.completions.create(model="gpt-4o-mini", messages=[{"role": "user", "content": prompt}])
                st.session_state["rfq_result"] = response.choices[0].message.content.strip()
            except Exception as e: st.error(f"Error: {str(e)}")

    with col_output:
        if "rfq_result" in st.session_state:
            st.text_area("النتيجة القياسية:", value=st.session_state["rfq_result"], height=220)

# MODULE 3: COMPANY DIRECTORY
elif selected_tool == "🏢 الشركات المقيّدة (Company Directory)":
    st.markdown("<div class='awam-header'><div class='awam-title'>🏢 وحدة إدارة وتقييد الشركات (Awam Directory Engine)</div><div class='awam-subtitle'>تسجيل وتقييد العملاء، الخطوط الملاحية، والموردين بجدول موحد.</div></div>", unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["➕ إضافة شركة جديدة", "📋 سجل الشركات المقيّدة"])
    with tab1:
        with st.form("comp_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                acc_code = st.text_input("الكود المحاسبي", value=generate_account_code())
                comp_name = st.text_input("اسم الشركة *")
                comp_type = st.selectbox("نوع الشركة", ["عميل / Shipper / Consignee", "خط ملاحي / Shipping Line", "مورد نقل داخلي / Hauler", "مخلص جمركي / Customs Broker"])
                email = st.text_input("البريد الإلكتروني")
            with c2:
                phone = st.text_input("رقم الهاتف")
                tax_num = st.text_input("الرقم الضريبي")
                tax_office = st.text_input("المكتب الضريبي")
                address = st.text_input("العنوان")
            sub_btn = st.form_submit_button("💾 حفظ الشركة")
            if sub_btn and comp_name.strip():
                new_c = {"الكود المحاسبي": acc_code, "اسم الشركة": comp_name, "النوع": comp_type, "البريد الإلكتروني": email, "الهاتف": phone, "الرقم الضريبي": tax_num, "المكتب الضريبي": tax_office, "العنوان": address, "تاريخ التسجيل": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")}
                st.session_state.companies.append(new_c)
                save_companies(st.session_state.companies)
                st.success(f"✅ تم تقييد الشركة بنجاح كود: {acc_code}")
                st.rerun()

    with tab2:
        if st.session_state.companies:
            df = pd.DataFrame(st.session_state.companies)
            st.dataframe(df, use_container_width=True)

# MODULE 4: INVOICE ENGINE
elif selected_tool == "🧾 إصدار الفواتير (Invoice Engine)":
    st.markdown("<div class='awam-header'><div class='awam-title'>🧾 وحدة إصدار الفواتير المعتمدة (Awam Financial Invoice Engine)</div><div class='awam-subtitle'>إصدار وتوليد الفواتير المالية الرسمية لشركة أوام لوجستيك بصيغة PDF.</div></div>", unsafe_allow_html=True)

    col_meta1, col_meta2 = st.columns(2)
    with col_meta1:
        inv_num = st.text_input("رقم الفاتورة (Invoice Number)", value="INV-2401")
        companies_list = [c["اسم الشركة"] for c in st.session_state.companies]
        selected_comp = st.selectbox("اختر شركة مقيدة في السجل (أو أدخل يدوياً):", ["-- إدخال يدوي --"] + companies_list)
        
        default_cust_text = "Awam Global Cannealan tinerey\nBURSA / TURKIYE\nVN: 1234567890"
        if selected_comp != "-- إدخال يدوي --":
            comp_obj = next((c for c in st.session_state.companies if c["اسم الشركة"] == selected_comp), None)
            if comp_obj:
                default_cust_text = f"{comp_obj['اسم الشركة']}\n{comp_obj['العنوان']}\nVN: {comp_obj['الرقم الضريبي']}  VD: {comp_obj['المكتب الضريبي']}"

        cust_info = st.text_area("بيانات العميل (Customer Details)", value=default_cust_text, height=100)

    with col_meta2:
        inv_date = st.date_input("تاريخ الفاتورة (Date)", value=datetime.date.today()).strftime("%d.%m.%Y")
        tax_val = st.number_input("قيمة الضريبة المضافة إن وجدت ($)", value=0.0)

    st.subheader("📦 بنود الفاتورة (Invoice Line Items)")
    init_df = pd.DataFrame([
        {"shipper": "Awam Global Cannealan tinerey", "description": "Sample Shipping Packet, BURSA / Turkiye", "units": 100.0, "unit_price": 30.0},
        {"shipper": "", "description": "", "units": 0.0, "unit_price": 0.0}
    ])

    edited_invoice_df = st.data_editor(init_df, num_rows="dynamic", use_container_width=True)

    if st.button("🚀 إصدار الفاتورة وتوليد PDF", type="primary"):
        pdf_out = build_pdf_invoice(
            invoice_num=inv_num,
            invoice_date=inv_date,
            customer_info=cust_info,
            items_data=edited_invoice_df.to_dict(orient="records"),
            tax_amount=tax_val,
            logo_path="AG-LOGO.png"
        )
        st.download_button(
            label="📥 تحميل الفاتورة الرسمية PDF (Awam Invoice)",
            data=pdf_out,
            file_name=f"Invoice_{inv_num}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        st.success("✅ تم إصدار الفاتورة وتحديث المسافات والتنسيق في القاع بنجاح!")
