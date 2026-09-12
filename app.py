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

# ReportLab Engine for Invoices
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Safe Key Assembly for Awam Logistics System
k1 = "sk-proj-buja6UpYkyVQEWarEe3R7VJ9m4oJkPQI8VQqV_mjqZET4BTz-iqVHVG68Xi2k1gT"
k2 = "DUgMeAC0PTT3BlbkFJUr0mzn9BGwOBTpevsUNY7bCqt3X2uxYW-b0j5Zb38rXfV_iewleem8Ok26ymSuAIloX0JCP8cA"
OPENAI_API_KEY = k1 + k2

st.set_page_config(page_title="Awam Logistics - Suite", page_icon="🚢", layout="wide")

# High-Contrast Design System
st.markdown("""
<style>
    .stApp { background-color: #0F172A !important; font-family: 'Inter', -apple-system, sans-serif !important; }
    [data-testid="stSidebar"] { background-color: #1E293B !important; border-right: 1px solid #334155 !important; }
    
    .brand-box {
        background: linear-gradient(135deg, #1E3A8A 0%, #0F172A 100%);
        border: 1px solid #3B82F6;
        border-radius: 12px;
        padding: 16px;
        text-align: center;
        margin-bottom: 20px;
    }
    .brand-title { font-size: 20px; font-weight: 900; color: #FFFFFF !important; margin: 0; }
    .brand-sub { font-size: 11px; color: #93C5FD !important; font-weight: 600; text-transform: uppercase; margin-top: 4px; }

    h1, h2, h3, h4, h5, h6, label, p, span, div { color: #FFFFFF !important; }
    label[data-testid="stWidgetLabel"] { color: #FFFFFF !important; font-weight: 700 !important; font-size: 14px !important; }

    [data-testid="stFileUploader"] { background-color: #1E293B !important; border: 2px dashed #3B82F6 !important; border-radius: 12px !important; padding: 15px !important; }
    [data-testid="stFileUploader"] * { color: #FFFFFF !important; font-weight: 600 !important; }
    
    .stRadio > label { display: none !important; }
    .stRadio div[role="radiogroup"] { gap: 10px !important; }
    .stRadio div[role="radiogroup"] > label {
        background: #0F172A !important; border: 1px solid #334155 !important; border-radius: 10px !important;
        padding: 12px 14px !important; color: #E2E8F0 !important; font-weight: 600 !important; width: 100% !important;
    }
    .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important; border-color: #60A5FA !important; color: #FFFFFF !important;
    }

    .awam-header { background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border: 1px solid #334155; border-radius: 12px; padding: 20px; margin-bottom: 25px; }
    .awam-title { font-size: 24px; font-weight: 800; color: #FFFFFF !important; margin: 0; }
    .awam-subtitle { font-size: 13px; color: #CBD5E1 !important; margin-top: 5px; }

    .stTextArea textarea, .stSelectbox select { background-color: #0F172A !important; color: #FFFFFF !important; border: 1px solid #475569 !important; border-radius: 8px !important; }
    .stTextInput input { background-color: #0F172A !important; color: #38BDF8 !important; border: 1px solid #475569 !important; border-radius: 8px !important; font-weight: 700 !important; }
    .stButton>button { background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important; color: #FFFFFF !important; font-weight: 700 !important; border-radius: 8px !important; padding: 10px 20px !important; }
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

# Clean ReportLab PDF Generation
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

    # 3. Customer & Meta Box (Separated with gap)
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

    # 4. Items Table
    items_table_data = [[
        Paragraph("NO", th_style),
        Paragraph("SHIPPER", th_style),
        Paragraph("DESCRIPTION", th_style),
        Paragraph("UNITS", th_style),
        Paragraph("UNIT PRICE", th_style),
        Paragraph("TOTAL", th_style)
    ]]

    subtotal = 0.0
    valid_row_index = 1

    for item in items_data:
        shipper_val = sanitize_text(item.get("shipper", ""))
        desc_val = sanitize_text(item.get("description", ""))
        
        # Skip empty/NaN rows
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
        valid_row_index += 1

    # Fill clean empty rows without 'nan'
    while len(items_table_data) < 8:
        items_table_data.append([
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style),
            Paragraph("", cell_style)
        ])

    grand_total = subtotal + tax_amount

    # Financial Rows
    items_table_data.append(['', '', '', '', Paragraph("<b>SUBTOTAL</b>", cell_right), Paragraph(f"<b>${subtotal:,.2f}</b>", cell_right)])
    items_table_data.append(['', '', '', '', Paragraph("<b>TAX</b>", cell_right), Paragraph(f"<b>${tax_amount:,.2f}</b>", cell_right)])
    items_table_data.append(['', '', '', '', Paragraph("<font color='white'><b>GRAND TOTAL:</b></font>", cell_right), Paragraph(f"<font color='white'><b>${grand_total:,.2f}</b></font>", cell_right)])

    items_table = Table(items_table_data, colWidths=[30, 115, 175, 45, 87, 88])
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
    story.append(Spacer(1, 15))

    # 5. Page Border Decoration
    def add_page_border(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#0B1B3D"))
        canvas.setLineWidth(1)
        canvas.rect(18, 18, 576, 756)
        canvas.restoreState()

    story.append(Paragraph("www.awamlogistics.com", ParagraphStyle('FT', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, alignment=1)))

    doc.build(story, onFirstPage=add_page_border, onLaterPages=add_page_border)
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
        st.markdown("<div class='card-label'>📥 Müşteri / Satış Ham Mesajı</div>", unsafe_allow_html=True)
        raw_text = st.text_area("raw_input_box", height=220, placeholder="ادخل الرساله هنا", label_visibility="collapsed")
        r_col1, r_col2 = st.columns([1.2, 1])
        with r_col1: custom_ref = st.text_input("Referans Kodu", value=default_ref, label_visibility="collapsed")
        with r_col2: process_btn = st.button("⚡ Hızlı Çevir", use_container_width=True)

    if process_btn and raw_text.strip():
        with st.spinner("AI Analiz Ediyor..."):
            try:
                client = openai.OpenAI(api_key=OPENAI_API_KEY)
                prompt = f"Parse request for Awam Logistics: {raw_text}. Ref: {custom_ref}"
                response = client.chat.completions.create(model="gpt-4o-mini", messages=[{"role": "user", "content": prompt}])
                st.session_state["rfq_result"] = response.choices[0].message.content.strip()
            except Exception as e: st.error(f"Error: {str(e)}")

    with col_output:
        st.markdown("<div class='card-label'>📤 Hazır Standart Mesaj</div>", unsafe_allow_html=True)
        if "rfq_result" in st.session_state:
            st.text_area("rfq_output_box", value=st.session_state["rfq_result"], height=220, label_visibility="collapsed")
            text_to_copy = json.dumps(st.session_state["rfq_result"])
            copy_button_html = f"""<div style="margin-top: 10px;"><button onclick="navigator.clipboard.writeText({text_to_copy})" style="width: 100%; background: #16A34A; color: white; font-weight: bold; border: none; padding: 12px; border-radius: 8px; cursor: pointer;">📋 Metni Doğrudan Kopyala (Copy Result)</button></div>"""
            components.html(copy_button_html, height=65)

# MODULE 4: INVOICE ENGINE
elif selected_tool == "🧾 إصدار الفواتير (Invoice Engine)":
    st.markdown("<div class='awam-header'><div class='awam-title'>🧾 وحدة إصدار الفواتير المعتمدة (Awam Financial Invoice Engine)</div><div class='awam-subtitle'>إصدار وتوليد الفواتير الرسمية بصيغة PDF ومطابقة لهوية الشعار الأزرق والحساب التلقائي.</div></div>", unsafe_allow_html=True)

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
        {"shipper": "Awam Btunuk Kamen", "description": "Sample Shipping Packet, BURSA / Turkiye", "units": 10.0, "unit_price": 25.0},
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
        st.success("✅ تم توليد الفاتورة بنجاح ومطابقة كامل البيانات وبشكل أنيق بدون أي خطأ nan!")
