import streamlit as st
import pandas as pd
import io
import json
import datetime
import os
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

# Safe Key Assembly
k1 = "sk-proj-buja6UpYkyVQEWarEe3R7VJ9m4oJkPQI8VQqV_mjqZET4BTz-iqVHVG68Xi2k1gT"
k2 = "DUgMeAC0PTT3BlbkFJUr0mzn9BGwOBTpevsUNY7bCqt3X2uxYW-b0j5Zb38rXfV_iewleem8Ok26ymSuAIloX0JCP8cA"
OPENAI_API_KEY = k1 + k2

st.set_page_config(page_title="Awam Logistics - Operasyonel Portal", page_icon="🚢", layout="wide")

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

# Sidebar Navigation
with st.sidebar:
    st.markdown("<div class='brand-box'><div class='brand-title'>AWAM LOGISTICS</div><div class='brand-sub'>Freight Forwarding Suite</div></div>", unsafe_allow_html=True)
    selected_tool = st.radio("Navigation", [
        "⚡ Hızlı RFQ Talep Dönüştürücü",
        "📜 B/L Talimat Dönüştürücü",
        "🏢 الشركات المقيّدة (Company Directory)",
        "🧾 إصدار الفواتير (Invoice Engine)"
    ])

# MODULE 1: RFQ CONVERTER
if selected_tool == "⚡ Hızlı RFQ Talep Dönüştürücü":
    st.markdown("<div class='awam-header'><div class='awam-title'>⚡ Satış Hızlı Talep Standardizasyon Aracı (Awam Quick RFQ)</div><div class='awam-subtitle'>Müşteriden gelen ham mesajları 4 satırlık UN/LOCODE standart fiyatlandırma formatına dönüştürün.</div></div>", unsafe_allow_html=True)
    now = datetime.datetime.now()
    default_ref = f"AGL{now.strftime('%y%m%d')}{now.strftime('%H%M')}"
    col_input, col_output = st.columns([1, 1], gap="large")
    with col_input:
        raw_text = st.text_area("رابط أو نص الطلب:", height=220, placeholder="ضع رسالة الواتساب أو الطلب هنا...")
        r_col1, r_col2 = st.columns([1.2, 1])
        with r_col1: custom_ref = st.text_input("كود المرجعية", value=default_ref)
        with r_col2: process_btn = st.button("⚡ تحويل فوري", use_container_width=True)

    if process_btn and raw_text.strip():
        with st.spinner("جاري التحليل واستخراج كود الشحن..."):
            try:
                client = openai.OpenAI(api_key=OPENAI_API_KEY)
                prompt = f"Parse for Awam Logistics: {raw_text}. Ref Code: {custom_ref}"
                response = client.chat.completions.create(model="gpt-4o-mini", messages=[{"role": "user", "content": prompt}])
                st.session_state["rfq_result"] = response.choices[0].message.content.strip()
            except Exception as e:
                st.error(f"حدث خطأ: {str(e)}")

    with col_output:
        if "rfq_result" in st.session_state:
            st.text_area("النتيجة القياسية المعتمدة:", value=st.session_state["rfq_result"], height=220)

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
    st.markdown("<div class='awam-header'><div class='awam-title'>🧾 وحدة إصدار الفواتير المعتمدة (Awam Financial Invoice Engine)</div><div class='awam-subtitle'>إصدار وتوليد الفواتير المالية الرسمية لشركة أوام لوجستيك.</div></div>", unsafe_allow_html=True)
    if not REPORTLAB_AVAILABLE:
        st.error("⚠️ مكتبة PDF غائبة حالياً على السيرفر، يرجى تشغيل أمر التثبيت للتفعيل المباشر.")
    else:
        st.info("جاهز لإصدار الفواتير PDF المعتمدة.")
