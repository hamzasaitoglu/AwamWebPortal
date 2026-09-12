import streamlit as st
import pandas as pd
import json
import os
import datetime

st.set_page_config(page_title="Awam Logistics - Company Management Module", page_icon="🏢", layout="wide")

DB_FILE = "companies_db.json"

# Load or init database
def load_data():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_data(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

if "companies" not in st.session_state:
    st.session_state.companies = load_data()

# Auto-generate Accounting Code for Awam Logistics
def generate_account_code():
    count = len(st.session_state.companies) + 1
    year = datetime.datetime.now().strftime("%Y")
    return f"AWM-ACC-{year}-{count:03d}"

# Custom CSS High-Contrast
st.markdown("""
<style>
    .stApp { background-color: #0F172A !important; font-family: 'Inter', sans-serif !important; }
    h1, h2, h3, h4, h5, h6, label, p, span, div { color: #FFFFFF !important; }
    .awam-header { background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%); border: 1px solid #334155; border-radius: 12px; padding: 20px; margin-bottom: 25px; }
    .awam-title { font-size: 24px; font-weight: 800; color: #FFFFFF !important; margin: 0; }
    .awam-subtitle { font-size: 13px; color: #CBD5E1 !important; margin-top: 5px; }
    .stTextInput input, .stSelectbox select { background-color: #0F172A !important; color: #38BDF8 !important; border: 1px solid #475569 !important; border-radius: 8px !important; }
    .stButton>button { background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important; color: #FFFFFF !important; font-weight: 700 !important; border-radius: 8px !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class='awam-header'>
    <div class='awam-title'>🏢 وحدة تقييد وتسجيل الشركات (Awam Directory Engine)</div>
    <div class='awam-subtitle'>تسجيل وتقييد العملاء، الخطوط الملاحية، والموردين ومنح الكود المحاسبي الموحد لشركة أوام لوجستيك.</div>
</div>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["➕ إضافة شركة جديدة", "📋 سجل الشركات المقيّدة"])

with tab1:
    st.subheader("تسجيل شركة جديدة في المنظومة")
    with st.form("company_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            acc_code = st.text_input("الكود المحاسبي (تلقائي)", value=generate_account_code())
            comp_name = st.text_input("اسم الشركة (Company Name) *")
            comp_type = st.selectbox("نوع الشركة (Role)", ["عميل / Shipper / Consignee", "خط ملاحي / Shipping Line", "مورد نقل داخلي / Hauler", "مخلص جمركي / Customs Broker"])
            email = st.text_input("البريد الإلكتروني (Email)")
        with c2:
            phone = st.text_input("رقم الهاتف (Phone / Mobile)")
            tax_num = st.text_input("الرقم الضريبي (Tax ID / VKN)")
            tax_office = st.text_input("المكتب الضريبي (Tax Office)")
            address = st.text_input("العنوان التفصيلي (Full Address)")
            
        submit_btn = st.form_submit_button("💾 حفظ وتقييد الشركة")

    if submit_btn:
        if comp_name.strip():
            new_comp = {
                "الكود المحاسبي": acc_code,
                "اسم الشركة": comp_name,
                "النوع": comp_type,
                "البريد الإلكتروني": email,
                "الهاتف": phone,
                "الرقم الضريبي": tax_num,
                "المكتب الضريبي": tax_office,
                "العنوان": address,
                "تاريخ التسجيل": datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            }
            st.session_state.companies.append(new_comp)
            save_data(st.session_state.companies)
            st.success(f"✅ تم حفظ وتأكيد تقييد شركة ({comp_name}) بنجاح بالكود المحاسبي: {acc_code}")
        else:
            st.error("❌ يرجى كتابة اسم الشركة على الأقل لتتم عملية التقييد.")

with tab2:
    st.subheader("قائمة ودليل الشركات المعتمدة لدى أوام لوجستيك")
    if st.session_state.companies:
        df = pd.DataFrame(st.session_state.companies)
        
        # Search Filter
        search_term = st.text_input("🔍 بحث عن شركة، كود محاسبي، أو رقم ضريبي:")
        if search_term:
            df = df[df.apply(lambda row: row.astype(str).str.contains(search_term, case=False).any(), axis=1)]
            
        st.dataframe(df, use_container_width=True)
        
        # Export Option
        csv = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 تحميل دليل الشركات (CSV)", data=csv, file_name="Awam_Logistics_Companies.csv", mime="text/csv")
    else:
        st.info("لا توجد شركات مقيدة حالياً في السجل. استخدم التبويب أعلاه لإضافة شركة.")