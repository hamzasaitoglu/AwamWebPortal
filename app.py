import streamlit as st
import pandas as pd
import io
import json
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
import openai

# Safe Key Assembly for Awam Logistics System
k1 = "sk-proj-buja6UpYkyVQEWarEe3R7VJ9m4oJkPQI8VQqV_mjqZET4BTz-iqVHVG68Xi2k1gT"
k2 = "DUgMeAC0PTT3BlbkFJUr0mzn9BGwOBTpevsUNY7bCqt3X2uxYW-b0j5Zb38rXfV_iewleem8Ok26ymSuAIloX0JCP8cA"
OPENAI_API_KEY = k1 + k2

st.set_page_config(page_title="Awam Logistics - Portal", page_icon="🚢", layout="wide")

st.markdown("""
<style>
    .main-header { font-size: 24px; font-weight: bold; color: #1A365D; text-align: center; margin-bottom: 5px; }
    .sub-header { font-size: 14px; color: #4A5568; text-align: center; margin-bottom: 20px; }
    .party-title { font-size: 14px; font-weight: bold; color: #1A365D; margin-top: 10px; margin-bottom: 5px; background: #EDF2F7; padding: 5px 10px; border-radius: 4px; }
    .summary-card { background-color: #EDF2F7; border-left: 5px solid #1A365D; padding: 12px; border-radius: 6px; margin-top: 10px; margin-bottom: 15px; }
    .summary-title { font-size: 13px; font-weight: bold; color: #2D3748; }
    .summary-value { font-size: 18px; font-weight: bold; color: #1A365D; }
    .missing-alert { color: #C53030; font-weight: bold; font-size: 12px; margin-top: 2px; }
    .stButton>button { border-radius: 6px; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# Sidebar Navigation
st.sidebar.title("AWAM LOGISTICS")
st.sidebar.caption("Operasyonel & Satış Portalı")
selected_tool = st.sidebar.radio(
    "Lütfen Kullanmak İstediğiniz Aracı Seçin:",
    [
        "💬 Satış - Fiyatlandırma Talep Dönüştürücü",
        "📜 B/L Talimat Dönüştürücü", 
        "🧹 Metin Karakter Temizleyici", 
        "💰 Navlun Teklif Oluşturucu"
    ]
)

# ---------------------------------------------------------
# TOOL 1: SATIŞ - FİYATLANDIRMA TALEP DÖNÜŞTÜRÜCÜ (AGL RFQ BUILDER)
# ---------------------------------------------------------
if selected_tool == "💬 Satış - Fiyatlandırma Talep Dönüştürücü":
    st.markdown("<div class='main-header'>Satış Talep Standardizasyon Aracı (RFQ Builder)</div>", unsafe_allow_html=True)
    st.markdown("<div class='sub-header'>Müşteriden veya Satış Ekibinden Gelen Düz Metni Standart Awam Fiyatlandırma Formatına Dönüştürün.</div>", unsafe_allow_html=True)

    # Generate Ref Code automatically (e.g. AGL268232304)
    now = datetime.datetime.now()
    default_ref = f"AGL{now.strftime('%y%m%d')}{now.strftime('%H%M')}"

    col_left, col_right = st.columns([1, 1])

    with col_left:
        st.subheader("📥 Müşteriden / Satıştan Gelen Ham Metin")
        raw_text = st.text_area("Metni buraya yapıştırın:", height=250, placeholder="Örn: Selamlar, Ambarlı'dan Aden'e 2x40 HC yükümüz var. Ağırlık 24 ton. Ready date önümüzdeki hafta...")
        
        c_ref, c_btn = st.columns([1.5, 1])
        with c_ref:
            custom_ref = st.text_input("Referans Kodu (Ref Code):", value=default_ref)
        with c_btn:
            st.write(" ")
            st.write(" ")
            process_btn = st.button("⚡ Standart Formata Çevir", type="primary", use_container_width=True)

    if process_btn and raw_text.strip():
        with st.spinner("AI Metni Analiz Ediyor ve Awam Formatına Getiriyor..."):
            try:
                client = openai.OpenAI(api_key=OPENAI_API_KEY)
                prompt = f"""
                You are an expert Freight Forwarding pricing agent at Awam Logistics.
                Parse the following raw inquiry text from sales/client and reformat it into a clean, professional, standardized rate inquiry request for pricing team/shipping line.

                Rules:
                1. Always start with Reference Code: {custom_ref}
                2. Extract POL, POD, Container Type & Quantity, Commodity, Gross Weight, Target Rate (if any), Ready Date, Cargo Type (FCL/LCL/AIR).
                3. Structure in clean bullet points in Turkish or English (match input language).
                4. Add Awam Logistics signature at the bottom.

                Raw Text:
                {raw_text}
                """
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "user", "content": prompt}]
                )
                formatted_result = response.choices[0].message.content.strip()
                st.session_state["rfq_result"] = formatted_result
            except Exception as e:
                st.error(f"Hata Oluştu: {str(e)}")

    with col_right:
        st.subheader("📤 Hazır Fiyatlandırma Mesajı (Awam Format)")
        if "rfq_result" in st.session_state:
            st.text_area("Kopyalamaya Hazır Mesaj:", value=st.session_state["rfq_result"], height=320)
            st.success("✅ Fiyatlandırma ekibine veya hatta gönderilmeye hazır!")

# ---------------------------------------------------------
# TOOL 2: B/L TALİMAT DÖNÜŞTÜRÜCÜ
# ---------------------------------------------------------
elif selected_tool == "📜 B/L Talimat Dönüştürücü":
    
    keys_to_init = [
        "booking_no", "shipping_line", "vessel", "pol", "pod", "freight_terms",
        "s_name", "s_addr", "s_tax", "s_tel", "s_email",
        "cn_name", "cn_addr", "cn_tax", "cn_tel", "cn_email",
        "nt_name", "nt_addr", "nt_tax", "nt_tel", "nt_email"
    ]
    for k in keys_to_init:
        if k not in st.session_state: st.session_state[k] = ""

    if "freight_terms" not in st.session_state or not st.session_state["freight_terms"]:
        st.session_state["freight_terms"] = "FREIGHT PREPAID"

    if "containers" not in st.session_state:
        st.session_state.containers = pd.DataFrame([
            {"Container No": "", "Seal No": "", "Type": "40' HC", "Packages": "", "Description": "", "Gross Weight (KG)": 0.0, "Volume (CBM)": 0.0}
        ])

    def read_docx_file(file):
        doc = docx.Document(file)
        full_text = []
        for p in doc.paragraphs:
            if p.text.strip(): full_text.append(p.text.strip())
        for table in doc.tables:
            for row in table.rows:
                row_cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_cells: full_text.append(" | ".join(row_cells))
        return "\n".join(full_text)

    def extract_with_ai(text_content, api_key):
        client = openai.OpenAI(api_key=api_key)
        prompt = f"""
        You are an expert shipping documentation parser for Awam Logistics.
        Parse the B/L instruction text and split party details into company name, address, tax_id, tel, email.
        Return ONLY a raw JSON object.

        Document Text:
        {text_content}
        """
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content.strip())

    def reset_all():
        for k in keys_to_init: st.session_state[k] = ""
        st.session_state["freight_terms"] = "FREIGHT PREPAID"
        st.session_state.containers = pd.DataFrame([
            {"Container No": "", "Seal No": "", "Type": "40' HC", "Packages": "", "Description": "", "Gross Weight (KG)": 0.0, "Volume (CBM)": 0.0}
        ])

    col_title, col_reset = st.columns([4, 1.2])
    with col_title:
        st.markdown("<div class='main-header'>B/L Talimat Dönüştürücü</div>", unsafe_allow_html=True)
    with col_reset:
        if st.button("🔄 Sıfırla", use_container_width=True):
            reset_all()
            st.rerun()

    uploaded_file = st.file_uploader("B/L Talimat Dosyasını Yükleyin (DOCX/TXT)", type=["docx", "txt"])
    
    if st.button("🚀 AI ile Oku ve Doldur", type="primary"):
        if uploaded_file is not None:
            with st.spinner("İşleniyor..."):
                try:
                    text_content = read_docx_file(uploaded_file)
                    res = extract_with_ai(text_content, OPENAI_API_KEY)
                    for k in keys_to_init:
                        if k in res and res[k]: st.session_state[k] = str(res[k])
                    if "containers" in res and res["containers"]:
                        st.session_state.containers = pd.DataFrame(res["containers"])
                    st.success("✅ Ayrıştırıldı!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Hata: {str(e)}")

    st.subheader("📋 Genel Sevkiyat Bilgileri")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.text_input("Booking No", key="booking_no")
        st.text_input("Shipping Line", key="shipping_line")
    with col2:
        st.text_input("Vessel & Voyage", key="vessel")
        st.text_input("Freight Terms", key="freight_terms")
    with col3:
        st.text_input("POL (Port of Loading)", key="pol")
        st.text_input("POD (Port of Discharge)", key="pod")

    st.subheader("👥 Partiler (Shipper / Consignee / Notify)")
    st.markdown("**1. SHIPPER / YÜKLEYİCİ**")
    c1, c2, c3, c4, c5 = st.columns([1.5, 2, 1.2, 1.2, 1.5])
    with c1: st.text_input("COMPANY NAME", key="s_name")
    with c2: st.text_input("ADDRESS", key="s_addr")
    with c3: st.text_input("TAX NUMBER", key="s_tax")
    with c4: st.text_input("TEL", key="s_tel")
    with c5: st.text_input("EMAIL", key="s_email")

    st.markdown("**2. CONSIGNEE / ALICI**")
    c1, c2, c3, c4, c5 = st.columns([1.5, 2, 1.2, 1.2, 1.5])
    with c1: st.text_input("COMPANY NAME", key="cn_name")
    with c2: st.text_input("ADDRESS", key="cn_addr")
    with c3: st.text_input("TAX NUMBER", key="cn_tax")
    with c4: st.text_input("TEL", key="cn_tel")
    with c5: st.text_input("EMAIL", key="cn_email")

    st.markdown("**3. NOTIFY / İHBAR TARAF**")
    c1, c2, c3, c4, c5 = st.columns([1.5, 2, 1.2, 1.2, 1.5])
    with c1: st.text_input("COMPANY NAME", key="nt_name")
    with c2: st.text_input("ADDRESS", key="nt_addr")
    with c3: st.text_input("TAX NUMBER", key="nt_tax")
    with c4: st.text_input("TEL", key="nt_tel")
    with c5: st.text_input("EMAIL", key="nt_email")

    st.subheader("📦 Konteyner ve Yük Detayları")
    edited_df = st.data_editor(st.session_state.containers, num_rows="dynamic", use_container_width=True)

    total_containers = len(edited_df[edited_df["Container No"].astype(str).str.strip() != ""])
    total_weight = edited_df["Gross Weight (KG)"].apply(lambda x: pd.to_numeric(x, errors='coerce')).sum()
    total_cbm = edited_df["Volume (CBM)"].apply(lambda x: pd.to_numeric(x, errors='coerce')).sum()

    st.markdown("### 📊 Totals Summary")
    s1, s2, s3 = st.columns(3)
    with s1: st.markdown(f"**Containers:** {total_containers}")
    with s2: st.markdown(f"**Total Weight:** {total_weight:,.2f} KG")
    with s3: st.markdown(f"**Total Volume:** {total_cbm:,.2f} CBM")

# Other Placeholder Tools
elif selected_tool == "🧹 Metin Karakter Temizleyici":
    st.subheader("🧹 Metin Karakter Temizleyici")
    st.info("Bu araç yakında eklenecektir.")
elif selected_tool == "💰 Navlun Teklif Oluşturucu":
    st.subheader("💰 Navlun Teklif Oluşturucu")
    st.info("Bu araç yakında eklenecektir.")
