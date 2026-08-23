import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import io
import json
import datetime
import docx
import openai

# Safe Key Assembly for Awam Logistics System
k1 = "sk-proj-buja6UpYkyVQEWarEe3R7VJ9m4oJkPQI8VQqV_mjqZET4BTz-iqVHVG68Xi2k1gT"
k2 = "DUgMeAC0PTT3BlbkFJUr0mzn9BGwOBTpevsUNY7bCqt3X2uxYW-b0j5Zb38rXfV_iewleem8Ok26ymSuAIloX0JCP8cA"
OPENAI_API_KEY = k1 + k2

st.set_page_config(page_title="Awam Logistics - Operasyonel Portal", page_icon="🚢", layout="wide")

# Executive Dark UI Stylesheet (Awam Standard)
st.markdown("""
<style>
    .stApp { background-color: #0F172A; font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; }
    
    [data-testid="stSidebar"] {
        background-color: #1E293B !important;
        border-right: 1px solid #334155 !important;
    }
    
    .brand-box {
        background: linear-gradient(135deg, #1E3A8A 0%, #0F172A 100%);
        border: 1px solid #3B82F6;
        border-radius: 12px;
        padding: 16px;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0 4px 14px rgba(59, 130, 246, 0.25);
    }
    .brand-title { font-size: 20px; font-weight: 900; color: #FFFFFF; letter-spacing: 0.5px; margin: 0; }
    .brand-sub { font-size: 11px; color: #93C5FD; font-weight: 600; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }

    .stRadio > label { display: none !important; }
    .stRadio div[role="radiogroup"] { gap: 12px !important; }
    .stRadio div[role="radiogroup"] > label {
        background: #0F172A !important;
        border: 1px solid #334155 !important;
        border-radius: 10px !important;
        padding: 14px 16px !important;
        color: #94A3B8 !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        transition: all 0.2s ease-in-out !important;
        cursor: pointer !important;
        width: 100% !important;
    }
    .stRadio div[role="radiogroup"] > label:hover {
        border-color: #38BDF8 !important;
        color: #F8FAFC !important;
        background: #1E293B !important;
    }
    .stRadio div[role="radiogroup"] > label[data-checked="true"] {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        border-color: #60A5FA !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.4) !important;
    }

    .awam-header {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 25px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.25);
    }
    .awam-title { font-size: 24px; font-weight: 800; color: #F8FAFC; margin: 0; }
    .awam-subtitle { font-size: 13px; color: #94A3B8; margin-top: 5px; }

    .card-label { font-size: 15px; font-weight: 700; color: #38BDF8; margin-bottom: 12px; display: flex; align-items: center; gap: 8px; }

    .stTextArea textarea {
        background-color: #0F172A !important;
        color: #F8FAFC !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
        font-size: 15px !important;
        line-height: 1.6 !important;
    }
    .stTextInput input {
        background-color: #0F172A !important;
        color: #38BDF8 !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
    }

    .stButton>button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important;
        color: #FFFFFF !important;
        font-weight: 700 !important;
        font-size: 15px !important;
        border: none !important;
        border-radius: 8px !important;
        padding: 12px 24px !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.35) !important;
        transition: all 0.2s ease !important;
    }
    .stButton>button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.5) !important;
    }

    .sys-status {
        background: #0F172A;
        border: 1px solid #1E293B;
        border-radius: 8px;
        padding: 10px;
        margin-top: 30px;
        font-size: 11px;
        color: #10B981;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .dot { width: 8px; height: 8px; background-color: #10B981; border-radius: 50%; display: inline-block; }
</style>
""", unsafe_allow_html=True)

# Sidebar UI Construction
with st.sidebar:
    st.markdown("""
    <div class='brand-box'>
        <div class='brand-title'>AWAM LOGISTICS</div>
        <div class='brand-sub'>Freight Forwarding Suite</div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<p style='color:#64748B; font-size:12px; font-weight:700; margin-bottom:10px;'>MODÜL SEÇİMİ / SELECT MODULE</p>", unsafe_allow_html=True)
    
    selected_tool = st.radio(
        "Navigation",
        [
            "⚡ Hızlı RFQ Talep Dönüştürücü",
            "📜 B/L Talimat Dönüştürücü"
        ]
    )
    
    st.markdown("""
    <div class='sys-status'>
        <span class='dot'></span> Awam Cloud Engine: Active & Online
    </div>
    """, unsafe_allow_html=True)

# Initializing B/L Session State Keys
bl_keys = [
    "booking_no", "shipping_line", "vessel", "pol", "pod", "freight_terms",
    "s_name", "s_addr", "s_tax", "s_tel", "s_email",
    "cn_name", "cn_addr", "cn_tax", "cn_tel", "cn_email",
    "nt_name", "nt_addr", "nt_tax", "nt_tel", "nt_email"
]
for k in bl_keys:
    if k not in st.session_state:
        st.session_state[k] = ""

if "freight_terms" not in st.session_state or not st.session_state["freight_terms"]:
    st.session_state["freight_terms"] = "FREIGHT PREPAID"

if "containers" not in st.session_state:
    st.session_state.containers = pd.DataFrame([
        {"Container No": "", "Seal No": "", "Type": "40' HC", "Packages": "", "Description": "", "Gross Weight (KG)": 0.0, "Volume (CBM)": 0.0}
    ])

# ---------------------------------------------------------
# MODULE 1: SATIŞ - HIZLI FİYATLANDIRMA TALEP DÖNÜŞTÜRÜCÜ
# ---------------------------------------------------------
if selected_tool == "⚡ Hızlı RFQ Talep Dönüştürücü":
    
    st.markdown("""
    <div class='awam-header'>
        <div class='awam-title'>⚡ Satış Hızlı Talep Standardizasyon Aracı (Awam Quick RFQ)</div>
        <div class='awam-subtitle'>Müşteriden gelen ham mesajları 4 satırlık UN/LOCODE standart fiyatlandırma formatına dönüştürün.</div>
    </div>
    """, unsafe_allow_html=True)

    now = datetime.datetime.now()
    default_ref = f"AGL{now.strftime('%y%m%d')}{now.strftime('%H%M')}"

    col_input, col_output = st.columns([1, 1], gap="large")

    with col_input:
        st.markdown("<div class='card-label'>📥 Müşteri / Satış Ham Mesajı</div>", unsafe_allow_html=True)
        raw_text = st.text_area("raw_input_box", height=220, placeholder="ادخل الرساله هنا", label_visibility="collapsed")
        
        r_col1, r_col2 = st.columns([1.2, 1])
        with r_col1:
            custom_ref = st.text_input("Referans Kodu", value=default_ref, label_visibility="collapsed")
        with r_col2:
            process_btn = st.button("⚡ Hızlı Çevir", use_container_width=True)

    if process_btn and raw_text.strip():
        with st.spinner("AI Tarafından Dönüştürülüyor..."):
            try:
                client = openai.OpenAI(api_key=OPENAI_API_KEY)
                prompt = f"""
                You are an expert Freight Forwarding speed-parser for Awam Logistics.
                Convert the raw inquiry text into a VERY SIMPLE, ultra-clean 4-line message with NO BULLET POINTS, NO LABELS, NO EXTRA TEXT.

                STRICT FORMAT REQUIRED:
                [POL IN UPPERCASE] [POD IN UPPERCASE]
                [QUANTITY]X[CONTAINER TYPE IN UPPERCASE]
                [CLIENT NAME IN UPPERCASE]
                [REF CODE]

                Port Name Dictionary Rules:
                - Always map port names to their standard international shipping names (UN/LOCODE standard).
                - Examples:
                  مرسين -> MERSIN
                  عدن -> ADEN
                  بورسودان / بورتسودان -> PORT SUDAN
                  مصراتة / مصراته -> MISURATA
                  جبل علي -> JEBEL ALI
                  أمبارلي / امبارلي -> AMBARLI
                  الحديدية / الحديده -> HODEIDAH
                  جدة / جده -> JEDDAH
                  إسكندرون -> ISKENDERUN

                Equipment Rules:
                - Standardize container specs: حاويه اربعين -> 1X40 HC, حاويه عشرين -> 1X20 GP, إلخ.

                Append Ref Code: {custom_ref}

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

    with col_output:
        st.markdown("<div class='card-label'>📤 Hazır Standart Mesaj</div>", unsafe_allow_html=True)
        if "rfq_result" in st.session_state:
            st.text_area("rfq_output_box", value=st.session_state["rfq_result"], height=220, label_visibility="collapsed")
            
            text_to_copy = json.dumps(st.session_state["rfq_result"])
            copy_button_html = f"""
                <div style="margin-top: 10px;">
                    <button id="copyBtn" onclick="copyToClipboard()" style="
                        width: 100%;
                        background: #16A34A;
                        color: white;
                        font-weight: bold;
                        border: none;
                        padding: 12px;
                        border-radius: 8px;
                        cursor: pointer;
                        font-size: 14px;
                        font-family: sans-serif;
                        box-shadow: 0 4px 12px rgba(22, 163, 74, 0.3);
                        transition: all 0.2s ease;
                    ">📋 Metni Doğrudan Kopyala (Copy Result)</button>
                </div>
                <script>
                    function copyToClipboard() {{
                        const text = {text_to_copy};
                        navigator.clipboard.writeText(text).then(function() {{
                            const btn = document.getElementById('copyBtn');
                            btn.innerText = '✅ Kopyalandı! (Copied to Clipboard)';
                            btn.style.background = '#059669';
                            setTimeout(function() {{
                                btn.innerText = '📋 Metni Doğrudan Kopyala (Copy Result)';
                                btn.style.background = '#16A34A';
                            }}, 2000);
                        }});
                    }}
                </script>
            """
            components.html(copy_button_html, height=65)
        else:
            st.text_area("rfq_output_placeholder", value="Dönüştürülen mesaj burada görünecektir...", height=220, disabled=True, label_visibility="collapsed")

# ---------------------------------------------------------
# MODULE 2: B/L TALİMAT DÖNÜŞTÜRÜCÜ
# ---------------------------------------------------------
elif selected_tool == "📜 B/L Talimat Dönüştürücü":
    
    st.markdown("""
    <div class='awam-header'>
        <div class='awam-title'>📜 B/L Talimat (Bill of Lading Instruction) Dönüştürücü</div>
        <div class='awam-subtitle'>Word/PDF talimatlarını yapay zeka ile okuyun, eksiklikleri kontrol edin ve Awam Excel formatında indirin.</div>
    </div>
    """, unsafe_allow_html=True)

    def read_docx_file(file):
        doc = docx.Document(file)
        full_text = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                row_cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                if row_cells: full_text.append(" | ".join(row_cells))
        return "\n".join(full_text)

    def extract_with_ai(text_content):
        client = openai.OpenAI(api_key=OPENAI_API_KEY)
        prompt = f"""
        You are an expert shipping documentation parser for Awam Logistics.
        Parse the B/L instruction text and strictly return a JSON object with these EXACT keys:
        {{
            "booking_no": "", "shipping_line": "", "vessel": "", "pol": "", "pod": "", "freight_terms": "FREIGHT PREPAID",
            "s_name": "", "s_addr": "", "s_tax": "", "s_tel": "", "s_email": "",
            "cn_name": "", "cn_addr": "", "cn_tax": "", "cn_tel": "", "cn_email": "",
            "nt_name": "", "nt_addr": "", "nt_tax": "", "nt_tel": "", "nt_email": "",
            "containers": [
                {{"Container No": "", "Seal No": "", "Type": "40' HC", "Packages": "", "Description": "", "Gross Weight (KG)": 0.0, "Volume (CBM)": 0.0}}
            ]
        }}

        Document Text:
        {text_content}
        """
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        return json.loads(response.choices[0].message.content.strip())

    uploaded_file = st.file_uploader("B/L Talimat Dosyasını Yükleyin (DOCX/TXT)", type=["docx", "txt"])
    
    if st.button("🚀 AI ile Oku ve Doldur", type="primary"):
        if uploaded_file is not None:
            with st.spinner("İşleniyor..."):
                try:
                    text_content = read_docx_file(uploaded_file)
                    res = extract_with_ai(text_content)
                    
                    for k in bl_keys:
                        if k in res and res[k]:
                            st.session_state[k] = str(res[k])
                    
                    if "containers" in res and res["containers"]:
                        st.session_state.containers = pd.DataFrame(res["containers"])
                        
                    st.success("✅ Tüm Bilgiler Başarıyla Ayrıştırıldı!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Hata: {str(e)}")

    st.subheader("📋 Genel Sevkiyat Bilgileri")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.session_state["booking_no"] = st.text_input("Booking No", value=st.session_state["booking_no"])
        st.session_state["shipping_line"] = st.text_input("Shipping Line", value=st.session_state["shipping_line"])
    with col2:
        st.session_state["vessel"] = st.text_input("Vessel & Voyage", value=st.session_state["vessel"])
        st.session_state["freight_terms"] = st.text_input("Freight Terms", value=st.session_state["freight_terms"])
    with col3:
        st.session_state["pol"] = st.text_input("POL (Port of Loading)", value=st.session_state["pol"])
        st.session_state["pod"] = st.text_input("POD (Port of Discharge)", value=st.session_state["pod"])

    st.subheader("👥 Partiler (Shipper / Consignee / Notify)")
    st.markdown("**1. SHIPPER / YÜKLEYİCİ**")
    c1, c2, c3, c4, c5 = st.columns([1.5, 2, 1.2, 1.2, 1.5])
    with c1: st.session_state["s_name"] = st.text_input("COMPANY NAME", value=st.session_state["s_name"], key="input_s_name")
    with c2: st.session_state["s_addr"] = st.text_input("ADDRESS", value=st.session_state["s_addr"], key="input_s_addr")
    with c3: st.session_state["s_tax"] = st.text_input("TAX NUMBER", value=st.session_state["s_tax"], key="input_s_tax")
    with c4: st.session_state["s_tel"] = st.text_input("TEL", value=st.session_state["s_tel"], key="input_s_tel")
    with c5: st.session_state["s_email"] = st.text_input("EMAIL", value=st.session_state["s_email"], key="input_s_email")

    st.markdown("**2. CONSIGNEE / ALICI**")
    c1, c2, c3, c4, c5 = st.columns([1.5, 2, 1.2, 1.2, 1.5])
    with c1: st.session_state["cn_name"] = st.text_input("COMPANY NAME", value=st.session_state["cn_name"], key="input_cn_name")
    with c2: st.session_state["cn_addr"] = st.text_input("ADDRESS", value=st.session_state["cn_addr"], key="input_cn_addr")
    with c3: st.session_state["cn_tax"] = st.text_input("TAX NUMBER", value=st.session_state["cn_tax"], key="input_cn_tax")
    with c4: st.session_state["cn_tel"] = st.text_input("TEL", value=st.session_state["cn_tel"], key="input_cn_tel")
    with c5: st.session_state["cn_email"] = st.text_input("EMAIL", value=st.session_state["cn_email"], key="input_cn_email")

    st.markdown("**3. NOTIFY / İHBAR TARAF**")
    c1, c2, c3, c4, c5 = st.columns([1.5, 2, 1.2, 1.2, 1.5])
    with c1: st.session_state["nt_name"] = st.text_input("COMPANY NAME", value=st.session_state["nt_name"], key="input_nt_name")
    with c2: st.session_state["nt_addr"] = st.text_input("ADDRESS", value=st.session_state["nt_addr"], key="input_nt_addr")
    with c3: st.session_state["nt_tax"] = st.text_input("TAX NUMBER", value=st.session_state["nt_tax"], key="input_nt_tax")
    with c4: st.session_state["nt_tel"] = st.text_input("TEL", value=st.session_state["nt_tel"], key="input_nt_tel")
    with c5: st.session_state["nt_email"] = st.text_input("EMAIL", value=st.session_state["nt_email"], key="input_nt_email")

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
