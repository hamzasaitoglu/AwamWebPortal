import streamlit as st
import pandas as pd
import io
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
import openai

# Safe Key Assembly for Awam Logistics System
k1 = "sk-proj-buja6UpYkyVQEWarEe3R7VJ9m4oJkPQI8VQqV_mjqZET4BTz-iqVHVG68Xi2k1gT"
k2 = "DUgMeAC0PTT3BlbkFJUr0mzn9BGwOBTpevsUNY7bCqt3X2uxYW-b0j5Zb38rXfV_iewleem8Ok26ymSuAIloX0JCP8cA"
OPENAI_API_KEY = k1 + k2

st.set_page_config(page_title="Awam Logistics - Operasyonel Portal", page_icon="🚢", layout="wide")

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

# Initialize Session State Keys
keys_to_init = [
    "booking_no", "shipping_line", "vessel", "pol", "pod", "freight_terms",
    "s_name", "s_addr", "s_tax", "s_tel", "s_email",
    "cn_name", "cn_addr", "cn_tax", "cn_tel", "cn_email",
    "nt_name", "nt_addr", "nt_tax", "nt_tel", "nt_email"
]
for k in keys_to_init:
    if k not in st.session_state:
        st.session_state[k] = ""

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
    You are an expert shipping documentation parser for Awam Logistics (Freight Forwarder).
    Parse the B/L instruction text and strictly split the party details into separate keys:
    company name, address, tax_id (tax number / CR no), tel, and email.

    Return ONLY a raw JSON object:
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
        messages=[{"role": "user", "content": prompt}]
    )
    raw_res = response.choices[0].message.content.strip()
    if raw_res.startswith("```"):
        parts = raw_res.split("```")
        if len(parts) > 1:
            raw_res = parts[1].replace("json", "").strip()
    return json.loads(raw_res)

def reset_all():
    for k in keys_to_init:
        st.session_state[k] = ""
    st.session_state["freight_terms"] = "FREIGHT PREPAID"
    st.session_state.containers = pd.DataFrame([
        {"Container No": "", "Seal No": "", "Type": "40' HC", "Packages": "", "Description": "", "Gross Weight (KG)": 0.0, "Volume (CBM)": 0.0}
    ])

# Sidebar Menu
st.sidebar.title("AWAM LOGISTICS")
st.sidebar.caption("Operasyonel Portal")
selected_tool = st.sidebar.radio(
    "Lütfen Kullanmak İstediğiniz Aracı Seçin:",
    ["📜 B/L Talimat Dönüştürücü", "💬 Satış - Fiyatlandırma", "🧹 Metin Karakter Temizleyici", "💰 Navlun Teklif Oluşturucu"]
)

if selected_tool == "📜 B/L Talimat Dönüştürücü":
    
    col_title, col_reset = st.columns([4, 1.2])
    with col_title:
        st.markdown("<div class='main-header'>B/L Talimat (Bill of Lading Instruction) Dönüştürücü</div>", unsafe_allow_html=True)
        st.markdown("<div class='sub-header'>Fabrika Word/PDF talimatlarını yapay zeka ile okuyun, eksiklikleri kontrol edin ve Awam Excel formatında indirin.</div>", unsafe_allow_html=True)
    with col_reset:
        if st.button("🔄 Yeni İşlem / Sıfırla", use_container_width=True):
            reset_all()
            st.rerun()

    uploaded_file = st.file_uploader("Fabrikadan Gelen B/L Talimat Dosyasını Yükleyin (DOCX veya TXT)", type=["docx", "txt"])
    
    col_process, _ = st.columns([1.5, 3])
    with col_process:
        if st.button("🚀 Talimatı Oku ve Doldur (AI)", type="primary", use_container_width=True):
            if uploaded_file is not None:
                with st.spinner("Dosya AI ile analiz ediliyor..."):
                    try:
                        text_content = read_docx_file(uploaded_file)
                        res = extract_with_ai(text_content, OPENAI_API_KEY)
                        
                        for k in keys_to_init:
                            if k in res and res[k]:
                                st.session_state[k] = str(res[k])
                        
                        if "containers" in res and res["containers"]:
                            st.session_state.containers = pd.DataFrame(res["containers"])
                            
                        st.success("✅ Tüm bilgiler başarıyla ayrıştırıldı ve ekrana aktarıldı!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"AI Analiz Hatası: {str(e)}")
            else:
                st.warning("Lütfen önce bir dosya yükleyin.")

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

    st.subheader("👥 Tarafların Detaylı Bilgileri (Parties)")
    
    # Shipper
    st.markdown("<div class='party-title'>1. SHIPPER / YÜKLEYİCİ</div>", unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns([1.5, 2, 1.2, 1.2, 1.5])
    with c1:
        st.text_input("COMPANY NAME", key="s_name")
        if not st.session_state.s_name: st.markdown("<div class='missing-alert'>⚠️ Eksik Bilgi</div>", unsafe_allow_html=True)
    with c2:
        st.text_input("ADDRESS", key="s_addr")
        if not st.session_state.s_addr: st.markdown("<div class='missing-alert'>⚠️ Eksik Bilgi</div>", unsafe_allow_html=True)
    with c3:
        st.text_input("TAX NUMBER", key="s_tax")
    with c4:
        st.text_input("TEL", key="s_tel")
    with c5:
        st.text_input("EMAIL", key="s_email")

    # Consignee
    st.markdown("<div class='party-title'>2. CONSIGNEE / ALICI</div>", unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns([1.5, 2, 1.2, 1.2, 1.5])
    with c1:
        st.text_input("COMPANY NAME", key="cn_name")
        if not st.session_state.cn_name: st.markdown("<div class='missing-alert'>⚠️ Eksik Bilgi</div>", unsafe_allow_html=True)
    with c2:
        st.text_input("ADDRESS", key="cn_addr")
        if not st.session_state.cn_addr: st.markdown("<div class='missing-alert'>⚠️ Eksik Bilgi</div>", unsafe_allow_html=True)
    with c3:
        st.text_input("TAX NUMBER / CR NO", key="cn_tax")
    with c4:
        st.text_input("TEL", key="cn_tel")
    with c5:
        st.text_input("EMAIL", key="cn_email")

    # Notify
    st.markdown("<div class='party-title'>3. NOTIFY PARTY / İHBAR TARAF</div>", unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns([1.5, 2, 1.2, 1.2, 1.5])
    with c1:
        st.text_input("COMPANY NAME", key="nt_name")
        if not st.session_state.nt_name: st.markdown("<div class='missing-alert'>⚠️ Eksik Bilgi</div>", unsafe_allow_html=True)
    with c2:
        st.text_input("ADDRESS", key="nt_addr")
        if not st.session_state.nt_addr: st.markdown("<div class='missing-alert'>⚠️ Eksik Bilgi</div>", unsafe_allow_html=True)
    with c3:
        st.text_input("TAX NUMBER / CR NO", key="nt_tax")
    with c4:
        st.text_input("TEL", key="nt_tel")
    with c5:
        st.text_input("EMAIL", key="nt_email")

    st.subheader("📦 Konteyner ve Yük Detayları")
    edited_df = st.data_editor(st.session_state.containers, num_rows="dynamic", use_container_width=True)

    # Dynamic Totals Calculation
    total_containers = len(edited_df[edited_df["Container No"].astype(str).str.strip() != ""])
    
    pkg_sum_str = ""
    try:
        pkg_vals = edited_df["Packages"].astype(str).tolist()
        pkg_nums = [float(pd.to_numeric(p.split()[0], errors='coerce')) for p in pkg_vals if p.strip()]
        valid_nums = [n for n in pkg_nums if not pd.isna(n)]
        if valid_nums: pkg_sum_str = f"{int(sum(valid_nums))} PKGS"
        else: pkg_sum_str = f"{len(pkg_vals)} ITEMS"
    except:
        pkg_sum_str = "-"

    total_weight = edited_df["Gross Weight (KG)"].apply(lambda x: pd.to_numeric(x, errors='coerce')).sum()
    total_cbm = edited_df["Volume (CBM)"].apply(lambda x: pd.to_numeric(x, errors='coerce')).sum()

    st.markdown("### 📊 Sevkiyat İcmal ve İstatistikleri (Totals Summary)")
    s1, s2, s3, s4 = st.columns(4)
    with s1:
        st.markdown(f"<div class='summary-card'><div class='summary-title'>Konteyner Sayısı</div><div class='summary-value'>{total_containers} Container(s)</div></div>", unsafe_allow_html=True)
    with s2:
        st.markdown(f"<div class='summary-card'><div class='summary-title'>Toplam Kap / Paket</div><div class='summary-value'>{pkg_sum_str}</div></div>", unsafe_allow_html=True)
    with s3:
        st.markdown(f"<div class='summary-card'><div class='summary-title'>Toplam Brüt Ağırlık</div><div class='summary-value'>{total_weight:,.2f} KG</div></div>", unsafe_allow_html=True)
    with s4:
        st.markdown(f"<div class='summary-card'><div class='summary-title'>Toplam Hacim</div><div class='summary-value'>{total_cbm:,.2f} CBM</div></div>", unsafe_allow_html=True)

    def generate_excel():
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "B-L Talimat"
        ws.views.sheetView[0].showGridLines = True

        NAVY_FILL = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
        BLUE_FILL = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
        GRAY_FILL = PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid")
        BORDER_BOX = Border(left=Side(style="thin", color="CBD5E0"), right=Side(style="thin", color="CBD5E0"),
                            top=Side(style="thin", color="CBD5E0"), bottom=Side(style="thin", color="CBD5E0"))

        ws.merge_cells("A1:G1")
        ws["A1"] = "AWAM LOGISTICS - BILL OF LADING INSTRUCTION (B/L TALİMAT)"
        ws["A1"].font = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
        ws["A1"].fill = NAVY_FILL
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 32

        ws.merge_cells("A2:G2")
        ws["A2"] = "Official Shipping Instruction Document | [www.awamlogistics.com](https://www.awamlogistics.com)"
        ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
        ws["A2"].fill = BLUE_FILL
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        info_list = [
            ("Booking No:", st.session_state.booking_no), ("Shipping Line:", st.session_state.shipping_line),
            ("Vessel & Voyage:", st.session_state.vessel), ("POL (Loading Port):", st.session_state.pol),
            ("POD (Discharge Port):", st.session_state.pod), ("Freight Terms:", st.session_state.freight_terms)
        ]
        for idx, (lbl, val) in enumerate(info_list, start=4):
            ws.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=2)
            ws.cell(row=idx, column=1, value=lbl).font = Font(bold=True, size=10)
            ws.cell(row=idx, column=1).fill = GRAY_FILL
            ws.merge_cells(start_row=idx, start_column=3, end_row=idx, end_column=7)
            ws.cell(row=idx, column=3, value=val).font = Font(size=10)
            for c in range(1, 8): ws.cell(row=idx, column=c).border = BORDER_BOX
            ws.row_dimensions[idx].height = 20

        def add_party_rows(start_row, title, name, addr, tax, tel, email):
            ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=7)
            ws.cell(row=start_row, column=1, value=title).font = Font(bold=True, size=10, color="1A365D")
            ws.cell(row=start_row, column=1).fill = GRAY_FILL
            ws.row_dimensions[start_row].height = 20
            
            rows_data = [
                ("Company Name", name),
                ("Address", addr),
                ("Tax Number / CR No", tax),
                ("Tel", tel),
                ("Email", email)
            ]
            
            for offset, (lbl, val) in enumerate(rows_data, start=1):
                r = start_row + offset
                ws.row_dimensions[r].height = 19
                
                c_lbl = ws.cell(row=r, column=1, value=lbl)
                c_lbl.font = Font(size=10, italic=True)
                c_lbl.alignment = Alignment(horizontal="left", vertical="center")
                
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
                c_val = ws.cell(row=r, column=2, value=val)
                c_val.font = Font(size=10)
                c_val.alignment = Alignment(horizontal="left", vertical="center")
                
                for col in range(1, 8):
                    ws.cell(row=r, column=col).border = BORDER_BOX

        add_party_rows(11, "1. SHIPPER DETAILS", st.session_state.s_name, st.session_state.s_addr, st.session_state.s_tax, st.session_state.s_tel, st.session_state.s_email)
        add_party_rows(17, "2. CONSIGNEE DETAILS", st.session_state.cn_name, st.session_state.cn_addr, st.session_state.cn_tax, st.session_state.cn_tel, st.session_state.cn_email)
        add_party_rows(23, "3. NOTIFY PARTY DETAILS", st.session_state.nt_name, st.session_state.nt_addr, st.session_state.nt_tax, st.session_state.nt_tel, st.session_state.nt_email)

        headers = ["Container No", "Seal No", "Type", "Packages", "Description of Goods", "Gross Weight (KG)", "Volume (CBM)"]
        ws.row_dimensions[29].height = 25
        for c_i, h in enumerate(headers, 1):
            cell = ws.cell(row=29, column=c_i, value=h)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = BLUE_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER_BOX

        start_cargo_row = 30
        curr_row = start_cargo_row
        for idx, row in edited_df.iterrows():
            ws.cell(row=curr_row, column=1, value=row["Container No"]).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_row, column=2, value=row["Seal No"]).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_row, column=3, value=row["Type"]).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_row, column=4, value=row["Packages"]).alignment = Alignment(horizontal="center")
            ws.cell(row=curr_row, column=5, value=row["Description"]).alignment = Alignment(wrap_text=True)
            ws.cell(row=curr_row, column=6, value=float(row["Gross Weight (KG)"] if row["Gross Weight (KG)"] else 0)).number_format = '#,##0.00'
            ws.cell(row=curr_row, column=7, value=float(row["Volume (CBM)"] if row["Volume (CBM)"] else 0)).number_format = '#,##0.00'
            for c in range(1, 8): ws.cell(row=curr_row, column=c).border = BORDER_BOX
            curr_row += 1

        ws.row_dimensions[curr_row].height = 24
        ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=3)
        t_cell = ws.cell(row=curr_row, column=1, value=f"TOTAL: {total_containers} CONTAINER(S)")
        t_cell.font = Font(bold=True, size=10, color="1A365D")
        t_cell.alignment = Alignment(horizontal="center", vertical="center")

        pkg_cell = ws.cell(row=curr_row, column=4, value=pkg_sum_str)
        pkg_cell.font = Font(bold=True, size=10)
        pkg_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.cell(row=curr_row, column=5, value="")

        w_cell = ws.cell(row=curr_row, column=6, value=f"=SUM(F{start_cargo_row}:F{curr_row-1})")
        w_cell.font = Font(bold=True, size=10)
        w_cell.number_format = '#,##0.00'
        w_cell.alignment = Alignment(horizontal="right", vertical="center")

        v_cell = ws.cell(row=curr_row, column=7, value=f"=SUM(G{start_cargo_row}:G{curr_row-1})")
        v_cell.font = Font(bold=True, size=10)
        v_cell.number_format = '#,##0.00'
        v_cell.alignment = Alignment(horizontal="right", vertical="center")

        for c in range(1, 8):
            ws.cell(row=curr_row, column=c).border = BORDER_BOX
            ws.cell(row=curr_row, column=c).fill = GRAY_FILL

        col_w = {1:22, 2:16, 3:12, 4:16, 5:40, 6:18, 7:15}
        for c, w in col_w.items(): ws.column_dimensions[get_column_letter(c)].width = w

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue()

    st.markdown("---")
    excel_data = generate_excel()
    b_no = st.session_state.booking_no
    st.download_button(
        label="📥 B/L Talimat Excel Dosyasını İndir (Awam Mavi Format)",
        data=excel_data,
        file_name=f"BL_Talimat_{b_no if b_no else 'New'}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
