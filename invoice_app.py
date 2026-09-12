import streamlit as st
import pandas as pd
import datetime
import os
from invoice_generator import generate_awam_invoice

st.set_page_config(page_title="Awam Logistics - Invoice Generator", page_icon="🧾", layout="wide")

st.title("🧾 وحدة إصدار الفواتير المالية - Awam Logistics")

col_meta1, col_meta2 = st.columns(2)

with col_meta1:
    invoice_num = st.text_input("رقم الفاتورة (Invoice Number)", value="INV-2401")
    customer_info = st.text_area("بيانات العملاء (Customer Details)", value="Awam Global Cannealan tinerey\nBURSA / TURKIYE\nVN: 1234567890")

with col_meta2:
    invoice_date = st.date_input("تاريخ الفاتورة (Date)", value=datetime.date.today()).strftime("%d.%m.%Y")

st.subheader("📦 بنود الفاتورة (Invoice Items)")

default_items = pd.DataFrame([
    {"shipper": "Awam Global Cannealan tinerey", "description": "Sample Shipping Packet, BURSA / Turkiye", "units": 100.0, "unit_price": 30.0},
    {"shipper": "Awam Btunuk Kamen", "description": "Sample Shipping Packet, BURSA / Turkiye", "units": 10.0, "unit_price": 25.0},
    {"shipper": "Awam Global Cannealan tinerey", "description": "Sample Shipping Packet, BURSA / Turkiye", "units": 5.0, "unit_price": 25.0},
    {"shipper": "Mvr", "description": "Sample Shipping Packet, BURSA / Turkiye", "units": 1.0, "unit_price": 30.0}
])

edited_df = st.data_editor(default_items, num_rows="dynamic", use_container_width=True)

if st.button("🚀 إنشاء الفاتورة PDF"):
    items_list = edited_df.to_dict(orient="records")
    output_pdf = f"Invoice_{invoice_num}.pdf"
    
    try:
        generate_awam_invoice(
            filename=output_pdf,
            invoice_num=invoice_num,
            invoice_date=invoice_date,
            customer_info=customer_info,
            items_data=items_list,
            logo_path="AG-LOGO.png"
        )
        
        with open(output_pdf, "rb") as file:
            st.download_button(
                label="📥 تحميل الفاتورة PDF",
                data=file,
                file_name=output_pdf,
                mime="application/pdf"
            )
        st.success("✅ تم إنشاء الفاتورة بنجاح مطابقة للنموذج المعتمد!")
    except Exception as e:
        st.error(f"حدث خطأ أثناء إنشاء الفاتورة: {str(e)}")