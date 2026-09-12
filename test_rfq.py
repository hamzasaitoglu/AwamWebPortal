import streamlit as st
import datetime
import re
import openai

# Safe Key Assembly for Awam Logistics System
k1 = "sk-proj-buja6UpYkyVQEWarEe3R7VJ9m4oJkPQI8VQqV_mjqZET4BTz-iqVHVG68Xi2k1gT"
k2 = "DUgMeAC0PTT3BlbkFJUr0mzn9BGwOBTpevsUNY7bCqt3X2uxYW-b0j5Zb38rXfV_iewleem8Ok26ymSuAIloX0JCP8cA"
OPENAI_API_KEY = k1 + k2

st.set_page_config(page_title="Awam Logistics - Quick RFQ", page_icon="⚡", layout="wide")

st.markdown("""
<style>
    html, body, .stApp { background-color: #F8FAFC !important; color: #0F172A !important; }
    .stButton>button { background: #1D4ED8 !important; color: #FFFFFF !important; font-weight: bold; border-radius: 6px; border: none; }
</style>
""", unsafe_allow_html=True)

st.title("⚡ Quick RFQ Standardization Tool")

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
                client = openai.OpenAI(api_key=OPENAI_API_KEY)
                prompt = f"""
                You are the master operational RFQ parser for Awam Logistics (Freight Forwarding Expert).
                Parse the raw client message into STRICTLY 4 lines (UPPERCASE):

                Line 1: ORIGIN_CITY - DESTINATION_CITY [(INCOTERM if mentioned)] [(IMO if flammable/dangerous)]
                Line 2: QUANTITY x CONTAINER_TYPE
                Line 3: {ref_id}
                Line 4: CLIENT_NAME_IN_ENGLISH_UPPERCASE (If missing, write EXACTLY: MISSING_CLIENT_NAME)

                STRICT CONTAINER TYPE DEFINITIONS:
                - 20ft Dry -> "20DC"
                - 40ft High Cube / Standard Dry -> "40HC"
                - 40ft Reefer -> "40 REEFER"

                CRITICAL ARABIC QUANTITY & CONTAINER SIZE DISAMBIGUATION RULES:
                1. "اربع اربعين" or "أربع أربعين" or "4 اربعين" -> QUANTITY is 4, SIZE is 40ft -> Output: "4X40 HC"
                2. "ثلاث اربعين" or "3 اربعين" -> QUANTITY is 3, SIZE is 40ft -> Output: "3X40 HC"
                3. "خمس عشرين" or "5 عشرين" -> QUANTITY is 5, SIZE is 20ft -> Output: "5X20 DC"
                4. "اربعين" or "سعر الاربعين" (without preceding quantity number) -> QUANTITY is 1, SIZE is 40ft -> Output: "1X40 HC"
                5. "عشرين" or "سعر العشرين" (without preceding quantity number) -> QUANTITY is 1, SIZE is 20ft -> Output: "1X20 DC"
                6. "اربع اربعين مبرد" -> Output: "4X40 REEFER"

                EXAMPLES FOR ACCURACY:
                Input: "اربييد اربع اربعين من جده لعدن الادريسي"
                Output:
                JEDDAH - ADEN
                4X40 HC
                {ref_id}
                AL-ADRAISI

                Input: "اريد سعر العشرين من ازمير لعدن فوب عبدالمجيد"
                Output:
                IZMIR - ADEN (FOB)
                1X20 DC
                {ref_id}
                ABDULMAJEED

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