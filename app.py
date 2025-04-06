import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="PDF Text to Excel Converter", layout="centered")
st.title("📄 PDF Text to Excel Converter")

uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

# Optional: Page Range
start_page = st.number_input("Start Page (0-indexed)", min_value=0, value=0)
end_page = st.number_input("End Page (Leave 0 for all pages)", min_value=0, value=0)

def parse_text_to_dicts(text):
    # Split sections by blank lines
    entries = re.split(r'\n\s*\n', text.strip())
    data = []
    for entry in entries:
        record = {}
        lines = entry.strip().split('\n')
        for line in lines:
            if ':' in line:
                key, value = line.split(':', 1)
                record[key.strip()] = value.strip()
        if record:
            data.append(record)
    return data

if uploaded_file:
    try:
        with pdfplumber.open(uploaded_file) as pdf:
            total_pages = len(pdf.pages)
            pages_to_process = range(start_page, end_page) if end_page > 0 else range(total_pages)
            full_text = ""

            for i in pages_to_process:
                page = pdf.pages[i]
                full_text += page.extract_text() + "\n"

            parsed_data = parse_text_to_dicts(full_text)

            if parsed_data:
                df = pd.DataFrame(parsed_data)
                st.subheader("Extracted & Structured Data")
                st.dataframe(df)

                def to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name='ExtractedData')
                    return output.getvalue()

                excel_data = to_excel(df)

                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_data,
                    file_name="structured_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("No structured data found in the text.")

    except Exception as e:
        st.error(f"⚠️ Error: {str(e)}")



