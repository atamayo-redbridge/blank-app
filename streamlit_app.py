import streamlit as st
import pandas as pd
import pdfplumber
from io import BytesIO

# Function to extract text while preserving layout
def extract_text_from_pdf(pdf_file):
    text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            extracted_text = page.extract_text(x_tolerance=2, y_tolerance=2)
            if extracted_text:
                text += extracted_text + "\n\n"
    return text

# Function to extract tables from a PDF
def extract_tables_from_pdf(pdf_file):
    tables = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            extracted_tables = page.extract_tables()
            for table in extracted_tables:
                df = pd.DataFrame(table)
                tables.append(df)
    return tables

# Function to save extracted text and tables to an Excel file
def save_to_excel(text_data, tables_data):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        text_df = pd.DataFrame({"Extracted Text": text_data.split("\n")})
        text_df.to_excel(writer, sheet_name="Extracted Text", index=False)

        if tables_data:
            for i, table in enumerate(tables_data):
                table.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)

    output.seek(0)
    return output

# ------------------ Streamlit UI ------------------
st.set_page_config(page_title="📄 PDF to Excel Converter", layout="wide")

st.title("📄 PDF to Excel Converter with Preview")
st.write("Upload a **PDF file**, preview extracted content, and download it as **Excel**.")

uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

if uploaded_file:
    extracted_text = extract_text_from_pdf(uploaded_file)
    extracted_tables = extract_tables_from_pdf(uploaded_file)

    # Create a two-column layout for side-by-side preview
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📜 Extracted Text Preview")
        st.text_area("Text Output", extracted_text, height=400)

    with col2:
        st.subheader("📊 Extracted Tables Preview")
        if extracted_tables:
            for i, table in enumerate(extracted_tables):
                st.write(f"Table {i+1}")
                st.dataframe(table)
        else:
            st.write("No tables detected.")

    # Convert extracted content to Excel
    excel_file = save_to_excel(extracted_text, extracted_tables)

    st.subheader("📥 Download Excel File")
    st.download_button(
        label="Download Extracted Data as Excel",
        data=excel_file,
        file_name="extracted_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
