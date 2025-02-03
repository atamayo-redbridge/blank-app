import streamlit as st
import pandas as pd
import pdfplumber
import fitz  # PyMuPDF
import pytesseract
from pdf2image import convert_from_path
from io import BytesIO
from PIL import Image
import numpy as np

# Function to extract text from PDF (with OCR fallback)
def extract_text_from_pdf(pdf_file):
    """Extracts text while detecting scanned pages & applying OCR when needed."""
    text = ""

    # Try extracting text normally (PyMuPDF)
    doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    for page in doc:
        text += page.get_text("text") + "\n\n"

    # If no text is found, assume it's scanned and apply OCR
    if not text.strip():
        pdf_file.seek(0)  # Reset file pointer
        images = convert_from_path(pdf_file)
        text = ""
        for img in images:
            text += pytesseract.image_to_string(img, config="--psm 6") + "\n\n"  # Optimized OCR
    return text.strip()

# Function to extract structured tables
def extract_tables_from_pdf(pdf_file):
    """Extracts tables from PDF with complex table handling (merged cells, nested rows)."""
    tables = []
    pdf_file.seek(0)  # Reset file pointer for pdfplumber
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            extracted_tables = page.extract_tables()
            for table in extracted_tables:
                df = pd.DataFrame(table)

                # Remove empty rows/columns
                df = df.dropna(how="all")
                df = df.dropna(axis=1, how="all")

                # Handle merged header cells
                if not df.empty and df.shape[1] > 1:
                    df.columns = df.iloc[0]  # Set first row as header
                    df = df[1:].reset_index(drop=True)

                    # Clean multi-line cells
                    df = df.applymap(lambda x: " ".join(x.split()) if isinstance(x, str) else x)
                    
                    tables.append(df)
    return tables

# Function to save extracted text and tables to Excel
def save_to_excel(text_data, tables_data):
    """Saves extracted content to an Excel file."""
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
st.set_page_config(page_title="📄 Advanced PDF to Excel Converter", layout="wide")

st.title("📄 PDF to Excel Converter with OCR & Advanced Table Handling")
st.write("Upload a **PDF file**, preview structured text and tables, and download as **Excel**.")

uploaded_file = st.file_uploader("Upload a PDF file", type=["pdf"])

if uploaded_file:
    extracted_text = extract_text_from_pdf(uploaded_file)
    extracted_tables = extract_tables_from_pdf(uploaded_file)

    # Create a two-column layout for better preview
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📜 Extracted Text Preview")
        st.text_area("Structured Text Output", extracted_text, height=400)

    with col2:
        st.subheader("📊 Extracted Tables Preview")
        if extracted_tables:
            for i, table in enumerate(extracted_tables):
                try:
                    st.write(f"Table {i+1}")
                    st.dataframe(table)  # Display table safely
                except ValueError:
                    st.warning(f"⚠️ Could not display Table {i+1}. Invalid format.")
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
