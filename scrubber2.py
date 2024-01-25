import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import pikepdf
import io
import re


st.title('Extracting Information from Part L Report')
st.write("""
Step 1: upload required PDFs\n
Step 2: download CSV after processing is completed""")

# Function to unlock a single PDF file using pikepdf
def unlock_pdf(uploaded_file):
    try:
        # Read the uploaded file into a BytesIO buffer
        input_pdf_stream = io.BytesIO(uploaded_file.getvalue())

        # Open the PDF with pikepdf and save the unlocked PDF to a BytesIO buffer
        with pikepdf.open(input_pdf_stream) as pdf:
            output_pdf_stream = io.BytesIO()
            pdf.save(output_pdf_stream)
            output_pdf_stream.seek(0)

        return output_pdf_stream
    except Exception as e:
        st.error(f'Error in unlocking PDF: {e}')
        return None

# Function to safely extract information using regex pattern
def safe_extraction(pattern, text):
    match = re.search(pattern, text)
    return match.group(1).strip() if match else None

# Function to extract dynamic information from PDF text
def extract_dynamic_info(text, field_patterns):
    info_dict = {field: None for field in field_patterns}
    for field, pattern in field_patterns.items():
        match = safe_extraction(pattern, text)
        if match:
            info_dict[field] = float(match) if match.replace('.', '', 1).isdigit() else match
    return info_dict

# Function to process and extract data from a single PDF file
def process_single_pdf(pdf_stream, field_patterns):
    full_text = ""
    with fitz.open(stream=pdf_stream) as doc:
        for page_num in range(doc.page_count):
            page = doc[page_num]
            full_text += page.get_text()
    return extract_dynamic_info(full_text, field_patterns)

# Upload multiple PDF files
uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)


# Define the regex patterns for fields to extract
field_patterns = {  
    "Address Line 1": r"Address line 1\n(.+?)\s*\n",
    "Address Line 2": r"Address line 2\n(.+?)\s*\n",
    "Address Line 3": r"Address line 3\n(.+?)\s*\n",
    "Dwelling Type": r"Dwelling Type\n(.+?)\s*\n",
    "Total Floor Area": r"Total Floor Area\n([\d.]+)",
    "BER Result": r"BER Result\n(\w+)\s*\n",
    "BER Number": r"BER Number\n(\d+)\s*\n",
    "EPC": r"EPC\n([\d.]+)",
    "CPC": r"CPC\n([\d.]+)",
    # Add additional patterns here as necessary
}

if uploaded_files:
    # Initialize progress bar
    progress_bar = st.progress(0)
    num_files = len(uploaded_files)
    data = pd.DataFrame(columns=field_patterns.keys())

    for i, uploaded_file in enumerate(uploaded_files):
        with st.spinner(f'Processing {uploaded_file.name}...'):
            unlocked_pdf_stream = unlock_pdf(uploaded_file)
            if unlocked_pdf_stream:
                extracted_info = process_single_pdf(unlocked_pdf_stream, field_patterns)
                data = pd.concat([data, pd.DataFrame([extracted_info])], ignore_index=True)
            else:
                st.error(f"Failed to process {uploaded_file.name}")

        # Update progress bar
        progress_bar.progress((i + 1) / num_files)

    # Hide the progress bar after processing
    progress_bar.empty()

    # Convert the dataframe to CSV
    csv_data = data.to_csv(index=False).encode('utf-8')

    # Download button for the CSV file
    st.download_button(
        label="Download Extracted Data as CSV",
        data=csv_data,
        file_name='extracted_data.csv',
        mime='text/csv',
    )
