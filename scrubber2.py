import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import pikepdf
import io
import re


st.set_page_config(
    page_title="Part-L / Dwelling Report Extraction",
    layout="wide",
    initial_sidebar_state="expanded")


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



with st.sidebar:
    st.title('PDF Extraction')
    file_type = st.radio(
        "Please select the type of file you are uploading below:",
        ["Part-L report", "Draft Part-L report", "Dwelling Report"],
        index=0,
    )
    st.write("You selected: ", file_type)

col = st.columns((6, 2), gap='medium')

if file_type == "Part-L report":
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
elif file_type == "Draft Part-L report":
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
elif file_type == "Dwelling Report":
    field_patterns = {  
    "Dwelling Type": r"Dwelling Type\n(.+?)\s*\n",
    "Address Line 1": r"Address line 1\n(.+?)\s*\n",
    "Address Line 2": r"Address line 2\n(.+?)\s*\n",
    "Address Line 3": r"Address line 3\n(.+?)\s*\n",
    "BER Number": r"\n([\d.]+)",
    "Total Floor Area": r"Totals\n([\d.]+)",
    "Air permeabilitiy Adjusted result [ac/h]": r"Adjusted result of air permeability test\n\[ac/h\]\n([\d.]+)",
    "Thermal Bridging Factor": r"Thermal bridging factor \[W/m K\]\n2\n([\d.]+)",
    "Heat Use during Heating Season": r"Heat use during heating season \[kWh/y\]\n([\d.]+)",
    "Main space heating system (Delivered energy)": r"Main space heating system\n([\d.]+)",
    "Secondary space heating system (Delivered energy)": r"Secondary space heating system\n([\d.]+)",
    "Main water heating system (Delivered energy)": r"Main water heating system\n([\d.]+)",
    "Supplementary water heating system (Delivered energy)": r"Supplementary water heating system\n([\d.]+)",
    "Pumps and fans (Delivered energy)": r"Pumps and fans\n([\d.]+)",
    "Energy for lighting (Delivered energy)": r"Energy for lighting\n([\d.]+)",
    "Energy Rating": r"Energy Rating\n([A-G\d]+)",
    "Primary energy (Per m2 floor area)": r"Per m\s* floor area\n2\n[\d.]+\n([\d.]+)",
    "CO2 emissions (Per m2 floor area)": r"Per m\s* floor area\n2\n[\d.]+\n[\d.]+\n([\d.]+)"
    # Add additional patterns here as necessary
}


with col[0]:
    if file_type == "Part-L report":
        st.title('Extracting Information from Part L reports')
    elif file_type == "Draft Part-L report":
        st.title('Extracting Information from Draft Part L reports')
    elif file_type == "Dwelling Report":
        st.title('Extracting Information from Dwelling Reports')

    st.write("Please upload your files below (Max ~250 files)")

    # Upload multiple PDF files
    uploaded_files = st.file_uploader("", type="pdf", accept_multiple_files=True)

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


with col[1]:
    st.title('User Instructions')
    st.write("""
1. Please select the file type in the sidebar on the left
2. Please upload the files you want information extracted from on the left column
3. Once the progress bar is completed, click the download button to get the excel sheet
""")
    with st.expander('Troubleshooting', expanded=True):
        st.write('''
            If the excel file is not outputting the correct information, see below:
            - Is the correct file type selected in the sidebar?.
            - For Draft Part-L reports,have you tried the regular Part-L report button? 
            - If this still doesn't work, please just email/text - Samreet
            ''')
