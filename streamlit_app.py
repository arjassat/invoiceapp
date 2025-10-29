import streamlit as st
import pandas as pd
from io import BytesIO
from google import genai
from google.genai import types
from google.genai.errors import APIError
import os

# --- 1. CONFIGURATION ---

# Structured output schema for reliable data extraction
INVOICE_SCHEMA = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "Invoice_Date": types.Schema(type=types.Type.STRING, description="The date the invoice was issued, in YYYY-MM-DD format (e.g., 2025-08-31)."),
        "Total_Amount_Excl_VAT": types.Schema(type=types.Type.STRING, description="The total amount of the invoice *excluding* VAT/Tax. Only the numerical value, no currency symbols or commas."),
        "Total_Amount_Incl_VAT": types.Schema(type=types.Type.STRING, description="The final total amount of the invoice *including* all VAT/Tax. Only the numerical value, no currency symbols or commas."),
        "VAT_Amount": types.Schema(type=types.Type.STRING, description="The total VAT/Tax amount, if explicitly listed. Only the numerical value, no currency symbols or commas."),
    },
    # Ensure these key fields are always extracted
    required=["Invoice_Date", "Total_Amount_Excl_VAT", "Total_Amount_Incl_VAT"]
)

PROMPT = (
    "You are an expert financial data extractor. Analyze the uploaded invoice document. "
    "Identify and extract the following information. Ensure all numerical amounts are clean, "
    "with no currency symbols, commas, or letters, and follow the YYYY-MM-DD date format. "
    "If a value cannot be found, use the text 'NOT_FOUND'. "
    "Return the results strictly in the provided JSON format."
)

def get_mime_type(file_name, file_type):
    """Determines the correct MIME type for the API."""
    if file_type == 'pdf':
        return 'application/pdf'
    elif file_type in ['jpg', 'jpeg']:
        return 'image/jpeg'
    elif file_type == 'png':
        return 'image/png'
    return 'application/octet-stream' # Default for safety


# --- 2. EXTRACTION FUNCTION ---

def extract_invoice_data(file_bytes, file_name):
    """Calls the Gemini API to extract structured data from a file."""
    try:
        # Initialize the client (API key is pulled from Streamlit secrets automatically)
        # We need to explicitly set the API key from secrets for the client
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("GEMINI_API_KEY is not configured in Streamlit secrets.")
        
        # Initialize client with the fetched key
        client = genai.Client(api_key=api_key)

        # Determine correct MIME type
        file_type = file_name.split('.')[-1].lower()
        mime_type = get_mime_type(file_name, file_type)
        
        # Create a content part from the uploaded file bytes
        invoice_part = types.Part.from_bytes(
            data=file_bytes,
            mime_type=mime_type
        )

        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[PROMPT, invoice_part],
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=INVOICE_SCHEMA,
            ),
        )

        # The response text will be a JSON string conforming to the schema
        data = pd.read_json(response.text, typ='series')
        data['File_Name'] = file_name
        return data

    except APIError as e:
        st.error(f"Gemini API Error for **{file_name}**: The API reported an issue. This can often be due to an unsupported file, a content issue, or an invalid API key. Details: `{e}`")
        return None
    except ValueError as e:
        st.error(f"Configuration Error: {e}")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred for **{file_name}**: {e}")
        return None

# --- 3. STREAMLIT APP LAYOUT ---

st.set_page_config(
    page_title="AI Invoice Extractor to Excel",
    layout="centered"
)

st.title("📄 AI Invoice Data Extractor")
st.markdown("Upload your PDF/Image invoices. The app will extract the date and total amounts and output a single Excel file.")

uploaded_files = st.file_uploader(
    "Upload PDF or image invoices (PDF, JPG, PNG)",
    type=['pdf', 'jpg', 'jpeg', 'png'],
    accept_multiple_files=True
)

process_button = st.button("Extract Data & Generate Excel")

# Initialize session state for results if not present
if 'results' not in st.session_state:
    st.session_state.results = []

# --- 4. MAIN LOGIC ---

if process_button and uploaded_files:
    # Clear previous results
    st.session_state.results = []
    total_files = len(uploaded_files)

    with st.spinner(f"Processing {total_files} invoice(s)..."):
        progress_bar = st.progress(0)
        
        for i, uploaded_file in enumerate(uploaded_files):
            st.info(f"Processing file {i+1} of {total_files}: **{uploaded_file.name}**")
            
            # The file is a Streamlit UploadedFile object. Read its bytes.
            file_bytes = uploaded_file.getvalue()
            
            # Perform the extraction
            extracted_series = extract_invoice_data(file_bytes, uploaded_file.name)
            
            if extracted_series is not None:
                st.session_state.results.append(extracted_series)
            
            # Update progress
            progress_bar.progress((i + 1) / total_files)

    if st.session_state.results:
        # Combine all results into a single DataFrame
        df = pd.DataFrame(st.session_state.results)
        
        # Reorder columns for a better view
        df = df[['File_Name', 'Invoice_Date', 'Total_Amount_Incl_VAT', 'Total_Amount_Excl_VAT', 'VAT_Amount']]
        
        st.success("✅ Extraction Complete! See the results below.")
        
        # Display the DataFrame
        st.dataframe(df)

        # --- Create Excel for Download ---
        @st.cache_data
        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Invoice_Data')
            return output.getvalue()
        
        excel_data = convert_df_to_excel(df)

        st.download_button(
            label="📥 Download Extracted Data as Excel",
            data=excel_data,
            file_name="invoice_extraction_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No data was successfully extracted. Check the error messages above for details.")
elif process_button:
    st.warning("Please upload at least one invoice file before clicking the button.")
