import streamlit as st
import pandas as pd
from io import BytesIO
from google import genai
from google.genai import types
from google.genai.errors import APIError
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# --- 1. CONFIGURATION ---

# Structured output schema for reliable data extraction
INVOICE_SCHEMA = types.Schema(
    type=types.Type.OBJECT,
    properties={
        "Invoice_Date": types.Schema(type=types.Type.STRING, description="The date the invoice was issued, in YYYY-MM-DD format (e.g., 2025-08-31)."),
        "Total_Amount_Excl_VAT": types.Schema(type=types.Type.STRING, description="The total amount of the invoice *excluding* VAT/Tax. Only the numerical value, no currency symbols or commas, using a DOT (.) for the decimal separator."),
        "Total_Amount_Incl_VAT": types.Schema(type=types.Type.STRING, description="The final total amount of the invoice *including* all VAT/Tax. Only the numerical value, no currency symbols or commas, using a DOT (.) for the decimal separator."),
        "VAT_Amount": types.Schema(type=types.Type.STRING, description="The total VAT/Tax amount, if explicitly listed. Only the numerical value, no currency symbols or commas, using a DOT (.) for the decimal separator."),
    },
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
    return 'application/octet-stream'

# --- 2. EXTRACTION FUNCTION ---

def extract_invoice_data(file_bytes, file_name):
    """Calls the Gemini API to extract structured data from a file."""
    try:
        # Initialize the client (API key is pulled from Streamlit secrets automatically)
        api_key = st.secrets.get("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("GEMINI_API_KEY is not configured in Streamlit secrets.")
        
        client = genai.Client(api_key=api_key)

        # Determine correct MIME type
        file_type = file_name.split('.')[-1].lower()
        mime_type = get_mime_type(file_name, file_type)
        
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

        data = pd.read_json(response.text, typ='series')
        data['File_Name'] = file_name
        return data

    except APIError as e:
        st.error(f"Gemini API Error for **{file_name}**: The API reported an issue. Details: `{e}`")
        return None
    except ValueError as e:
        st.error(f"Configuration Error: {e}")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred for **{file_name}**: {e}")
        return None

# --- 3. DATA CLEANUP & CONVERSION FUNCTION (NEW) ---

def format_amount_to_number(value):
    """Converts a string amount (potentially with commas as thousands separators) to a clean float."""
    if pd.isna(value) or str(value).upper() in ['NOT_FOUND', 'NAN', '']:
        return None  # Use None for missing data, which is better than text
    
    # 1. Clean the string: remove commas (thousands separator), and ensure dot is the decimal.
    # We are assuming the API returns standard US/International format (dot decimal, comma thousands).
    value_str = str(value).replace(',', '').strip() 
    
    try:
        return float(value_str)
    except ValueError:
        return None # Return None if conversion to a number fails

# --- 4. EXCEL GENERATION FUNCTION (MODIFIED) ---

@st.cache_data
def convert_df_to_excel(df):
    """Generates the Excel file, applying the comma decimal format."""
    
    # We need to use openpyxl directly to apply a specific number format
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice_Data"
    
    # 1. Write the DataFrame headers and rows to the worksheet
    for r_idx, row in enumerate(dataframe_to_rows(df, header=True, index=False)):
        ws.append(row)
        
    # 2. Define the columns to format (based on the reordered DF)
    # Total_Amount_Incl_VAT is in column C (index 3), Excl_VAT in D (index 4), VAT_Amount in E (index 5)
    formatted_cols = [3, 4, 5]
    
    # 3. Define the custom number format (tells Excel to use comma as decimal)
    # Example format: 1.234,56 (0 is for whole number, 00 is for decimal places)
    comma_decimal_format = '#,##0.00' 

    # 4. Apply the format to all data cells in the target columns (starting from row 2, skipping header)
    for col_idx in formatted_cols:
        # Iterate over rows starting from the first data row (row index 2)
        for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value is not None:
                    # Set the number format
                    cell.number_format = comma_decimal_format
                    
    # Save the workbook to the BytesIO object
    wb.save(output)
    return output.getvalue()


# --- 5. STREAMLIT APP LAYOUT & LOGIC ---

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

if 'results' not in st.session_state:
    st.session_state.results = []

if process_button and uploaded_files:
    st.session_state.results = []
    total_files = len(uploaded_files)

    with st.spinner(f"Processing {total_files} invoice(s)..."):
        progress_bar = st.progress(0)
        
        for i, uploaded_file in enumerate(uploaded_files):
            st.info(f"Processing file {i+1} of {total_files}: **{uploaded_file.name}**")
            
            file_bytes = uploaded_file.getvalue()
            extracted_series = extract_invoice_data(file_bytes, uploaded_file.name)
            
            if extracted_series is not None:
                st.session_state.results.append(extracted_series)
            
            progress_bar.progress((i + 1) / total_files)

    if st.session_state.results:
        # Combine all results into a single DataFrame
        df = pd.DataFrame(st.session_state.results)
        
        # --- NEW CONVERSION SECTION ---
        amount_cols = ['Total_Amount_Excl_VAT', 'Total_Amount_Incl_VAT', 'VAT_Amount']
        for col in amount_cols:
            # Convert string extraction to actual float numbers
            df[col] = df[col].apply(format_amount_to_number)
        
        # --- END NEW CONVERSION SECTION ---
        
        # Reorder columns for a better view
        df = df[['File_Name', 'Invoice_Date', 'Total_Amount_Incl_VAT', 'Total_Amount_Excl_VAT', 'VAT_Amount']]
        
        st.success("✅ Extraction Complete! See the results below.")
        
        # Display the DataFrame (Note: Streamlit's display will still use dot decimals)
        st.dataframe(df.fillna('NOT_FOUND'))

        # Generate the Excel file with the custom number formatting
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
