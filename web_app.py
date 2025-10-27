import streamlit as st
import pandas as pd
import matplotlib
matplotlib.use('Agg') # Use Agg backend for web servers
import matplotlib.pyplot as plt
import fitz # PyMuPDF
from openai import OpenAI
from fpdf import FPDF
import io

# --- Page Configuration ---
st.set_page_config(layout="wide")
st.title("SICOM AI Report Generator")
st.write("Please upload your financial data files. The report will be generated based on these documents.")

# --- API Key ---
# Load the API key from Streamlit's secure secrets manager
try:
    api_key = st.secrets["OPENAI_API_KEY"]
except KeyError:
    st.error("ERROR: OPENAI_API_KEY not found in Streamlit Secrets. Please add it.")
    st.stop()
# --- END API KEY ---


# --- Helper Functions ---
def generate_summary(text_from_pdf, latest_revenue, previous_revenue, client):
    st.write("Checkpoint A: Entered generate_summary function.")
    revenue_growth = ((latest_revenue - previous_revenue) / previous_revenue) * 100
    prompt = f"""
You are a professional financial analyst for SICOM.
The latest quarterly revenue is {latest_revenue:,.0f}.
The revenue for the same quarter last year was {previous_revenue:,.0f}.
This represents a {revenue_growth:.2f}% change.

The quarterly data also shows a significant negative revenue of -1,388.1M in Q3 2016.
Flag this in your analysis as a "Data Integrity Alert" and advise that it appears to be a data-entry error in the source file.

Analyze the following text from the financial report and provide a brief summary
of key performance indicators, risks, and outlook.

Text:
{text_from_pdf[:4000]}
"""
    st.write("Checkpoint B: Connecting to OpenAI API...")
    response = client.chat.completions.create(model="gpt-4o", messages=[{"role": "user", "content": prompt}], temperature=0.5)
    st.write("Checkpoint C: Received response from OpenAI API.")
    return response.choices[0].message.content

# --- File Uploaders ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Financial Data File")
    excel_file = st.file_uploader("Select the .xlsx file", type=["xlsx"])

with col2:
    st.subheader("2. PDF Report File")
    pdf_file = st.file_uploader("Select the .pdf file", type=["pdf"])

st.divider()

# --- Generate Report Button ---
if st.button("Generate Full Report"):

    # 1. Check if both files are uploaded
    if excel_file is not None and pdf_file is not None:

        with st.spinner("Generating report... This may take a moment."):

            try:
                # 2. Initialize OpenAI Client
                client = OpenAI(api_key=api_key)
                st.write("Checkpoint 4: OpenAI client initialized successfully.")

                # --- 3. Process the Excel File (from memory) ---
                st.write("Checkpoint 5: Loading and cleaning Excel file.")

                df = pd.read_excel(excel_file,
                                   sheet_name='Income Statement',
                                   header=3,
                                   index_col=0)

                df = df.drop(df.columns[0], axis=1)
                df = df.drop('3 Months Ending')
                st.write("Checkpoint 5.2: Excel sheet cleaned successfully.")

                revenue_data = pd.to_numeric(df.loc['Net Revenue'], errors='coerce')
                st.write("Checkpoint 5.5: Revenue data prepared.")

                # --- 4. Create the Chart (in memory) ---
                st.write("Checkpoint 6: Generating chart.")
                plt.figure(figsize=(10, 6))
                revenue_data.plot(kind='bar')
                plt.title('CIM Financials Quarterly Revenue')
                plt.xlabel('Quarter')
                plt.ylabel('Revenue (in thousands)')
                plt.xticks(rotation=45)
                plt.tight_layout()

                img_buffer = io.BytesIO()
                plt.savefig(img_buffer, format='png')
                img_buffer.seek(0)
                st.write("Checkpoint 6.1: Chart saved to memory.")

                # --- 5. Process the PDF File (from memory) ---
                st.write("Checkpoint 7: Extracting text from PDF.")
                text_content = ""

                with fitz.open(stream=pdf_file.read(), filetype="pdf") as doc:
                    for page in doc:
                        text_content += page.get_text()
                st.write("Checkpoint 7.1: PDF text extracted.")

                # --- 6. Run AI Analysis ---
                latest_rev = revenue_data.iloc[-1]
                previous_rev = revenue_data.iloc[-5]
                summary = generate_summary(text_content, latest_rev, previous_rev, client)

                # --- 7. Generate the PDF Report (in memory) ---
                st.write("Checkpoint 8: Generating PDF report...")
                pdf = FPDF()
                
                # --- START FONT FIX ---
                # Add Unicode fonts (must be in the same directory)
                try:
                    pdf.add_font('DejaVu', '', 'DejaVuSans.ttf', uni=True)
                    pdf.add_font('DejaVu', 'B', 'DejaVuSans-Bold.ttf', uni=True)
                    font_family = 'DejaVu'
                except FileNotFoundError:
                    st.error("Font files (DejaVuSans.ttf, DejaVuSans-Bold.ttf) not found. "
                             "Falling back to Arial, which may cause errors with special characters.")
                    font_family = 'Arial'
                # --- END FONT FIX ---
                
                pdf.add_page()
                
                # Use the new Unicode font
                pdf.set_font(font_family, 'B', 16)
                pdf.cell(0, 10, "SICOM Financial Analysis", ln=True, align='C')

                # Use the new Unicode font
                pdf.set_font(font_family, '', 11)
                pdf.multi_cell(0, 5, summary) # This will now work
                pdf.ln(10)

                # Pass only the buffer and dimensions
                pdf.image(img_buffer, x=10, y=None, w=190)

                st.write("Checkpoint 9: PDF report finished.")

                # --- 8. Show Success and Download Button ---
                st.success("--- Report finished successfully! ---")
                st.balloons()

                st.download_button(
                    label="Download SICOM_Report.pdf",
                    data=bytes(pdf.output()),
                    file_name="SICOM_Report.pdf",
                    mime="application/pdf"
                )

            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.exception(e)

    else:
        st.error("Please make sure both files are uploaded before generating the report.")
