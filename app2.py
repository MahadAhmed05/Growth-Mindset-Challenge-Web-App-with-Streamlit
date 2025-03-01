import streamlit as st
from pdf2docx import Converter
from docx import Document
from fpdf import FPDF
import tempfile
import os
import fitz  # PyMuPDF for text extraction
from datetime import datetime

# Custom CSS without dark mode
def load_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Poppins', sans-serif;
    }
    
    /* Main container styling */
    .main-container {
        background-color: rgba(255, 255, 255, 0.95);
        border-radius: 20px;
        padding: 30px;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
        margin: 20px 0;
        backdrop-filter: blur(10px);
        color: #333333;
    }
    
    /* Header styling */
    .header {
        text-align: center;
        margin-bottom: 30px;
    }
    
    .header h1 {
        font-size: 2.5rem;
        font-weight: 700;
        color: #1E88E5;
        margin-bottom: 10px;
    }
    
    .header p {
        font-size: 1.1rem;
        color: #546E7A;
        font-weight: 300;
    }
    
    /* Card styling */
    .card {
        background-color: white;
        border-radius: 15px;
        padding: 25px;
        margin-bottom: 20px;
        box-shadow: 0 5px 15px rgba(0, 0, 0, 0.05);
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        color: #333333;
    }
    
    .card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
    }
    
    /* Button styling */
    .stButton > button {
        border-radius: 30px !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        letter-spacing: 1px !important;
        padding: 10px 25px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 10px rgba(30, 136, 229, 0.3) !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 7px 15px rgba(30, 136, 229, 0.4) !important;
    }
    
    .stDownloadButton > button {
        background: linear-gradient(90deg, #1E88E5 0%, #42A5F5 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 30px !important;
        padding: 12px 30px !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        letter-spacing: 1px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 4px 10px rgba(30, 136, 229, 0.3) !important;
    }
    
    .stDownloadButton > button:hover {
        background: linear-gradient(90deg, #1976D2 0%, #1E88E5 100%) !important;
        box-shadow: 0 7px 15px rgba(30, 136, 229, 0.4) !important;
        transform: translateY(-2px) !important;
    }
    
    /* File uploader styling */
    .upload-area {
        border: 2px dashed #BBDEFB;
        border-radius: 15px;
        padding: 30px 20px;
        text-align: center;
        margin: 20px 0;
        cursor: pointer;
        transition: all 0.3s ease;
    }
    
    .upload-area:hover {
        border-color: #64B5F6;
        background-color: rgba(187, 222, 251, 0.1);
    }
    
    /* Success/info message styling */
    .success-msg {
        background-color: #E8F5E9;
        border-left: 5px solid #4CAF50;
        color: #2E7D32;
        padding: 15px;
        border-radius: 5px;
        margin: 20px 0;
        font-weight: 500;
    }
    
    .info-msg {
        background-color: #E3F2FD;
        border-left: 5px solid #2196F3;
        color: #0D47A1;
        padding: 15px;
        border-radius: 5px;
        margin: 20px 0;
        font-weight: 500;
    }
    
    .warning-msg {
        background-color: #FFF8E1;
        border-left: 5px solid #FFC107;
        color: #FF8F00;
        padding: 15px;
        border-radius: 5px;
        margin: 20px 0;
        font-weight: 500;
    }
    
    /* Progress bar styling */
    .stProgress > div > div {
        background-color: #1E88E5 !important;
    }
    
    /* Features section */
    .features {
        display: flex;
        flex-wrap: wrap;
        justify-content: space-between;
        margin: 30px 0;
    }
    
    .feature-item {
        flex-basis: 30%;
        text-align: center;
        padding: 15px;
        border-radius: 10px;
        background-color: #F5F5F5;
        margin-bottom: 20px;
        color: #333333;
    }
    
    /* Date display */
    .date-display {
        text-align: center;
        margin: 20px 0;
        font-size: 1.2rem;
        color: #1E88E5;
        font-weight: 500;
    }
    
    .date-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1976D2;
        margin-top: 5px;
    }
    
    /* Footer styling */
    .footer {
        text-align: center;
        margin-top: 50px;
        padding-top: 20px;
        border-top: 1px solid #EEEEEE;
        font-size: 0.9rem;
        color: #9E9E9E;
    }
    
    /* Text area styling */
    .styled-textarea textarea {
        border-radius: 10px !important;
        border: 1px solid #E0E0E0 !important;
        padding: 15px !important;
        font-family: 'Poppins', sans-serif !important;
        width: 100% !important;
        background-color: #F9F9F9 !important;
        color: #333333 !important;
    }
    
    /* Override Streamlit's default text input and textarea */
    .stTextInput > div > div > input {
        background-color: #F9F9F9;
        color: #333333;
        border-radius: 10px;
    }
    
    /* Override for file uploader */
    .stFileUploader > div {
        background-color: transparent !important;
    }
    
    /* Animations */
    @keyframes fadeIn {
        from { opacity: 0; }
        to { opacity: 1; }
    }
    
    .animate-fade {
        animation: fadeIn 0.5s ease-in-out;
    }
    
    /* Mobile responsive adjustments */
    @media (max-width: 768px) {
        .header h1 {
            font-size: 2rem;
        }
        
        .feature-item {
            flex-basis: 100%;
        }
    }
    </style>
    """, unsafe_allow_html=True)

# App title and configuration
st.set_page_config(
    page_title="DocumentFlow | PDF ↔ Word Converter",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Load custom CSS
load_css()

# Main Container
st.markdown('<div class="main-container">', unsafe_allow_html=True)

# Header
st.markdown("""
<div class="header">
    <h1>DocumentFlow <span style="font-weight:300;">| PDF ↔ Word Converter</span></h1>
    <p>Professional document conversion with seamless preservation of formatting</p>
</div>
""", unsafe_allow_html=True)

# Feature highlights
st.markdown("""
<div class="features animate-fade">
    <div class="feature-item">
        <h3>🔄 Bidirectional</h3>
        <p>Convert from PDF to Word and Word to PDF with equal precision</p>
    </div>
    <div class="feature-item">
        <h3>🎨 Format Preservation</h3>
        <p>Maintain your document's layout, fonts, and styling</p>
    </div>
    <div class="feature-item">
        <h3>🔍 Text Extraction</h3>
        <p>Extract and edit text content from your PDFs</p>
    </div>
</div>
""", unsafe_allow_html=True)

# Date display - simplified version with only the date
current_date = datetime.now().strftime("%B %d, %Y")
st.markdown(f"""
<div class="date-display">
    <p>Today's Date</p>
    <div class="date-value">{current_date}</div>
</div>
""", unsafe_allow_html=True)

# Card container for file upload
st.markdown('<div class="card">', unsafe_allow_html=True)

# File Upload with improved styling
st.markdown('<div class="upload-area">', unsafe_allow_html=True)
st.markdown('### 📁 Drag and drop your file here')
st.markdown('Supported formats: PDF & DOCX')
uploaded_file = st.file_uploader("Choose a file", type=["pdf", "docx"], label_visibility="collapsed")
st.markdown('</div>', unsafe_allow_html=True)

# Show file info if uploaded
if uploaded_file:
    file_ext = os.path.splitext(uploaded_file.name)[-1].lower()
    file_size = uploaded_file.size / 1024  # Size in KB
    
    # Show file details
    st.markdown(f"""
    <div class="info-msg">
        <strong>File Details:</strong><br>
        📄 Filename: {uploaded_file.name}<br>
        📏 Size: {file_size:.2f} KB<br>
        🔄 Type: {file_ext.upper().replace(".", "")}
    </div>
    """, unsafe_allow_html=True)
    
    # Create a progress bar for visual feedback
    progress_bar = st.progress(0)
    for i in range(101):
        # Update progress bar
        progress_bar.progress(i)
        if i == 100:
            break
    
    # Convert PDF to Word
    if file_ext == ".pdf":
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf.write(uploaded_file.read())
            temp_pdf_path = temp_pdf.name

        docx_path = temp_pdf_path.replace(".pdf", ".docx")

        st.markdown('<div class="info-msg">🔄 Converting PDF to Word document...</div>', unsafe_allow_html=True)
        
        # Conversion
        cv = Converter(temp_pdf_path)
        cv.convert(docx_path)
        cv.close()

        with open(docx_path, "rb") as docx_file:
            btn = st.download_button(
                label="⬇️ Download Word Document",
                data=docx_file,
                file_name="converted.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        st.markdown('<div class="success-msg">✅ PDF successfully converted to Word!</div>', unsafe_allow_html=True)

        # New card for text extraction
        st.markdown('</div><div class="card">', unsafe_allow_html=True)
        st.markdown('<h3>📝 Extract Text from PDF</h3>', unsafe_allow_html=True)
        
        with open(temp_pdf_path, "rb") as pdf_file:
            doc = fitz.open(pdf_file)
            extracted_text = "\n".join([page.get_text() for page in doc])

        # Custom styled text area
        st.markdown('<div class="styled-textarea">', unsafe_allow_html=True)
        st.text_area("Extracted Text:", extracted_text, height=300, label_visibility="collapsed")
        st.markdown('</div>', unsafe_allow_html=True)

        st.download_button(
            label="⬇️ Download Text Content",
            data=extracted_text,
            file_name="extracted_text.txt",
            mime="text/plain"
        )

    # Convert Word to PDF
    elif file_ext == ".docx":
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
            temp_docx.write(uploaded_file.read())
            temp_docx_path = temp_docx.name

        pdf_path = temp_docx_path.replace(".docx", ".pdf")

        st.markdown('<div class="info-msg">🔄 Converting Word document to PDF...</div>', unsafe_allow_html=True)
        
        # Conversion
        doc = Document(temp_docx_path)
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=10)
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        for para in doc.paragraphs:
            pdf.cell(200, 10, txt=para.text, ln=True)

        pdf.output(pdf_path)

        with open(pdf_path, "rb") as pdf_file:
            st.download_button(
                label="⬇️ Download PDF Document",
                data=pdf_file,
                file_name="converted.pdf",
                mime="application/pdf"
            )

        st.markdown('<div class="success-msg">✅ Word document successfully converted to PDF!</div>', unsafe_allow_html=True)

else:
    st.markdown("""
    <div class="warning-msg">
        📂 Please upload a PDF or Word document to begin conversion.
    </div>
    """, unsafe_allow_html=True)

# Close the card div
st.markdown('</div>', unsafe_allow_html=True)

# How-to-use Guide
st.markdown('<div class="card" style="margin-top: 30px;">', unsafe_allow_html=True)
st.markdown("""
<h3>🔍 How to Use DocumentFlow</h3>
<ol>
    <li><strong>Upload your document</strong> - Drag and drop your PDF or Word file into the upload area.</li>
    <li><strong>Wait for conversion</strong> - Our intelligent converter will process your document while preserving formatting.</li>
    <li><strong>Download result</strong> - Click the download button to save your converted document.</li>
    <li><strong>For PDFs</strong> - You can also extract and download the text content separately.</li>
</ol>
""", unsafe_allow_html=True)
st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown(f"""
<div class="footer">
    <p>DocumentFlow Converter • Streamlit Edition • {datetime.now().year}</p>
</div>
""", unsafe_allow_html=True)

# Close the main container div
st.markdown('</div>', unsafe_allow_html=True)

# Try to clean up temp files (may not work in all environments)
try:
    if 'temp_pdf_path' in locals():
        os.unlink(temp_pdf_path)
        os.unlink(docx_path)
    if 'temp_docx_path' in locals():
        os.unlink(temp_docx_path)
        os.unlink(pdf_path)
except:
    pass