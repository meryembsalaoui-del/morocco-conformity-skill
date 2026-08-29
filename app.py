import os
import io
import fitz  # PyMuPDF for high-speed PDF text parsing
import streamlit as st
import google.generativeai as genai
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# =====================================================================
# 1. APPLICATION & CONFIGURATION SETUP
# =====================================================================
st.set_page_config(page_title="Morocco Conformity AI Skill", layout="wide")
st.title("🇲🇦 Morocco Product Conformity Verification Skill")
st.write("Analyze technical sheets and cross-reference your custom Google Drive norms repository.")

# ⚠️ PLACE YOUR OPENAI OR GEMINI API KEY HERE TO POWER THE REASONING
# You can get a free Gemini API key from Google AI Studio
# Pull the Anthropic Claude API Key securely from Cloud Secrets
API_KEY = st.secrets["AI_MODEL_API_KEY"]

# Pull your unique Moroccan Standards root Google Drive Folder ID
GOOGLE_DRIVE_FOLDER_ID = st.secrets["DRIVE_FOLDER_ID"]

# This configuration variable is no longer needed since we authenticate from memory 
SERVICE_ACCOUNT_FILE = None 


# =====================================================================
# 2. GOOGLE DRIVE BACKEND FUNCTIONS
# =====================================================================
def get_drive_service():
    """Initializes secure service account connection to Google Drive."""
        SCOPES = ['https://googleapis.com']
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        st.error(f"Missing {SERVICE_ACCOUNT_FILE} in your project directory!")
        return None
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

def search_standards(service, keyword):
    """
    Scans the parent directory and recursively looks inside 
    all nested sub-folders for files matching the product name.
    """
    try:
        # Step A: Find all child sub-folder IDs within the main database
        folder_query = f"'{GOOGLE_DRIVE_FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder'"
        folder_results = service.files().list(q=folder_query, fields="files(id, name)").execute()
        all_folder_ids = [f['id'] for f in folder_results.get('files', [])]
        
        # Always include the root directory ID in the search sweep list
        all_folder_ids.append(GOOGLE_DRIVE_FOLDER_ID)
        
        # Step B: Build a combined query covering all discovered directories
        parent_constraints = " or ".join([f"'{fid}' in parents" for fid in all_folder_ids])
        file_query = f"({parent_constraints}) and (name contains '{keyword}' or fullText contains '{keyword}') and mimeType = 'application/pdf'"
        
        # Step C: Execute full file discovery sweep
        file_results = service.files().list(
            q=file_query, 
            fields="files(id, name)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True
        ).execute()
        
        files = file_results.get('files', [])
        return files[0] if files else None  # Return top relevant document match
        
    except Exception as e:
        st.error(f"Error traversing Google Drive sub-directories: {e}")
        return None


def download_and_extract_pdf_text(service, file_id):
    """Downloads matching document directly into RAM and extracts text strings."""
    request = service.files().get_media(fileId=file_id)
    file_stream = io.BytesIO()
    downloader = MediaIoBaseDownload(file_stream, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    
    file_stream.seek(0)
    pdf_document = fitz.open(stream=file_stream, filetype="pdf")
    text_content = ""
    for page in pdf_document:
        text_content += page.get_text()
    return text_content

# =====================================================================
# 3. AI EXTRACTION ENGINE (THE SKILL LOGIC)
# =====================================================================
def generate_conformity_report(product, technical_context, standard_text, standard_name):
    """Feeds norm text to AI engine to synthesize structured verification outputs."""
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    prompt = f"""
    You are an expert Moroccan market control and conformity verification engineer working under Law 24-09.
    Analyze the following official regulatory standard text to create a strict verification profile.
    
    TARGET PRODUCT: {product}
    USER TECHNICAL DATA: {technical_context}
    REFERENCED DOCUMENT: {standard_name}
    
    STANDARD REGULATORY TEXT DATA:
    {standard_text[:15000]}  # Safe token optimization truncation
    
    Format your response cleanly using markdown with the following structure:
    
    ### 1. Applied Norm(s)
    * State the exact norm designation found in the text (e.g., NM EN 60335-1).
    
    ### 2. Simplified Scope Explanation
    * Provide a 2-3 sentence non-technical summary explaining exactly what products this norm protects, applies to, or excludes.
    
    ### 3. Mandatory Laboratory Tests Required
    * Create a markdown table with columns: [Test Name / Characteristic] | [Reference Clause] | [Success Criteria / Threshold]
    * Populate with explicit tests found in the text (e.g., Dielectric strength, mechanical impact, flame resistance).
    
    ### 4. Labeling & Marking Requirements
    * Create a markdown table with columns: [Required Element] | [Placement Location] | [Language / Legibility Rules]
    * Detail things like CMIM mark placement, rated voltage markings, manufacturer name visibility, or required languages (e.g., Arabic).
    """
    
    response = model.generate_content(prompt)
    return response.text

# =====================================================================
# 4. USER INTERFACE (FRONTEND)
# =====================================================================
product_input = st.text_input("Enter product name or keyword (e.g., 'Chauffe-eau', 'Jouet', 'Câble'):")
tech_sheet_input = st.text_area("Optional: Paste specifications or technical sheet text here to narrow the match:")

if st.button("Run Conformity Analysis ✨"):
    if not product_input:
        st.warning("Please specify a target product first.")
    else:
        with st.spinner("Connecting to Google Drive Repository..."):
            drive_api = get_drive_service()
            
            if drive_api:
                # Find matching documents in your standards repository
                matching_files = search_standards(drive_api, product_input)
                
                if not matching_files:
                    st.error(f"No corresponding standard document found in your Drive folder matching: '{product_input}'")
                else:
                    st.success(f"Found standard: **{matching_files[0]['name']}**")
                    
                    # Extract document layers
                    raw_text = download_and_extract_pdf_text(drive_api, matching_files[0]['id'])
                    
                    st.info("Analyzing standard text and generating compliance profiles...")
                    # Run AI synthesis
                    compliance_report = generate_conformity_report(
                        product=product_input,
                        technical_context=tech_sheet_input,
                        standard_text=raw_text,
                        standard_name=matching_files[0]['name']
                    )
                    
                    # Output final formatted tables to user dashboard screen
                    st.markdown("---")
                    st.markdown(compliance_report)
