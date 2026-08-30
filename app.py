import io
import fitz  # PyMuPDF – reads text from PDFs
import streamlit as st
import anthropic
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ============================================================
# 1. CONFIG
# ============================================================
st.set_page_config(page_title="Morocco Conformity Skill", layout="wide")
st.title("🇲🇦 Morocco Product Conformity Skill")
st.caption("Type a product → get the applied norm(s), scope, tests and marking "
           "rules, pulled from your Google Drive standards folder.")

# Read-only access to Google Drive (this exact scope string matters)
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

# Model to use for the analysis
CLAUDE_MODEL = "claude-sonnet-5"

# Values come from Streamlit → Advanced settings → Secrets
API_KEY = st.secrets["AI_MODEL_API_KEY"]
GOOGLE_DRIVE_FOLDER_ID = st.secrets["DRIVE_FOLDER_ID"]

client = anthropic.Anthropic(api_key=API_KEY)

# ============================================================
# 2. GOOGLE DRIVE
# ============================================================
@st.cache_resource
def get_drive_service():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def list_all_folder_ids(service, root_id):
    """Walk the whole tree under the root folder and return every folder id."""
    all_ids = [root_id]
    to_visit = [root_id]
    while to_visit:
        parent = to_visit.pop()
        q = (f"'{parent}' in parents "
             f"and mimeType = 'application/vnd.google-apps.folder' "
             f"and trashed = false")
        resp = service.files().list(
            q=q,
            fields="files(id, name)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        ).execute()
        for f in resp.get("files", []):
            all_ids.append(f["id"])
            to_visit.append(f["id"])
    return all_ids


def search_standards(service, keyword):
    """Find PDFs anywhere under the root whose name OR content matches keyword."""
    folder_ids = list_all_folder_ids(service, GOOGLE_DRIVE_FOLDER_ID)
    matches = []
    # Query the folders in small chunks so the parent list never gets too long
    for i in range(0, len(folder_ids), 15):
        chunk = folder_ids[i:i + 15]
        parents = " or ".join([f"'{fid}' in parents" for fid in chunk])
        q = (f"({parents}) and mimeType = 'application/pdf' and trashed = false "
             f"and (name contains '{keyword}' or fullText contains '{keyword}')")
        resp = service.files().list(
            q=q,
            fields="files(id, name)",
            includeItemsFromAllDrives=True,
            supportsAllDrives=True,
        ).execute()
        matches.extend(resp.get("files", []))
    # de-duplicate by id
    seen, unique = set(), []
    for f in matches:
        if f["id"] not in seen:
            seen.add(f["id"])
            unique.append(f)
    return unique


def extract_pdf_text(service, file_id):
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    doc = fitz.open(stream=buf, filetype="pdf")
    return "".join(page.get_text() for page in doc)


# ============================================================
# 3. CLAUDE ANALYSIS
# ============================================================
def generate_report(product, tech_context, standard_text, standard_name):
    prompt = f"""You are an expert Moroccan market-control and conformity-verification
engineer (VOC / PortNet, Law 24-09). Analyse the official standard text below and
build a verification profile for the product.

TARGET PRODUCT: {product}
USER TECHNICAL DATA: {tech_context or "(none provided)"}
SOURCE DOCUMENT: {standard_name}

STANDARD TEXT:
{standard_text[:60000]}

Reply in English, in markdown, with EXACTLY these four sections:

### 1. Applied norm(s)
State the exact norm designation(s) and any referenced or equivalent standards
(EN, ISO, JIS, ...).

### 2. Simplified scope
2-3 plain-language sentences: what this norm covers, applies to, or excludes.

### 3. Mandatory tests
A markdown table with columns:
| Test / characteristic | Clause | Acceptance criteria / threshold |

### 4. Labelling & marking requirements
A markdown table with columns:
| Required element | Placement | Language / legibility |

If a value is missing or looks garbled in the source, say so rather than inventing it.
"""
    msg = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=4000,
        temperature=0,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text


# ============================================================
# 4. USER INTERFACE
# ============================================================
product = st.text_input(
    "Product name or keyword (e.g. 'embrayage', 'chauffe-eau', 'jouet', 'câble'):"
)
tech = st.text_area("Optional — paste technical-sheet text to narrow the match:")

# --- Step 1: search Drive ---
if st.button("🔍 Search standards"):
    if not product.strip():
        st.warning("Type a product first.")
        st.session_state.pop("files", None)
    else:
        try:
            service = get_drive_service()
            with st.spinner("Searching your Drive standards..."):
                st.session_state["files"] = search_standards(service, product.strip())
                st.session_state["product"] = product.strip()
                st.session_state["tech"] = tech
        except Exception as e:
            st.error(f"Google Drive error — check your [gcp_service_account] secret "
                     f"and that the folder is shared with the bot. Details: {e}")
            st.session_state.pop("files", None)

# --- Step 2: pick a document and analyse it ---
files = st.session_state.get("files")
if files is not None:
    if not files:
        st.error(
            "No PDF matched. Your files are named by NM code, so matching relies on "
            "Drive's full-text index of the PDF contents — try another keyword, or the NM code."
        )
    else:
        names = [f["name"] for f in files]
        choice = st.selectbox(f"{len(files)} document(s) found — pick one:", names)
        if st.button("✨ Analyse this document"):
            chosen = files[names.index(choice)]
            service = get_drive_service()
            with st.spinner(f"Reading {chosen['name']}..."):
                text = extract_pdf_text(service, chosen["id"])
            if not text.strip():
                st.warning("This PDF has no extractable text (likely a scan — OCR needed).")
            else:
                with st.spinner("Generating report..."):
                    report = generate_report(
                        st.session_state["product"],
                        st.session_state.get("tech", ""),
                        text,
                        chosen["name"],
                    )
                st.markdown("---")
                st.markdown(report)
