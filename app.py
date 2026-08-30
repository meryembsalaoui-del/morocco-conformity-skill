import io
import os
import re
import fitz  # PyMuPDF – reads text from PDFs
import streamlit as st
import anthropic
from docx import Document
from docx.shared import RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ============================================================
# 1. CONFIG
# ============================================================
MAROON = "#78352A"
LOGO = "ttec_logo.png"
CLAUDE_MODEL = "claude-sonnet-5"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

st.set_page_config(
    page_title="TTEC - Morocco Conformity Skill",
    page_icon=LOGO if os.path.exists(LOGO) else "🇲🇦",
    layout="wide",
)

API_KEY = st.secrets["AI_MODEL_API_KEY"]
GOOGLE_DRIVE_FOLDER_ID = st.secrets["DRIVE_FOLDER_ID"]
client = anthropic.Anthropic(api_key=API_KEY)

# ============================================================
# 2. STYLING
# ============================================================
st.markdown(f"""
<style>
.stApp {{ background:#ffffff; }}
.ttec-title {{ color:{MAROON}; font-size:1.9rem; font-weight:700; margin:0; }}
.ttec-sub {{ color:#555; margin:.2rem 0 1rem 0; }}
.ttec-rule {{ border:0; border-top:3px solid {MAROON}; margin:.3rem 0 1rem 0; }}
.stButton>button {{ background:{MAROON}; color:#fff; border:none; border-radius:6px; font-weight:600; }}
.stButton>button:hover {{ background:#5c2820; color:#fff; }}
.stDownloadButton>button {{ background:{MAROON}; color:#fff; border:none; border-radius:6px; font-weight:600; }}
div[data-testid="stMarkdownContainer"] table {{ width:100%; border-collapse:collapse; }}
div[data-testid="stMarkdownContainer"] th {{ background:{MAROON}; color:#fff; padding:8px; text-align:left; }}
div[data-testid="stMarkdownContainer"] td {{ border:1px solid #e6dcd7; padding:8px; }}
</style>
""", unsafe_allow_html=True)

# ---------- header ----------
c1, c2 = st.columns([1, 6])
with c1:
    if os.path.exists(LOGO):
        st.image(LOGO, width=170)
with c2:
    st.markdown('<p class="ttec-title">Morocco Product Conformity Skill</p>', unsafe_allow_html=True)
    st.markdown('<p class="ttec-sub">Enter a product or norm codes - applied norm(s), scope, '
                'tests and marking rules, from the TTEC standards library.</p>', unsafe_allow_html=True)
st.markdown('<hr class="ttec-rule">', unsafe_allow_html=True)

# ---------- sidebar ----------
with st.sidebar:
    if os.path.exists(LOGO):
        st.image(LOGO, width=150)
    st.markdown("### How it works")
    st.markdown("1. Type a product **or several norm codes**, comma-separated.\n"
                "2. Click **Search standards**.\n"
                "3. **Tick every norm** that applies.\n"
                "4. Click **Analyse** - one consolidated report.")
    st.markdown("### Examples")
    st.code("embrayage")
    st.code("joint torique, ISO 3601, ISO 681")
    st.code("jouet, EN 71, EN 62115")
    st.caption("Many products need several norms - add each family as a keyword.")

# ============================================================
# 3. GOOGLE DRIVE
# ============================================================
@st.cache_resource
def get_drive_service():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def list_all_folder_ids(service, root_id):
    all_ids, to_visit = [root_id], [root_id]
    while to_visit:
        parent = to_visit.pop()
        q = (f"'{parent}' in parents and mimeType = 'application/vnd.google-apps.folder' "
             f"and trashed = false")
        resp = service.files().list(q=q, fields="files(id, name)",
                                    includeItemsFromAllDrives=True,
                                    supportsAllDrives=True).execute()
        for f in resp.get("files", []):
            all_ids.append(f["id"])
            to_visit.append(f["id"])
    return all_ids


def search_standards(service, keywords):
    """Search every folder under the root for PDFs matching ANY of the keywords."""
    folder_ids = list_all_folder_ids(service, GOOGLE_DRIVE_FOLDER_ID)
    matches = []
    for kw in keywords:
        kw = kw.strip()
        if not kw:
            continue
        for i in range(0, len(folder_ids), 15):
            chunk = folder_ids[i:i + 15]
            parents = " or ".join([f"'{fid}' in parents" for fid in chunk])
            q = (f"({parents}) and mimeType = 'application/pdf' and trashed = false "
                 f"and (name contains '{kw}' or fullText contains '{kw}')")
            resp = service.files().list(q=q, fields="files(id, name)",
                                        includeItemsFromAllDrives=True,
                                        supportsAllDrives=True).execute()
            matches.extend(resp.get("files", []))
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
# 4. CLAUDE ANALYSIS (consolidated, multi-norm)
# ============================================================
def generate_report(product, tech_context, docs):
    """docs = list of (filename, text). Produces one consolidated report."""
    per_doc = 40000
    blocks = [f"=== SOURCE DOCUMENT: {name} ===\n{text[:per_doc]}" for name, text in docs]
    combined = "\n\n".join(blocks)
    prompt = f"""You are an expert Moroccan market-control and conformity-verification engineer
(TTEC, VOC / PortNet, Law 24-09). You are given ONE OR MORE official standards that all apply to
the same product. Produce a single CONSOLIDATED verification profile combining them.

TARGET PRODUCT: {product}
USER TECHNICAL DATA: {tech_context or "(none provided)"}

{combined}

Reply in English, in markdown, with EXACTLY these sections:

### 1. Applied norm(s)
List every norm found. For each, state in one line whether it is:
- a GENERAL / BASE norm (whole product family),
- a PRODUCT-SPECIFIC norm, or
- a CONDITIONAL norm (applies only if the product has a feature - e.g. electrical parts,
  flammable materials); state the condition.
Also list referenced/equivalent standards (EN, ISO, JIS...).

### 2. Simplified scope
2-4 plain-language sentences on what these norms together apply to or exclude.

### 3. Mandatory tests
One combined markdown table:
| Norm & clause | Test / characteristic | Acceptance criteria / threshold |

### 4. Labelling & marking requirements
One combined markdown table:
| Norm | Required element | Placement | Language / legibility |

### 5. Notes & gaps
- Flag any value that is missing or looks garbled in the source (never invent one).
- If a product feature might trigger an ADDITIONAL norm that is NOT among the documents
  provided (e.g. an electrical toy needing EN 62115), name it as a reminder.
"""
    msg = client.messages.create(model=CLAUDE_MODEL, max_tokens=6000,
                                 messages=[{"role": "user", "content": prompt}])
    return msg.content[0].text


# ============================================================
# 5. WORD EXPORT
# ============================================================
def _add_runs(paragraph, text, force_bold=False, white=False):
    for part in re.split(r"(\*\*.+?\*\*)", text):
        if part.startswith("**") and part.endswith("**"):
            r = paragraph.add_run(part[2:-2]); r.bold = True
        else:
            r = paragraph.add_run(part)
        if force_bold:
            r.bold = True
        if white:
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)


def _shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def _add_table(doc, table_lines):
    rows = [[c.strip() for c in ln.strip().strip("|").split("|")] for ln in table_lines]
    rows = [r for r in rows if not all(set(c) <= set("-: ") and c != "" for c in r)]
    if not rows:
        return
    ncol = max(len(r) for r in rows)
    table = doc.add_table(rows=0, cols=ncol)
    table.style = "Table Grid"
    for ridx, r in enumerate(rows):
        cells = table.add_row().cells
        for cidx in range(ncol):
            txt = r[cidx] if cidx < len(r) else ""
            p = cells[cidx].paragraphs[0]
            if ridx == 0:
                _shade(cells[cidx], "78352A")
                _add_runs(p, txt, force_bold=True, white=True)
            else:
                _add_runs(p, txt)


def report_to_docx(markdown_text, title):
    doc = Document()
    h = doc.add_heading(level=0)
    run = h.add_run(f"TTEC - Conformity Report: {title}")
    run.font.color.rgb = RGBColor(0x78, 0x35, 0x2A)

    lines = markdown_text.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i].rstrip()
        if not line.strip():
            i += 1; continue
        if line.startswith("#"):
            level = min(len(line) - len(line.lstrip("#")), 3)
            hh = doc.add_heading(line.lstrip("#").strip(), level=max(1, level))
            for r in hh.runs:
                r.font.color.rgb = RGBColor(0x78, 0x35, 0x2A)
            i += 1; continue
        if line.lstrip().startswith("|"):
            tbl = []
            while i < len(lines) and lines[i].lstrip().startswith("|"):
                tbl.append(lines[i].strip()); i += 1
            _add_table(doc, tbl); continue
        if line.lstrip().startswith(("- ", "* ")):
            p = doc.add_paragraph(style="List Bullet")
            _add_runs(p, line.lstrip()[2:]); i += 1; continue
        _add_runs(doc.add_paragraph(), line); i += 1

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# 6. USER INTERFACE
# ============================================================
kw_input = st.text_input("Product or norm codes (comma-separated):",
                         placeholder="e.g. jouet, EN 71, EN 62115")
tech = st.text_area("Optional - paste technical-sheet text to narrow the match:")

# --- Step 1: search ---
if st.button("🔍 Search standards"):
    kws = [k for k in kw_input.split(",") if k.strip()]
    st.session_state.pop("report", None)
    if not kws:
        st.warning("Type at least one product or norm code.")
        st.session_state.pop("files", None)
    else:
        try:
            service = get_drive_service()
            with st.spinner("Searching the TTEC standards library..."):
                st.session_state["files"] = search_standards(service, kws)
                st.session_state["product"] = kw_input.strip()
                st.session_state["tech"] = tech
        except Exception as e:
            st.error(f"Google Drive error - check the [gcp_service_account] secret, that the "
                     f"folder is shared with the bot, and that the Drive API is enabled. Details: {e}")
            st.session_state.pop("files", None)

# --- Step 2: choose norms + analyse ---
files = st.session_state.get("files")
if files is not None:
    if not files:
        st.error("No PDF matched. Try the NM code or another keyword - matching relies on "
                 "Drive's full-text index of the PDF contents.")
    else:
        names = [f["name"] for f in files]
        st.success(f"{len(files)} document(s) found.")
        chosen = st.multiselect("Tick every norm that applies to this product:", names, default=names)
        if st.button("✨ Analyse selected norms") and chosen:
            service = get_drive_service()
            docs, empty = [], []
            with st.spinner("Reading documents..."):
                for f in files:
                    if f["name"] in chosen:
                        t = extract_pdf_text(service, f["id"])
                        (docs if t.strip() else empty).append((f["name"], t))
            if empty:
                st.warning("No extractable text (likely scans - OCR needed): "
                           + ", ".join(n for n, _ in empty))
            if docs:
                with st.spinner("Generating consolidated report..."):
                    st.session_state["report"] = generate_report(
                        st.session_state["product"], st.session_state.get("tech", ""), docs)
                    st.session_state["report_sources"] = [n for n, _ in docs]

# --- Step 3: show report + download ---
report = st.session_state.get("report")
if report:
    st.markdown('<hr class="ttec-rule">', unsafe_allow_html=True)
    with st.expander("📄 Documents used in this report"):
        for n in st.session_state.get("report_sources", []):
            st.write("•", n)
    st.markdown(report)
    st.download_button(
        "⬇️ Download as Word (.docx)",
        report_to_docx(report, st.session_state.get("product", "report")),
        file_name="TTEC_conformity_report.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
