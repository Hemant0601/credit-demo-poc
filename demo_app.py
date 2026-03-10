"""
Credit Intelligence AI Copilot — Conversational Chat Demo
==========================================================
ChatGPT-style interface for credit officers.
AI has full access to uploaded loan documents and answers with citations.

Requirements:
    pip install streamlit pymupdf openai python-docx openpyxl pandas python-dotenv

Run:
    streamlit run demo_app.py
"""

import streamlit as st
import fitz                  # PyMuPDF
import openai
import hashlib
import os
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path

from dotenv import load_dotenv

load_dotenv()

try:
    import docx as python_docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
# OPENAI MODELS
# ─────────────────────────────────────────────────────────────────────────────

OPENAI_MODELS = [
    "gpt-5.4",
    "gpt-5-mini",
    "o3",
    "o3-pro",
    "o4-mini",
    "gpt-4o",
    "gpt-4o-mini",
]

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Credit Intelligence AI Copilot",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CSS — Modern dark-navy theme
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* ── Global ────────────────────────────────────────────────────────── */
    .stApp { background: #f8f9fb; }

    /* ── Top banner ────────────────────────────────────────────────────── */
    .top-banner {
        background: linear-gradient(135deg, #0f1b33 0%, #1a3a6b 50%, #2563eb 100%);
        padding: 20px 28px; border-radius: 14px; margin-bottom: 20px;
        box-shadow: 0 4px 20px rgba(15,27,51,0.25);
        position: relative; overflow: hidden;
    }
    .top-banner::before {
        content: ''; position: absolute; top: -40%; right: -10%;
        width: 300px; height: 300px; border-radius: 50%;
        background: rgba(255,255,255,0.04);
    }
    .top-banner h2 {
        color: white; margin: 0; font-size: 22px; font-weight: 700;
        letter-spacing: -0.3px;
    }
    .top-banner p {
        color: rgba(200,220,255,0.85); margin: 6px 0 0 0;
        font-size: 13px; line-height: 1.5;
    }

    /* ── Stat pills row ────────────────────────────────────────────────── */
    .stat-row {
        display: flex; gap: 10px; margin-top: 12px; flex-wrap: wrap;
    }
    .stat-pill {
        background: rgba(255,255,255,0.12); backdrop-filter: blur(6px);
        border: 1px solid rgba(255,255,255,0.15);
        border-radius: 8px; padding: 6px 14px;
        color: white; font-size: 12px; font-weight: 500;
        display: flex; align-items: center; gap: 6px;
    }
    .stat-pill .val { font-weight: 700; font-size: 14px; }

    /* ── Sidebar ───────────────────────────────────────────────────────── */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f1b33 0%, #162d50 100%);
    }
    section[data-testid="stSidebar"] * {
        color: #c8d8f0 !important;
    }
    section[data-testid="stSidebar"] .stMarkdown h4 {
        color: #e2ecff !important;
        font-size: 13px; text-transform: uppercase;
        letter-spacing: 1px; margin-top: 8px;
    }
    section[data-testid="stSidebar"] .stTextInput input,
    section[data-testid="stSidebar"] .stSelectbox > div > div {
        background: rgba(255,255,255,0.08) !important;
        border: 1px solid rgba(255,255,255,0.15) !important;
        color: white !important; border-radius: 8px !important;
    }
    section[data-testid="stSidebar"] .stDivider {
        border-color: rgba(255,255,255,0.1) !important;
    }
    section[data-testid="stSidebar"] .stAlert {
        background: rgba(255,255,255,0.06) !important;
        border: 1px solid rgba(255,255,255,0.1) !important;
        color: #a0b8d8 !important;
    }

    /* Sidebar brand */
    .sidebar-brand {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        border-radius: 12px; padding: 16px 18px; margin-bottom: 18px;
        border: 1px solid rgba(255,255,255,0.1);
    }
    .sidebar-brand .title {
        color: white !important; font-size: 18px; font-weight: 700;
        letter-spacing: -0.3px;
    }
    .sidebar-brand .subtitle {
        color: rgba(200,220,255,0.7) !important; font-size: 11px;
        margin-top: 2px; letter-spacing: 0.5px;
    }

    /* Document cards in sidebar */
    .doc-card {
        background: rgba(255,255,255,0.06); border-radius: 10px;
        padding: 10px 14px; margin: 6px 0;
        border: 1px solid rgba(255,255,255,0.08);
        transition: background 0.2s;
    }
    .doc-card:hover { background: rgba(255,255,255,0.1); }
    .doc-card strong { font-size: 13px; color: #93b5ff !important; }
    .doc-card .meta { color: #7a94b8 !important; font-size: 11px; margin-top: 3px; }

    /* ── Welcome screen ────────────────────────────────────────────────── */
    .welcome-box {
        text-align: center; padding: 60px 20px 40px; color: #888;
    }
    .welcome-icon {
        width: 80px; height: 80px; margin: 0 auto 20px;
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
        border-radius: 20px; display: flex; align-items: center;
        justify-content: center; font-size: 36px;
        box-shadow: 0 8px 30px rgba(37,99,235,0.25);
    }
    .welcome-title {
        font-size: 24px; font-weight: 700; color: #0f1b33;
        margin: 0 0 10px;
    }
    .welcome-desc {
        font-size: 15px; max-width: 560px; margin: 0 auto;
        line-height: 1.8; color: #64748b;
    }

    /* ── Suggestion chips ──────────────────────────────────────────────── */
    .stButton > button {
        border: 1px solid #e2e8f0 !important;
        background: white !important; color: #334155 !important;
        border-radius: 10px !important; font-size: 13px !important;
        font-weight: 500 !important; padding: 8px 16px !important;
        transition: all 0.2s !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04) !important;
    }
    .stButton > button:hover {
        background: #f0f5ff !important;
        border-color: #2563eb !important;
        color: #1d4ed8 !important;
        box-shadow: 0 2px 8px rgba(37,99,235,0.12) !important;
        transform: translateY(-1px);
    }

    /* ── Chat messages ─────────────────────────────────────────────────── */
    .stChatMessage {
        background: white !important;
        border: 1px solid #e8ecf2 !important;
        border-radius: 14px !important;
        padding: 16px 20px !important;
        margin-bottom: 12px !important;
        box-shadow: 0 1px 4px rgba(0,0,0,0.03) !important;
    }
    /* User messages — subtle blue tint */
    div[data-testid="stChatMessage"]:has(.stMarkdown) {
        border-left: 3px solid #2563eb !important;
    }

    /* ── Chat input ────────────────────────────────────────────────────── */
    .stChatInput {
        border-radius: 14px !important;
    }
    .stChatInput > div {
        border: 2px solid #e2e8f0 !important;
        border-radius: 14px !important;
        background: white !important;
        box-shadow: 0 2px 12px rgba(0,0,0,0.04) !important;
        transition: border-color 0.2s !important;
    }
    .stChatInput > div:focus-within {
        border-color: #2563eb !important;
        box-shadow: 0 2px 16px rgba(37,99,235,0.1) !important;
    }

    /* ── Status widget ─────────────────────────────────────────────────── */
    .stStatus {
        background: #f0f5ff !important;
        border: 1px solid #bfdbfe !important;
        border-radius: 10px !important;
    }

    /* ── Metric widget (queries) ───────────────────────────────────────── */
    div[data-testid="stMetricValue"] {
        color: #93b5ff !important; font-size: 28px !important;
    }

    /* ── Hide Streamlit chrome ─────────────────────────────────────────── */
    footer { visibility: hidden; }
    #MainMenu { visibility: hidden; }
    header[data-testid="stHeader"] { background: transparent; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
for k, v in {
    "messages":       [],
    "documents":      {},
    "loan_id":        "LOAN-2024-CORP-001",
    "officer":        "Credit Officer",
    "prefill":        "",
    "doc_context":    "",
    "model":          "gpt-5.4",
    "data_dir_loaded": False,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─────────────────────────────────────────────────────────────────────────────
# PARSERS
# ─────────────────────────────────────────────────────────────────────────────

def extract_pdf(raw: bytes, name: str) -> dict:
    doc  = fitz.open(stream=raw, filetype="pdf")
    full = ""
    for p in doc:
        full += f"\n--- Page {p.number+1} ---\n{p.get_text()}"
    return {"filename":name,"type":"PDF","page_count":len(doc),
            "full_text":full,"word_count":len(full.split()),
            "hash":hashlib.sha256(raw).hexdigest()[:10]}


def extract_docx(raw: bytes, name: str) -> dict:
    if not DOCX_AVAILABLE:
        return {"error":"python-docx not installed"}
    doc  = python_docx.Document(BytesIO(raw))
    full = ""
    for p in doc.paragraphs:
        if p.text.strip():
            full += f"\n{p.text}"
    for i, t in enumerate(doc.tables):
        full += f"\n[Table {i+1}]\n"
        for row in t.rows:
            full += " | ".join(c.text.strip() for c in row.cells) + "\n"
    return {"filename":name,"type":"WORD","page_count":"—",
            "full_text":full,"word_count":len(full.split()),
            "hash":hashlib.sha256(raw).hexdigest()[:10]}


def extract_excel(raw: bytes, name: str) -> dict:
    if not PANDAS_AVAILABLE:
        return {"error":"pandas not installed"}
    xl   = pd.ExcelFile(BytesIO(raw))
    full = ""
    for s in xl.sheet_names:
        df    = xl.parse(s).dropna(how="all")
        full += f"\n[Sheet: {s}]\n{df.to_string(index=False)}\n"
    return {"filename":name,"type":"EXCEL",
            "page_count":f"{len(xl.sheet_names)} sheets",
            "full_text":full,"word_count":len(full.split()),
            "hash":hashlib.sha256(raw).hexdigest()[:10]}


PARSERS = {
    ".pdf":  extract_pdf,
    ".docx": extract_docx,
    ".doc":  extract_docx,
    ".xlsx": extract_excel,
    ".xls":  extract_excel,
}


def parse_file(f) -> dict:
    raw = f.read()
    ext = Path(f.name).suffix.lower()
    parser = PARSERS.get(ext)
    if parser:
        return parser(raw, f.name)
    return {"error": f"Unsupported: {ext}"}


DATA_DIR = Path("data")

def auto_load_data_dir():
    """Load all supported files from the data/ directory on first run."""
    if not DATA_DIR.is_dir():
        return
    for filepath in sorted(DATA_DIR.iterdir()):
        ext = filepath.suffix.lower()
        if ext in PARSERS and filepath.name not in st.session_state.documents:
            try:
                raw = filepath.read_bytes()
                parsed = PARSERS[ext](raw, filepath.name)
                if "error" not in parsed:
                    st.session_state.documents[filepath.name] = parsed
            except Exception:
                pass  # skip files that fail to parse
    if st.session_state.documents:
        st.session_state.doc_context = rebuild_context()


def rebuild_context() -> str:
    if not st.session_state.documents:
        return ""
    ctx  = f"LOAN FILE: {st.session_state.loan_id}\nDOCUMENTS: {len(st.session_state.documents)}\n{'='*60}\n\n"
    used = len(ctx)
    for name, doc in st.session_state.documents.items():
        if "error" in doc:
            continue
        hdr   = f"DOCUMENT: {name}  |  {doc['type']}  |  {doc.get('page_count','?')} pages\n{'-'*50}\n"
        space = 80000 - used - len(hdr) - 100
        if space < 300:
            ctx += "[Context limit reached]\n"
            break
        ctx  += hdr + doc.get("full_text","")[:space] + "\n\n"
        used  = len(ctx)
    return ctx

# ─────────────────────────────────────────────────────────────────────────────
# SYSTEM PROMPT
# ─────────────────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are the Credit Intelligence AI Copilot for a corporate banking Loan Origination System (LOS).

You assist credit officers in analyzing corporate loan applications.

YOUR RULES (strictly follow every one):

1. ANSWER ONLY FROM THE UPLOADED DOCUMENTS
   Use ONLY the document context provided below. Never fill gaps with general knowledge.
   You must NOT invent, assume, or use any data that is not explicitly present in the uploaded documents.
   If something is not in the documents say: "⚠️ Not available in the current documents."

2. CITE EVERY CLAIM
   Every factual statement must end with a citation tag:
   - For PDF:   📄 [Source: <filename>, Page <N>]
   - For Excel: 📊 [Source: <filename>, Sheet: <sheetname>]
   - For Word:  📝 [Source: <filename>]
   No citation = do not make the claim.

3. CALCULATIONS AND WHAT-IF SCENARIOS
   Use this structure every time:

   📥 Inputs:
   • [value 1] — from 📄 [Source: ...]
   • [value 2] — from 📄 [Source: ...]

   🔢 Calculation:
   • Step 1: ...
   • Step 2: ...
   • Result: [final number]

   📌 Assessment:
   [What this means for the credit decision]

4. RISK FLAGS
   🔴 HIGH RISK — [what and why]
   🟡 MEDIUM RISK — [what and why]
   🟢 POSITIVE — [what and why]

5. CONVERSATION MEMORY
   You remember previous questions in this conversation. Refer back when relevant.

6. MISSING DATA
   If key data is missing, say what you found and what document type would normally contain the missing info.

7. DISCLAIMER — end EVERY response with this line:
   ⚠️ AI Disclaimer: For decision support only. Verify all findings independently. AI does not approve or reject loans.

LOAN DOCUMENTS PROVIDED:
{doc_context}
"""

# ─────────────────────────────────────────────────────────────────────────────
# AI CALL
# ─────────────────────────────────────────────────────────────────────────────

def get_api_key() -> str:
    """Get API key from environment variable."""
    return os.environ.get("OPENAI_API_KEY", "")


def build_ai_messages(user_msg: str) -> list:
    """Build the message list for the OpenAI API call."""
    # Always rebuild context from current documents to ensure it's fresh
    st.session_state.doc_context = rebuild_context()
    doc_ctx = st.session_state.doc_context or "No documents loaded."

    # List the document names explicitly so the model knows what it has
    if st.session_state.documents:
        doc_list = "\n".join(
            f"  - {name} ({doc.get('type','?')}, {doc.get('word_count',0)} words)"
            for name, doc in st.session_state.documents.items()
            if "error" not in doc
        )
        doc_ctx = f"You have access to EXACTLY these {len(st.session_state.documents)} document(s):\n{doc_list}\n\nFull document contents below:\n{doc_ctx}"

    system = SYSTEM_PROMPT.replace("{doc_context}", doc_ctx)
    msgs = [{"role": "system", "content": system}]
    for m in st.session_state.messages[-20:]:
        msgs.append({"role": m["role"], "content": m["content"]})
    msgs.append({"role": "user", "content": user_msg})
    return msgs


def call_ai_stream(user_msg: str, api_key: str, model: str):
    """Stream the AI response, yielding chunks of text."""
    client = openai.OpenAI(api_key=api_key)
    msgs = build_ai_messages(user_msg)

    stream = client.chat.completions.create(
        model=model, messages=msgs, temperature=0.1, max_tokens=2500,
        stream=True,
    )

    first_token = True
    for chunk in stream:
        if chunk.choices and chunk.choices[0].delta.content:
            if first_token:
                first_token = False
            yield chunk.choices[0].delta.content

    # If no tokens were yielded, yield a fallback message
    if first_token:
        yield "No response generated. Please try again."

# ─────────────────────────────────────────────────────────────────────────────
# AUTO-LOAD documents from data/ directory on first run
# ─────────────────────────────────────────────────────────────────────────────
if not st.session_state.data_dir_loaded:
    auto_load_data_dir()
    st.session_state.data_dir_loaded = True

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="title">Credit AI Copilot</div>
        <div class="subtitle">CORPORATE LOAN ORIGINATION</div>
    </div>""", unsafe_allow_html=True)

    # API key — from .env or manual override
    env_key = get_api_key()
    st.markdown("#### Configuration")

    if env_key:
        st.success("✅ API Key loaded from .env")
        api_key = env_key
    else:
        api_key = st.text_input("OpenAI API Key", type="password",
                                placeholder="sk-...",
                                help="Set OPENAI_API_KEY in .env or enter here")
    api_ready = bool(api_key and api_key.startswith("sk-"))

    # Model selection
    st.session_state.model = st.selectbox(
        "OpenAI Model",
        OPENAI_MODELS,
        index=OPENAI_MODELS.index(st.session_state.model),
        help="Select the OpenAI model to use for analysis",
    )

    st.divider()

    # Loan context
    st.markdown("#### Loan Context")
    st.session_state.loan_id = st.text_input("Loan ID",      value=st.session_state.loan_id)
    st.session_state.officer = st.text_input("Credit Officer", value=st.session_state.officer)

    st.divider()

    # ── Simplified Document Upload ──────────────────────────────────────────
    st.markdown("#### Documents")
    st.caption("Drag & drop or click to upload — PDF, Word, Excel")
    uploaded = st.file_uploader(
        "Upload loan documents",
        type=["pdf","docx","doc","xlsx","xls"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    # Auto-process uploaded files and save to data/ for persistence
    if uploaded:
        new_count = 0
        for f in uploaded:
            if f.name not in st.session_state.documents:
                raw = f.read()
                f.seek(0)  # reset for parse_file
                parsed = parse_file(f)
                if "error" not in parsed:
                    st.session_state.documents[f.name] = parsed
                    # Save to data/ so it persists across redeployments
                    DATA_DIR.mkdir(exist_ok=True)
                    (DATA_DIR / f.name).write_bytes(raw)
                    new_count += 1
                else:
                    st.error(f"✗ {f.name}: {parsed['error']}")
        if new_count > 0:
            st.session_state.doc_context = rebuild_context()
            st.toast(f"✓ {new_count} document(s) loaded and saved", icon="📄")

    # Show loaded documents
    if st.session_state.documents:
        st.markdown(f"**{len(st.session_state.documents)} document(s) loaded**")
        for name, doc in st.session_state.documents.items():
            icon = {"PDF":"📕","WORD":"📘","EXCEL":"📗"}.get(doc.get("type",""),"📄")
            st.markdown(f"""<div class="doc-card">
                <strong>{icon} {name[:32]}</strong>
                <div class="meta">{doc.get('type')} &middot; {doc.get('page_count','?')} pages
                &middot; {doc.get('word_count',0):,} words</div>
            </div>""", unsafe_allow_html=True)
        if st.button("🗑 Clear All Documents", use_container_width=True):
            # Remove files from data/ directory
            for name in st.session_state.documents:
                fpath = DATA_DIR / name
                if fpath.exists():
                    fpath.unlink()
            st.session_state.documents      = {}
            st.session_state.doc_context    = ""
            st.session_state.messages       = []
            st.session_state.data_dir_loaded = True  # don't re-load cleared files
            st.rerun()
    else:
        st.info("No documents loaded")

    st.divider()

    # Conversation controls
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🔄 New Chat", use_container_width=True):
            st.session_state.messages = []
            st.rerun()
    with col2:
        q_count = len([m for m in st.session_state.messages if m["role"]=="user"])
        st.metric("Queries", q_count)

    st.divider()
    st.caption("Production: integrates with TCS BaNCS via REST API · "
               "On-prem LLM via TCS AI Compass · All queries audit-logged to PostgreSQL")

# ─────────────────────────────────────────────────────────────────────────────
# MAIN — CHAT INTERFACE
# ─────────────────────────────────────────────────────────────────────────────

doc_cnt   = len(st.session_state.documents)
doc_words = sum(d.get("word_count",0) for d in st.session_state.documents.values())

# Top banner with stat pills
q_count = len([m for m in st.session_state.messages if m["role"] == "user"])
st.markdown(f"""
<div class="top-banner">
    <h2>Credit Intelligence AI Copilot</h2>
    <p>Intelligent document analysis for credit decisions</p>
    <div class="stat-row">
        <div class="stat-pill">Loan <span class="val">{st.session_state.loan_id}</span></div>
        <div class="stat-pill">Docs <span class="val">{doc_cnt}</span></div>
        <div class="stat-pill">Words <span class="val">{doc_words:,}</span></div>
        <div class="stat-pill">Model <span class="val">{st.session_state.model}</span></div>
        <div class="stat-pill">Queries <span class="val">{q_count}</span></div>
    </div>
</div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CHAT SUGGESTIONS — contextual suggested questions
# ─────────────────────────────────────────────────────────────────────────────

SUGGESTIONS = [
    ("📋 Risk Summary",
     "Summarize all key risk factors identified across the uploaded documents. Highlight any red flags."),
    ("💰 Financial Highlights",
     "What are the key financial highlights? Include revenue, profit, net worth, debt metrics for all available years."),
    ("🏠 Collateral Analysis",
     "Analyze the collateral for this loan. What is the value, LTV ratio, and any concerns?"),
    ("📊 Credit Bureau",
     "What does the credit bureau report show? Summarize credit score, repayment history, existing loans and defaults."),
    ("👤 Guarantor Assessment",
     "Who are the guarantors? What is their net worth and capacity to support the guarantee?"),
    ("⚠️ Stress Test: Collateral -10%",
     "What-if scenario: If the collateral value decreases by 10%, calculate the revised LTV. Show the calculation step by step and state if any threshold is breached."),
]

# ── Welcome screen with suggestions ─────────────────────────────────────────
if not st.session_state.messages:
    st.markdown("""
    <div class="welcome-box">
        <div class="welcome-icon">&#x1F3E6;</div>
        <div class="welcome-title">Credit Intelligence AI Copilot</div>
        <div class="welcome-desc">
            Upload loan documents in the sidebar, then ask me anything about the loan file.
            I answer <strong>only from uploaded documents</strong> with full citations and risk analysis.
        </div>
    </div>""", unsafe_allow_html=True)

    # Suggestion chips on welcome screen
    st.markdown("#### Suggested Questions")
    cols = st.columns(3)
    for i, (label, prompt) in enumerate(SUGGESTIONS):
        with cols[i % 3]:
            if st.button(label, key=f"sug_{i}", use_container_width=True):
                st.session_state.prefill = prompt
                st.rerun()

# ── Render conversation history ──────────────────────────────────────────────
for msg in st.session_state.messages:
    if msg["role"] == "user":
        with st.chat_message("user", avatar="👤"):
            st.markdown(msg["content"])

    elif msg["role"] == "assistant":
        with st.chat_message("assistant", avatar="🤖"):
            st.markdown(msg["content"])
            meta = msg.get("meta", {})
            if meta:
                elapsed = meta.get("elapsed_ms", "—")
                model   = meta.get("model", "—")
                qid     = meta.get("query_id", "—")
                st.caption(f"Model: {model}  |  {elapsed}ms  |  Audit ID: {qid}")

# ── Follow-up suggestions after AI response ─────────────────────────────────
if st.session_state.messages and st.session_state.messages[-1]["role"] == "assistant":
    st.markdown("#### Ask next")
    follow_ups = [
        ("Risk Summary",  SUGGESTIONS[0][1]),
        ("Financials",    SUGGESTIONS[1][1]),
        ("Credit Bureau", SUGGESTIONS[3][1]),
        ("Stress Test",   SUGGESTIONS[5][1]),
    ]
    cols = st.columns(4)
    for i, (label, prompt) in enumerate(follow_ups):
        with cols[i]:
            if st.button(label, key=f"fu_{i}", use_container_width=True):
                st.session_state.prefill = prompt
                st.rerun()

# ── Status line ──────────────────────────────────────────────────────────────
if not api_ready:
    st.info("Set OPENAI_API_KEY in .env file or enter it in the sidebar.")
elif not st.session_state.documents:
    st.info("Upload at least one loan document in the sidebar to get started.")

# ── Chat input — always render this widget ───────────────────────────────────
user_input = st.chat_input(
    placeholder="Ask anything about this loan...",
    disabled=(not api_ready or not st.session_state.documents),
)

# If a suggestion button was clicked, use that as input instead
if st.session_state.prefill:
    user_input = st.session_state.prefill
    st.session_state.prefill = ""

# ── Handle send ──────────────────────────────────────────────────────────────
if user_input:
    # Show user message immediately
    st.session_state.messages.append({"role": "user", "content": user_input.strip()})
    with st.chat_message("user", avatar="👤"):
        st.markdown(user_input.strip())

    # Stream AI response with visible status
    ai_error = False
    with st.chat_message("assistant", avatar="🤖"):
        t0 = time.time()
        try:
            with st.status(f"Connecting to {st.session_state.model}...", expanded=True) as status:
                status.write(f"Sending query to OpenAI ({st.session_state.model})...")
                status.write(f"Documents in context: {len(st.session_state.documents)}")
                # Create the stream generator
                stream_gen = call_ai_stream(user_input.strip(), api_key, st.session_state.model)
                # Get the first chunk to confirm connection
                first_chunk = next(stream_gen, None)
                if first_chunk:
                    status.update(label="Receiving response...", state="running")
                    status.write("Streaming tokens from OpenAI...")

            # Now stream the full response (first chunk + rest)
            def full_stream():
                if first_chunk:
                    yield first_chunk
                yield from stream_gen

            full_response = st.write_stream(full_stream())

            elapsed = int((time.time() - t0) * 1000)
            qid = hashlib.sha256(f"{user_input}{time.time()}".encode()).hexdigest()[:12]
            meta = {
                "elapsed_ms": elapsed,
                "model": st.session_state.model,
                "timestamp": datetime.utcnow().isoformat(),
                "query_id": qid,
            }
            st.caption(f"Model: {st.session_state.model}  |  {elapsed}ms  |  Audit ID: {qid}")
            st.session_state.messages.append({
                "role": "assistant",
                "content": full_response,
                "meta": meta,
            })
        except openai.AuthenticationError:
            ai_error = True
            st.error("**Authentication Failed** — Invalid API key. Check your `.env` file or sidebar input.")
        except openai.RateLimitError:
            ai_error = True
            st.error("**Rate Limit Exceeded** — Please wait 30 seconds and try again.")
        except openai.APIConnectionError:
            ai_error = True
            st.error("**Connection Failed** — Could not reach OpenAI. Check your internet connection.")
        except openai.NotFoundError:
            ai_error = True
            st.error(f"**Model Not Found** — `{st.session_state.model}` is not available on your API plan. Try a different model.")
        except openai.APIStatusError as e:
            ai_error = True
            st.error(f"**API Error ({e.status_code})** — {e.message}")
        except Exception as e:
            ai_error = True
            st.error(f"**Unexpected Error** — {type(e).__name__}: {e}")

    if ai_error:
        # Remove the user message that got no AI reply so it doesn't look orphaned
        if st.session_state.messages and st.session_state.messages[-1]["role"] == "user":
            st.session_state.messages.pop()
        # Store error in session so it persists across reruns
        st.session_state["_last_error"] = True
    else:
        st.session_state.pop("_last_error", None)
        st.rerun()
