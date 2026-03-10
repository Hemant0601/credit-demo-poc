"""
Credit Intelligence AI Copilot — Conversational Chat Demo
==========================================================
ChatGPT-style interface for credit officers.
AI has full access to uploaded loan documents and answers with citations.

Requirements:
    pip install streamlit pymupdf openai python-docx openpyxl pandas

Run:
    streamlit run demo_app.py
"""

import streamlit as st
import fitz                  # PyMuPDF
import openai
import hashlib
import re
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path

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
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Credit Intelligence AI Copilot",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background: #f0f4f8; }

    .top-banner {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        padding: 14px 24px; border-radius: 10px; margin-bottom: 16px;
    }
    .top-banner h2 { color: white; margin: 0; font-size: 20px; }
    .top-banner p  { color: #cce4f7; margin: 2px 0 0 0; font-size: 13px; }

    .chat-user {
        display: flex; justify-content: flex-end; margin: 12px 0;
    }
    .bubble-user {
        background: #2E75B6; color: white;
        padding: 12px 16px; border-radius: 18px 18px 4px 18px;
        max-width: 72%; font-size: 14px; line-height: 1.6;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15);
    }

    .chat-ai {
        display: flex; justify-content: flex-start;
        margin: 12px 0; align-items: flex-start; gap: 10px;
    }
    .ai-avatar {
        background: #1F3864; color: white;
        width: 36px; height: 36px; border-radius: 50%;
        display: flex; align-items: center; justify-content: center;
        font-size: 16px; flex-shrink: 0; margin-top: 4px;
    }
    .bubble-ai {
        background: white; color: #1a1a2e;
        padding: 14px 18px; border-radius: 4px 18px 18px 18px;
        max-width: 82%; font-size: 14px; line-height: 1.8;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        border-left: 3px solid #2E75B6;
    }
    .msg-meta {
        margin-top: 5px; font-size: 11px; color: #aaa; padding-left: 4px;
    }
    .audit-pill {
        display: inline-block; background: #e8f5e9; color: #2e7d32;
        border-radius: 12px; padding: 1px 8px; font-size: 11px;
    }
    .src-tag-pdf  { background:#fff3cd; color:#856404; border-radius:4px; padding:1px 6px; font-size:12px; white-space:nowrap; }
    .src-tag-xls  { background:#e8f5e9; color:#2e7d32; border-radius:4px; padding:1px 6px; font-size:12px; white-space:nowrap; }
    .src-tag-doc  { background:#e3f2fd; color:#0d47a1; border-radius:4px; padding:1px 6px; font-size:12px; white-space:nowrap; }
    .risk-high    { background:#fde8e8; color:#c62828; padding:2px 8px; border-radius:4px; font-weight:600; }
    .risk-med     { background:#fff8e1; color:#f57f17; padding:2px 8px; border-radius:4px; font-weight:600; }
    .risk-pos     { background:#e8f5e9; color:#2e7d32; padding:2px 8px; border-radius:4px; font-weight:600; }

    .doc-card {
        background: white; border-radius: 8px; padding: 8px 12px;
        margin: 5px 0; border-left: 3px solid #2E75B6; font-size: 12px;
    }
    .doc-card strong { font-size: 13px; color: #1F3864; }

    .welcome-box {
        text-align: center; padding: 50px 20px; color: #888;
    }
    footer { visibility: hidden; }
    #MainMenu { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
for k, v in {
    "messages":    [],
    "documents":   {},
    "loan_id":     "LOAN-2024-CORP-001",
    "officer":     "Credit Officer",
    "prefill":     "",
    "doc_context": "",
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


def parse_file(f) -> dict:
    raw = f.read()
    ext = Path(f.name).suffix.lower()
    if ext == ".pdf":             return extract_pdf(raw, f.name)
    if ext in (".docx",".doc"):   return extract_docx(raw, f.name)
    if ext in (".xlsx",".xls"):   return extract_excel(raw, f.name)
    return {"error": f"Unsupported: {ext}"}


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

1. ANSWER ONLY FROM DOCUMENTS
   Use only the document context provided. Never fill gaps with general knowledge.
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

def call_ai(user_msg: str, api_key: str) -> dict:
    client = openai.OpenAI(api_key=api_key)

    system = SYSTEM_PROMPT.replace(
        "{doc_context}", st.session_state.doc_context or "No documents loaded."
    )

    msgs = [{"role":"system","content":system}]
    for m in st.session_state.messages[-20:]:
        msgs.append({"role":m["role"],"content":m["content"]})
    msgs.append({"role":"user","content":user_msg})

    t0   = time.time()
    resp = client.chat.completions.create(
        model="gpt-4o", messages=msgs, temperature=0.1, max_tokens=2500
    )
    elapsed = int((time.time()-t0)*1000)
    answer  = resp.choices[0].message.content

    return {
        "answer":            answer,
        "elapsed_ms":        elapsed,
        "prompt_tokens":     resp.usage.prompt_tokens,
        "completion_tokens": resp.usage.completion_tokens,
        "timestamp":         datetime.utcnow().isoformat(),
        "query_id":          hashlib.sha256(f"{user_msg}{time.time()}".encode()).hexdigest()[:12],
    }

# ─────────────────────────────────────────────────────────────────────────────
# FORMAT AI ANSWER — highlight citations and risk flags
# ─────────────────────────────────────────────────────────────────────────────

def format_answer(text: str) -> str:
    html = text.replace("\n", "<br>")
    # PDF citations
    html = re.sub(r'📄\s*\[Source:([^\]]+)\]',
        r'<span class="src-tag-pdf">📄 Source:\1</span>', html)
    # Excel citations
    html = re.sub(r'📊\s*\[Source:([^\]]+)\]',
        r'<span class="src-tag-xls">📊 Source:\1</span>', html)
    # Word citations
    html = re.sub(r'📝\s*\[Source:([^\]]+)\]',
        r'<span class="src-tag-doc">📝 Source:\1</span>', html)
    # Risk flags
    html = re.sub(r'🔴 (HIGH RISK[^\n<]*)',
        r'<span class="risk-high">🔴 \1</span>', html)
    html = re.sub(r'🟡 (MEDIUM RISK[^\n<]*)',
        r'<span class="risk-med">🟡 \1</span>', html)
    html = re.sub(r'🟢 (POSITIVE[^\n<]*)',
        r'<span class="risk-pos">🟢 \1</span>', html)
    return html

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("""
    <div style="background:linear-gradient(135deg,#1F3864,#2E75B6);
                border-radius:10px;padding:14px 16px;margin-bottom:14px">
        <div style="color:white;font-size:17px;font-weight:700">🏦 Credit AI Copilot</div>
        <div style="color:#cce4f7;font-size:12px">Corporate Loan Origination · POC</div>
    </div>""", unsafe_allow_html=True)

    # API key
    st.markdown("#### ⚙️ OpenAI API Key")
    api_key   = st.text_input("", type="password", placeholder="sk-...",
                               label_visibility="collapsed")
    api_ready = bool(api_key and api_key.startswith("sk-"))
    st.markdown("✅ API Key set" if api_ready else "⚠️ API key required")

    st.divider()

    # Loan context
    st.markdown("#### 📋 Loan Context")
    st.session_state.loan_id = st.text_input("Loan ID",      value=st.session_state.loan_id)
    st.session_state.officer = st.text_input("Credit Officer", value=st.session_state.officer)

    st.divider()

    # Document upload
    st.markdown("#### 📂 Upload Documents")
    st.caption("PDF · Word · Excel")
    uploaded = st.file_uploader("", type=["pdf","docx","doc","xlsx","xls"],
                                  accept_multiple_files=True,
                                  label_visibility="collapsed")
    if uploaded:
        for f in uploaded:
            if f.name not in st.session_state.documents:
                with st.spinner(f"Parsing {f.name}..."):
                    parsed = parse_file(f)
                    if "error" not in parsed:
                        st.session_state.documents[f.name]  = parsed
                        st.session_state.doc_context        = rebuild_context()
                        st.toast(f"✓ {f.name}", icon="📄")
                    else:
                        st.error(f"✗ {parsed['error']}")

    if st.session_state.documents:
        for name, doc in st.session_state.documents.items():
            icon = {"PDF":"📕","WORD":"📘","EXCEL":"📗"}.get(doc.get("type",""),"📄")
            st.markdown(f"""<div class="doc-card">
                <strong>{icon} {name[:28]}</strong><br>
                <span style="color:#666">{doc.get('type')} · {doc.get('page_count','?')} pages
                · {doc.get('word_count',0):,} words</span>
            </div>""", unsafe_allow_html=True)
        if st.button("🗑 Clear All", use_container_width=True):
            st.session_state.documents   = {}
            st.session_state.doc_context = ""
            st.session_state.messages    = []
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

# Top banner
st.markdown(f"""
<div class="top-banner">
    <h2>🏦 Credit Intelligence AI Copilot</h2>
    <p>Loan: <strong>{st.session_state.loan_id}</strong> &nbsp;·&nbsp;
       {doc_cnt} document(s) · {doc_words:,} words in context &nbsp;·&nbsp;
       Officer: {st.session_state.officer}</p>
</div>""", unsafe_allow_html=True)

# Quick prompts
st.markdown("**Quick Prompts:**")
c1,c2,c3 = st.columns(3)
quick_prompts = [
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
    ("⚠️ What-If: Collateral -10%",
     "What-if scenario: If the collateral value decreases by 10%, calculate the revised LTV. Show the calculation step by step and state if any threshold is breached."),
]
for i,(label,prompt) in enumerate(quick_prompts):
    col = [c1,c2,c3][i%3]
    with col:
        if st.button(label, key=f"q{i}", use_container_width=True):
            st.session_state.prefill = prompt

st.divider()

# ── Welcome screen ───────────────────────────────────────────────────────────
if not st.session_state.messages:
    st.markdown("""
    <div class="welcome-box">
        <div style="font-size:52px">🤖</div>
        <div style="font-size:18px;font-weight:600;color:#1F3864;margin:12px 0">
            Hello! I'm your Credit Intelligence AI Copilot.
        </div>
        <div style="font-size:14px;max-width:540px;margin:0 auto;line-height:1.8;color:#555">
            Upload the loan documents in the sidebar (PDF, Word, Excel).<br>
            Then ask me anything about the loan file.<br><br>
            I will answer <strong>only from the uploaded documents</strong>,
            cite every claim with its source and page number,
            and show full step-by-step reasoning for calculations and what-if scenarios.
        </div>
        <div style="margin-top:20px;color:#aaa;font-size:13px">
            Try a Quick Prompt above or type your own question below ↓
        </div>
    </div>""", unsafe_allow_html=True)

# ── Render conversation ───────────────────────────────────────────────────────
for msg in st.session_state.messages:
    if msg["role"] == "user":
        st.markdown(f"""
        <div class="chat-user">
            <div class="bubble-user">
                <div style="font-size:11px;opacity:.8;margin-bottom:5px">
                    👤 {st.session_state.officer}
                </div>
                {msg["content"]}
            </div>
        </div>""", unsafe_allow_html=True)

    elif msg["role"] == "assistant":
        meta    = msg.get("meta", {})
        elapsed = meta.get("elapsed_ms","—")
        tokens  = meta.get("completion_tokens","—")
        qid     = meta.get("query_id","—")
        ts      = meta.get("timestamp","")[:19]
        answer  = format_answer(msg["content"])

        st.markdown(f"""
        <div class="chat-ai">
            <div class="ai-avatar">🤖</div>
            <div style="flex:1;min-width:0">
                <div class="bubble-ai">
                    <div style="font-size:11px;color:#888;margin-bottom:8px">
                        <strong style="color:#1F3864">Credit AI Copilot</strong>
                        &nbsp;·&nbsp; {ts}
                    </div>
                    {answer}
                </div>
                <div class="msg-meta">
                    ⏱ {elapsed}ms &nbsp;|&nbsp; 🔤 {tokens} tokens &nbsp;|&nbsp;
                    <span class="audit-pill">✓ Audit logged · ID: {qid}</span>
                </div>
            </div>
        </div>""", unsafe_allow_html=True)

# ── Input row ────────────────────────────────────────────────────────────────
st.markdown("---")
inp_col, btn_col = st.columns([5,1])

with inp_col:
    user_input = st.text_area(
        "msg", value=st.session_state.prefill, height=90,
        placeholder=(
            "Ask anything about this loan...  e.g.  "
            "'What are the risk factors?'  |  "
            "'Calculate DSCR'  |  "
            "'If collateral drops 15%, what happens to LTV?'"
        ),
        label_visibility="collapsed",
    )
    st.session_state.prefill = ""

with btn_col:
    st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)
    send = st.button(
        "Send ➤", type="primary", use_container_width=True,
        disabled=(not user_input.strip() or not api_ready or not st.session_state.documents),
    )

# Status line
if not api_ready:
    st.caption("⚠️ Enter your OpenAI API key in the sidebar.")
elif not st.session_state.documents:
    st.caption("⚠️ Upload at least one loan document in the sidebar.")
else:
    st.caption(f"✅ Ready · {doc_cnt} doc(s) · {doc_words:,} words · "
               "Shift+Enter for new line · Every answer is cited and audit-logged")

# ── Handle send ──────────────────────────────────────────────────────────────
if send and user_input.strip():
    st.session_state.messages.append({"role":"user","content":user_input.strip()})

    with st.spinner("🤖 Analyzing documents and preparing cited response..."):
        try:
            result = call_ai(user_input.strip(), api_key)
            st.session_state.messages.append({
                "role":    "assistant",
                "content": result["answer"],
                "meta":    result,
            })
        except openai.AuthenticationError:
            st.error("❌ Invalid API key.")
        except openai.RateLimitError:
            st.error("❌ Rate limit — wait 30 seconds and retry.")
        except Exception as e:
            st.error(f"❌ {e}")

    st.rerun()
