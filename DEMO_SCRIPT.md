# Credit Intelligence AI Copilot — Demo Script
## Recorded Video Walkthrough Guide

---

## BEFORE YOU RECORD — CHECKLIST

- [ ] App is running at http://localhost:8501
- [ ] OpenAI API key entered in sidebar
- [ ] 2-3 sample documents uploaded (valuation PDF + any Word/Excel)
- [ ] Loan ID set to something realistic (e.g. LOAN-2024-CORP-001)
- [ ] Screen resolution: 1920x1080 preferred
- [ ] Browser: Chrome, full screen, no personal bookmarks visible
- [ ] Microphone tested
- [ ] Close all other apps (no notifications)
- [ ] Run through once without recording to check responses look good

---

## DEMO SCRIPT (Target: 8-10 minutes)

---

### SEGMENT 1 — OPENING (60 seconds)

**Say:**
"Good [morning/afternoon]. I'm going to walk you through a Proof of Concept
for the Credit Intelligence AI Copilot — an AI-assisted decision support tool
that we're building to integrate directly into our Loan Origination System.

The core problem it solves: a credit officer reviewing a corporate loan
application typically has to read through 15 to 20 documents — valuation
reports, legal opinions, bank statements, credit bureau reports.
That takes hours, and critical risk signals can be missed.

This AI Copilot reads all those documents, understands the relationships
between entities, and answers the credit officer's questions in plain English —
with full citations back to the source document and page.

Let me show you how it works."

**[Show the main screen — 'Ask AI Copilot' tab]**

---

### SEGMENT 2 — DOCUMENT INGESTION (90 seconds)

**Say:**
"The first thing a credit officer does is upload the documents for a loan.
I'm going to upload [name your documents] for Loan [your loan ID]."

**[Drag and drop your PDF into the sidebar uploader]**

**Say:**
"The system immediately parses the document — extracting text page by page,
identifying tables, capturing metadata. You can see in the sidebar that
it's extracted [X] words from [X] pages."

**[Upload a second document if you have one]**

**Say:**
"We support PDF, Word, and Excel — the three most common formats in
any corporate lending workflow. In the production system, documents
are automatically ingested from the LOS via batch jobs — no manual upload needed."

**[Point to the Document Viewer tab briefly]**

"The Document Viewer tab shows exactly what text was extracted from each
document — this is important for transparency. The AI never sees more
than what's here."

---

### SEGMENT 3 — FIRST QUERY: RISK SUMMARY (2 minutes)

**Say:**
"Now let me ask the most common question a credit officer asks."

**[Click the quick question button: "Summarize the key risk factors"]**

**Say:**
"I'm asking the AI to summarize the key risk factors from the uploaded documents.
Notice what it's doing — it's not searching the internet, it's not using
general knowledge. It is reading only the documents I've uploaded."

**[Wait for response — 10-20 seconds]**

**Say:**
"Look at the response. Every single claim has a citation — [Source: document name, Page X].
This is non-negotiable in a regulated environment. If a credit officer can't
trace an AI statement back to a specific page in a specific document, that
finding has no evidentiary value.

And see at the bottom — the mandatory disclaimer. The AI explicitly states
this is for decision support only and requires human verification.
The AI does not approve or reject loans. The credit officer does."

---

### SEGMENT 4 — WHAT-IF SCENARIO (90 seconds)

**Say:**
"Now let me show you something more powerful — What-If analysis."

**[Click: "If collateral value decreases by 10%, what is the revised LTV ratio?"]**

**Say:**
"This is a classic credit stress test. If the property market drops,
does our collateral still adequately cover the loan?"

**[Wait for response]**

**Say:**
"The AI pulls the current collateral value from the valuation report,
pulls the outstanding loan amount, applies the 10% stress, and gives
you the revised LTV — with the source citations for each number it used.

In the production system, this computation runs against structured data
extracted and stored in PostgreSQL — so it's even faster and more precise.
The AI then explains the result in plain English for the credit officer."

---

### SEGMENT 5 — AUDIT LOG (60 seconds)

**Say:**
"One of our core design requirements is full regulatory auditability.
Let me show you the audit trail."

**[Switch to 'Query History & Audit' tab]**

**Say:**
"Every single query asked to this AI Copilot is logged — the query text,
which documents were used, how many tokens were consumed, the response,
and the timestamp.

In the production system, this is written to a PostgreSQL table called
ai_query_log with an append-only trigger — meaning it cannot be modified
or deleted. Any AI recommendation can be traced back to:
who asked it, when, what documents were analyzed, and what the AI said.

This is the audit trail that regulators and internal audit require."

**[Click 'Download Audit Log' to show the JSON]**

"And the audit log is exportable for regulatory submissions."

---

### SEGMENT 6 — PRODUCTION ARCHITECTURE (90 seconds)

**[Switch to 'About This System' tab]**

**Say:**
"Let me close with where this is going.

What you've seen today is the POC — a focused demo showing the core
value: document ingestion, grounded AI responses, citations, audit trail.

The production architecture replaces the direct GPT-4o call with a
full RAG pipeline — Retrieval Augmented Generation — backed by three
specialized databases.

PostgreSQL stores structured facts — financial figures, extracted data points.
Neo4j, a graph database, stores entity relationships — who guarantees what,
which director is connected to which companies, which collateral secures
multiple loans.
And pgvector, inside PostgreSQL, stores semantic embeddings of document text
for intelligent search.

The LLM layer — currently OpenAI GPT-4o for the POC — will be replaced by
TCS AI Compass running on-premise. No data leaves the bank's infrastructure.

The final integration point is an embedded widget inside TCS BaNCS —
so the credit officer sees the AI Copilot panel right next to the loan
application, without switching systems."

---

### SEGMENT 7 — CLOSING (30 seconds)

**Say:**
"To summarize what this POC demonstrates:
- Multi-format document ingestion works
- AI responses are grounded and cited — no hallucination
- Audit trail is built-in from day one
- The architecture scales from this POC to full production

The next step is Phase 1 — building the full PDF ingestion pipeline
with vector embeddings, which we expect to complete in [X weeks].

Thank you."

---

## TIPS FOR A GOOD RECORDING

1. **Slow down** — you know the system, they don't. Pause after each response appears.
2. **Hover over citations** — draw attention to the [Source: ...] markers
3. **If AI gives a poor response** — say "In production, the full RAG pipeline
   gives more precise results because it searches at the chunk level rather than
   full-document context." Then move on.
4. **Don't demo live API errors** — test all questions first. If the API fails,
   re-record that segment.
5. **Zoom in** on the citation text so viewers can read it clearly.

---

## COMMON HARD QUESTIONS FROM STAKEHOLDERS

**Q: How do we know the AI isn't making things up?**
A: Two controls. First, the citation requirement — every claim must cite a source page.
If the AI can't cite it, it must say 'not available'. Second, in production, we use
RAG — the AI only sees retrieved document chunks, not its general training data.

**Q: What if the AI gives wrong information?**
A: That's why it's a copilot, not an autopilot. The credit officer must verify every
AI finding independently. The disclaimer is shown on every response. The AI reduces
reading time from hours to minutes — the human makes the decision.

**Q: Is data secure? Does it go to OpenAI?**
A: In this POC, yes — it uses OpenAI. In production, we switch to TCS AI Compass
running on-premise. No data leaves the bank's infrastructure.

**Q: How long to build the full system?**
A: Phase 1 (PDF pipeline + real RAG) in 3-4 weeks. Full system in 3-4 months.
We're building iteratively — each phase delivers working functionality.

**Q: Can it integrate with our existing Oracle database?**
A: Yes — the design explicitly integrates with Oracle/DB2 via REST API.
The AI layer reads from LOS, never writes to it.
