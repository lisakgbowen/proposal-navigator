import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from io import BytesIO
import re
import time
import os

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

st.set_page_config(
    page_title="Proposal Navigator",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# OPENAI SETUP
# -----------------------------
client = None
api_ready = False
try:
    api_key = None
    try:
        api_key = st.secrets.get("OPENAI_API_KEY", None)
    except Exception:
        api_key = None
    if api_key and OpenAI is not None:
        client = OpenAI(api_key=api_key)
        api_ready = True
except Exception:
    client = None
    api_ready = False

# -----------------------------
# SESSION STATE
# -----------------------------
if "review_phase" not in st.session_state:
    st.session_state.review_phase = "ready"   # ready, running, complete
if "review_results" not in st.session_state:
    st.session_state.review_results = None

# -----------------------------
# STYLING
# -----------------------------
st.markdown("""
<style>
:root{
    --pink:#ef3b73;
    --pink-dark:#d82e65;
    --pink-soft:#fff4f8;
    --border:#f1d4df;
    --ink:#2c1732;
    --muted:#77697a;
    --success:#168f5a;
    --warning:#ba7a00;
    --panel:#ffffff;
}
.stApp{
    background:linear-gradient(180deg,#fff8fb 0%,#ffffff 34%);
}
.block-container{
    max-width:1500px;
    padding-top:2.6rem;
    padding-bottom:1rem;
    padding-left:.9rem;
    padding-right:.9rem;
}
[data-testid="stSidebar"]{
    min-width:170px !important;
    max-width:170px !important;
    background:linear-gradient(180deg,#fff2f6,#fffafa);
    border-right:1px solid var(--border);
}
[data-testid="stSidebar"] .block-container{
    padding:.8rem .65rem;
}

.pn-side-compass{
    font-size:2.6rem;
    line-height:1;
    margin-bottom:.28rem;
    text-align:left;
}

.pn-side-title{
    font-family:Georgia,"Times New Roman",serif;
    color:var(--ink);
    font-size:1.45rem;
    font-weight:700;
    line-height:1;
}
.pn-side-sub{
    font-family:Georgia,"Times New Roman",serif;
    color:var(--pink);
    font-style:italic;
    font-size:.72rem;
    line-height:1.15;
    margin:.25rem 0 .7rem 0;
}
.pn-step{
    display:flex;
    align-items:center;
    gap:.42rem;
    margin:.36rem 0;
    color:var(--ink);
    font-size:.72rem;
    font-weight:600;
}
.pn-dot{
    width:21px;
    height:21px;
    border-radius:50%;
    background:var(--pink);
    color:#fff;
    display:inline-flex;
    align-items:center;
    justify-content:center;
    font-size:.68rem;
    font-weight:700;
    flex:0 0 21px;
}
.pn-help{
    margin-top:.8rem;
    padding:.6rem;
    border:1px solid var(--border);
    border-radius:12px;
    background:rgba(255,255,255,.82);
    font-size:.67rem;
    color:var(--muted);
    line-height:1.35;
}
.pn-help b{color:var(--ink);font-size:.78rem;}
.pn-header{
    border:1px solid var(--border);
    border-radius:18px;
    background:linear-gradient(135deg,#fff7fa 0%,#ffffff 75%);
    padding:.75rem 1rem;
    margin-bottom:.55rem;
}
.pn-title{
    font-family:Georgia,"Times New Roman",serif;
    color:var(--ink);
    font-size:2rem;
    font-weight:700;
    line-height:1.05;
    margin:0;
}
.pn-tag{
    font-family:Georgia,"Times New Roman",serif;
    color:var(--pink);
    font-style:italic;
    font-size:.95rem;
    margin-top:.12rem;
}
.pn-copy{
    color:var(--muted);
    font-size:.72rem;
    margin-top:.35rem;
}
[data-testid="stVerticalBlockBorderWrapper"]{
    background:rgba(255,255,255,.96) !important;
    border-color:var(--border) !important;
    border-radius:16px !important;
    box-shadow:0 3px 12px rgba(82,38,61,.035);
}
[data-testid="stVerticalBlockBorderWrapper"] > div{
    padding:.58rem .7rem !important;
}
.pn-cardhead{
    display:flex;
    align-items:center;
    gap:.45rem;
    margin-bottom:.3rem;
}
.pn-num{
    width:24px;
    height:24px;
    border-radius:50%;
    background:var(--pink);
    color:#fff;
    display:inline-flex;
    align-items:center;
    justify-content:center;
    font-size:.74rem;
    font-weight:700;
}
.pn-cardtitle{
    color:var(--ink);
    font-size:.92rem;
    font-weight:700;
}
[data-testid="stFileUploader"] section{
    min-height:68px;
    border:1.5px dashed #ef8daf;
    background:#fff8fa;
    border-radius:12px;
    padding:.4rem .55rem;
}
[data-testid="stRadio"]{margin-bottom:-.3rem;}
[data-testid="stRadio"] > label{
    font-size:.7rem !important;
    margin-bottom:-.12rem !important;
}
[data-testid="stRadio"] [role="radiogroup"] label{
    font-size:.69rem !important;
}
[data-testid="stMetric"]{
    border:1px solid var(--border);
    background:#fffafb;
    border-radius:12px;
    padding:.45rem .55rem;
    min-height:72px;
}
[data-testid="stMetricLabel"]{font-size:.68rem !important;}
[data-testid="stMetricValue"]{font-size:1.35rem !important;}
[data-testid="stExpander"]{
    border:1px solid var(--border) !important;
    border-radius:10px !important;
    overflow:hidden;
}
[data-testid="stExpander"] summary{
    min-height:31px !important;
    padding:.18rem .35rem !important;
    font-size:.7rem !important;
}
div.stButton > button{
    min-height:36px;
    font-size:.76rem;
    font-weight:700;
    border-radius:10px;
    border:0;
    color:white;
    background:linear-gradient(90deg,var(--pink-dark),var(--pink));
}
div.stButton > button:hover{
    color:white;
    border:0;
}
[data-testid="stDownloadButton"] button{
    min-height:36px;
    font-size:.76rem;
    font-weight:700;
    border-radius:10px;
    border:0;
    color:white;
    background:linear-gradient(90deg,var(--pink-dark),var(--pink));
    width:100%;
}
.stAlert{
    padding:.42rem .55rem !important;
    font-size:.69rem !important;
}
.stCaption{font-size:.64rem !important;}
.pn-mini{
    padding:.5rem .58rem;
    border-radius:10px;
    border:1px solid var(--border);
    background:#fff6f9;
    color:var(--muted);
    font-size:.67rem;
    line-height:1.32;
}
.pn-status{
    padding:.42rem .55rem;
    border-radius:10px;
    border:1px solid var(--border);
    background:#fff7fa;
    color:var(--ink);
    font-size:.69rem;
    margin:.28rem 0;
}
.pn-footer{
    margin-top:.45rem;
    padding:.45rem .55rem;
    border-radius:10px;
    border:1px solid var(--border);
    background:#fff6f9;
    color:var(--muted);
    font-size:.66rem;
}
.pn-readiness{
    --pct:0;
    width:98px;
    height:98px;
    border-radius:50%;
    background:conic-gradient(var(--pink) calc(var(--pct)*1%), #f7dbe5 0);
    display:grid;
    place-items:center;
    margin:.2rem auto .15rem auto;
    position:relative;
}
.pn-readiness::before{
    content:"";
    position:absolute;
    width:72px;
    height:72px;
    border-radius:50%;
    background:white;
}
.pn-readiness-inner{
    position:relative;
    z-index:1;
    text-align:center;
    line-height:1;
    color:var(--ink);
}
.pn-readiness-inner strong{display:block;font-size:1.2rem;}
.pn-readiness-inner span{font-size:.6rem;color:var(--muted);}

/* Review animation */
.pn-review-wrap{
    display:flex;
    justify-content:center;
    align-items:center;
    padding:.15rem 0 .4rem 0;
}
.pn-loader{
    position:relative;
    width:112px;
    height:112px;
}
.pn-loader .orbit-ring{
    position:absolute;
    inset:8px;
    border:3px solid #f8dce5;
    border-top-color:var(--pink);
    border-radius:50%;
}
.pn-loader .compass{
    position:absolute;
    left:50%;
    top:50%;
    transform:translate(-50%,-50%);
    width:54px;
    height:54px;
    border:2px solid #f0c4d5;
    border-radius:50%;
    background:linear-gradient(180deg,#fff7fa,#fff);
    display:flex;
    align-items:center;
    justify-content:center;
    color:var(--pink);
    font-size:1.25rem;
}
.pn-loader .orbiter{
    position:absolute;
    inset:0;
    animation:pnorbit 2.2s linear infinite;
}
.pn-loader .dragonfly{
    position:absolute;
    top:2px;
    left:50%;
    transform:translateX(-50%);
    width:26px;
    height:26px;
}
.pn-loader .dragonfly .body{
    position:absolute;
    left:11px;
    top:5px;
    width:4px;
    height:16px;
    border-radius:999px;
    background:var(--pink);
}
.pn-loader .dragonfly .head{
    position:absolute;
    left:9px;
    top:2px;
    width:8px;
    height:8px;
    border-radius:50%;
    background:var(--pink);
}
.pn-loader .dragonfly .wing1,
.pn-loader .dragonfly .wing2,
.pn-loader .dragonfly .wing3,
.pn-loader .dragonfly .wing4{
    position:absolute;
    width:10px;
    height:6px;
    border:1.6px solid var(--pink);
    background:rgba(255,255,255,.8);
    border-radius:10px;
}
.pn-loader .dragonfly .wing1{left:1px;top:6px;transform:rotate(-24deg);}
.pn-loader .dragonfly .wing2{left:14px;top:6px;transform:rotate(24deg);}
.pn-loader .dragonfly .wing3{left:1px;top:12px;transform:rotate(18deg);}
.pn-loader .dragonfly .wing4{left:14px;top:12px;transform:rotate(-18deg);}
@keyframes pnorbit{to{transform:rotate(360deg)}}
.pn-check-circle{
    width:82px;
    height:82px;
    border-radius:50%;
    border:8px solid #dff4e9;
    color:var(--success);
    display:flex;
    align-items:center;
    justify-content:center;
    font-size:2rem;
    font-weight:700;
    margin:.2rem auto .35rem auto;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# HELPERS
# -----------------------------
def reset_intake():
    st.session_state["personnel"] = None
    st.session_state["travel"] = None
    st.session_state["equipment"] = None
    st.session_state["participant_incentives"] = None
    st.session_state["subawards"] = None
    st.session_state["cost_share"] = None
    st.session_state["indirect_costs"] = None
    st.session_state.review_phase = "ready"
    st.session_state.review_results = None


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def extract_document(uploaded_file):
    document_text = ""
    document_pages = []
    if uploaded_file is None:
        return document_text, document_pages, None

    try:
        file_name = uploaded_file.name.lower()

        if file_name.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            for i, page in enumerate(reader.pages, start=1):
                page_text = page.extract_text() or ""
                if page_text.strip():
                    document_text += page_text + "\n"
                    document_pages.append((i, page_text))
        elif file_name.endswith(".txt"):
            document_text = uploaded_file.read().decode("utf-8", errors="ignore")
            document_pages.append((1, document_text))
        elif file_name.endswith(".docx"):
            doc = Document(BytesIO(uploaded_file.read()))
            paragraphs = [p.text for p in doc.paragraphs if p.text]
            document_text = "\n".join(paragraphs)
            document_pages.append((1, document_text))

        meta = {
            "characters": len(document_text),
            "pages": len(document_pages),
        }
        return document_text, document_pages, meta
    except Exception as e:
        return "", [], {"error": str(e)}


def find_source_match(pages, keywords):
    for page_number, page_text in pages:
        normalized_text = clean_text(page_text)
        lower_text = normalized_text.lower()

        for keyword in keywords:
            normalized_keyword = clean_text(keyword).lower()
            position = lower_text.find(normalized_keyword)
            if position != -1:
                start = max(0, position - 150)
                end = min(len(normalized_text), position + 450)
                return {
                    "found": True,
                    "page": page_number,
                    "keyword": keyword,
                    "excerpt": normalized_text[start:end],
                }

    return {"found": False, "page": None, "keyword": None, "excerpt": None}


def interpret_source_with_ai(category, intake_answer, match):
    if not match["found"]:
        return (
            "No relevant sponsor language was located using the current retrieval method. "
            "Human review is recommended."
        )

    if not api_ready or client is None:
        return (
            f"Source match found on page {match['page']}, but AI interpretation is unavailable. "
            "Review the excerpt manually."
        )

    prompt = f"""
You are assisting with proposal-readiness review in research administration.

Review ONLY the source excerpt below.
Do not invent sponsor requirements.
Do not assume that silence means something is allowed or prohibited.

Proposal category: {category}
User intake answer: {intake_answer}
Source page: {match['page']}

SOURCE EXCERPT:
{match['excerpt']}

Provide a concise interpretation using this exact format:
Sponsor Requirement:
Budget Impact:
Alignment With Intake:
Human Review Needed:
Reason:

Keep the response under 140 words.
If the excerpt is ambiguous or incomplete, say so clearly.
"""

    try:
        response = client.responses.create(
            model="gpt-4.1-mini",
            reasoning={"effort": "low"},
            input=prompt,
        )
        return response.output_text
    except Exception as e:
        return f"AI interpretation unavailable: {e}"


def compute_results(responses, document_pages):
    keyword_map = {
        "Personnel costs": ["salary", "salaries", "personnel", "compensation", "fringe"],
        "Travel": ["travel", "mileage", "airfare", "lodging"],
        "Equipment": ["equipment", "capital equipment"],
        "Participant incentives": ["participant incentive", "incentive", "gift card", "participant support"],
        "Subawards or contracts": ["subaward", "subrecipient", "subcontract", "contract"],
        "Cost share or matching": ["cost share", "cost sharing", "matching", "match requirement"],
        "Indirect costs": ["indirect costs", "indirect cost", "indirect", "f&a", "facilities and administrative"],
    }

    confirmed = []
    clarifications = []
    findings = []
    source_matches = {}
    interpretations = {}
    human_review = []

    for item, answer in responses.items():
        match = find_source_match(document_pages, keyword_map[item])
        source_matches[item] = match

        if answer == "Yes":
            status = "Confirmed"
            budget = "Potential impact"
            follow = "Review source" if match["found"] else "Human review"
            confirmed.append(item)
        elif answer == "No":
            status = "Not included"
            budget = "No current impact"
            follow = "None" if not match["found"] else "Review if applicable"
        elif answer == "Not sure":
            status = "Needs clarification"
            budget = "Unknown"
            follow = "Human review"
            clarifications.append(item)
            human_review.append(item)
        else:
            status = "Unanswered"
            budget = "Unknown"
            follow = "Complete intake"
            clarifications.append(item)
            human_review.append(item)

        # v0 institutional business rules (can be refined later)
        if answer == "Yes" and item in ["Personnel costs", "Subawards or contracts"]:
            if item not in human_review:
                human_review.append(item)

        findings.append({
            "Requirement": item,
            "Status": status,
            "Budget Impact": budget,
            "Source": f"p. {match['page']}" if match["found"] else "No source",
            "Follow-Up": follow,
        })

        interpretations[item] = interpret_source_with_ai(item, answer, match)

    answered_count = sum(v is not None for v in responses.values())
    source_count = sum(1 for m in source_matches.values() if m["found"])
    readiness_pct = round(((answered_count / 7) * 45) + ((source_count / 7) * 30) + ((len(confirmed) / 7) * 25))
    readiness_pct = max(0, min(100, readiness_pct))

    report_lines = [
        "PROPOSAL NAVIGATOR — READINESS REPORT",
        "",
        f"Overall readiness: {readiness_pct}%",
        f"Answered intake items: {answered_count} of 7",
        f"Confirmed areas: {len(confirmed)}",
        f"Clarifications needed: {len(clarifications)}",
        f"Source matches: {source_count}",
        "",
        "DETAILED FINDINGS",
    ]

    for row in findings:
        report_lines.extend([
            "",
            f"Requirement: {row['Requirement']}",
            f"Status: {row['Status']}",
            f"Budget Impact: {row['Budget Impact']}",
            f"Source: {row['Source']}",
            f"Follow-Up: {row['Follow-Up']}",
        ])

    report_lines.extend(["", "SOURCE INTERPRETATIONS"])
    for item, text in interpretations.items():
        match = source_matches[item]
        report_lines.extend([
            "",
            item,
            f"Source page: {match['page'] if match['found'] else 'Not located'}",
            text,
        ])

    report_lines.extend([
        "",
        "IMPORTANT",
        "Proposal Navigator is a decision-support prototype. Final decisions require sponsor guidance, institutional policy, and Research Administration review.",
    ])

    report_text = "\n".join(report_lines)

    return {
        "confirmed": confirmed,
        "clarifications": clarifications,
        "findings": findings,
        "source_matches": source_matches,
        "interpretations": interpretations,
        "human_review": human_review,
        "answered_count": answered_count,
        "source_count": source_count,
        "readiness_pct": readiness_pct,
        "report_text": report_text,
    }


def render_review_animation(running=True):
    if running:
        return """
        <div class="pn-review-wrap">
            <div class="pn-loader">
                <div class="orbit-ring"></div>
                <div class="compass">🧭</div>
                <div class="orbiter">
                    <div class="dragonfly">
                        <div class="head"></div>
                        <div class="body"></div>
                        <div class="wing1"></div>
                        <div class="wing2"></div>
                        <div class="wing3"></div>
                        <div class="wing4"></div>
                    </div>
                </div>
            </div>
        </div>
        """
    return '<div class="pn-check-circle">✓</div>'


# -----------------------------
# SIDEBAR
# -----------------------------
with st.sidebar:
    st.markdown('<div class="pn-side-compass">🧭</div>', unsafe_allow_html=True)
    st.markdown('<div class="pn-side-title">Proposal<br>Navigator</div>', unsafe_allow_html=True)
    st.markdown('<div class="pn-side-sub">From funding opportunity to proposal ready.</div>', unsafe_allow_html=True)

    for n, label in [
        ("1", "Upload Funding Opportunity"),
        ("2", "Guided Intake Questions"),
        ("3", "AI Review"),
        ("4", "Readiness Summary"),
        ("5", "Detailed Findings"),
        ("6", "Human Review"),
    ]:
        st.markdown(
            f'<div class="pn-step"><span class="pn-dot">{n}</span><span>{label}</span></div>',
            unsafe_allow_html=True,
        )

    st.markdown(
        '<div class="pn-help"><b>Need help?</b><br><br>'
        'Proposal Navigator supports review — it does not replace Research Administration judgment.'
        '</div>',
        unsafe_allow_html=True,
    )

# -----------------------------
# HEADER
# -----------------------------
st.markdown(
    """
    <div class="pn-header">
        <div class="pn-title">Proposal Navigator</div>
        <div class="pn-tag">From funding opportunity to proposal ready.</div>
        <div class="pn-copy">
            Upload sponsor guidance, answer a guided intake, and compare proposal details
            against source language to surface requirements, uncertainty, and items that need human review.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# EXTRACT DOCUMENT
# -----------------------------
col1, col2, col3 = st.columns([1.12, 0.92, 1.0], gap="small")

with col1:
    with st.container(border=True):
        st.markdown(
            '<div class="pn-cardhead"><span class="pn-num">1</span>'
            '<span class="pn-cardtitle">Upload Funding Opportunity</span></div>',
            unsafe_allow_html=True,
        )

        uploaded_file = st.file_uploader(
            "Upload a NOFO, RFP, or sponsor guidance document",
            type=["pdf", "txt", "docx"],
            key="funding_file",
        )

        document_text, document_pages, document_meta = extract_document(uploaded_file)

        if uploaded_file is not None:
            st.success(f"Uploaded: {uploaded_file.name}")
            if document_meta and document_meta.get("error"):
                st.error(f"Document reading error: {document_meta['error']}")
            elif document_text.strip():
                st.success("Document text extracted successfully.")
                if document_meta:
                    st.caption(f"{document_meta['characters']:,} characters extracted • {document_meta['pages']} readable page(s)")
                with st.expander("Preview extracted source text"):
                    st.text_area(
                        "Preview",
                        document_text[:5000],
                        height=180,
                        label_visibility="collapsed",
                    )
            else:
                st.warning("No readable text could be extracted from the uploaded file.")

        st.markdown(
            '<div class="pn-mini"><b>What happens next?</b><br>'
            'Proposal Navigator retrieves sponsor language and compares it with the guided intake.</div>',
            unsafe_allow_html=True,
        )

with col2:
    with st.container(border=True):
        st.markdown(
            '<div class="pn-cardhead"><span class="pn-num">2</span>'
            '<span class="pn-cardtitle">Guided Intake Questions</span></div>',
            unsafe_allow_html=True,
        )

        st.caption("Answer each item. 'Not sure' counts as an answer and routes that item for clarification.")

        personnel = st.radio("Personnel costs?", ["Yes", "No", "Not sure"], horizontal=True, index=None, key="personnel")
        travel = st.radio("Travel?", ["Yes", "No", "Not sure"], horizontal=True, index=None, key="travel")
        equipment = st.radio("Equipment?", ["Yes", "No", "Not sure"], horizontal=True, index=None, key="equipment")
        participant_incentives = st.radio("Participant incentives?", ["Yes", "No", "Not sure"], horizontal=True, index=None, key="participant_incentives")
        subawards = st.radio("Subawards or contracts?", ["Yes", "No", "Not sure"], horizontal=True, index=None, key="subawards")
        cost_share = st.radio("Cost share or matching?", ["Yes", "No", "Not sure"], horizontal=True, index=None, key="cost_share")
        indirect_costs = st.radio("Indirect costs?", ["Yes", "No", "Not sure"], horizontal=True, index=None, key="indirect_costs")

        answered_count = sum(
            value is not None
            for value in [
                personnel,
                travel,
                equipment,
                participant_incentives,
                subawards,
                cost_share,
                indirect_costs,
            ]
        )

        st.progress(answered_count / 7)
        st.caption(f"{answered_count} of 7 answered")

        st.button(
            "Reset Intake",
            on_click=reset_intake,
            use_container_width=True
        )

with col3:
    with st.container(border=True):
        st.markdown(
            '<div class="pn-cardhead"><span class="pn-num">3</span>'
            '<span class="pn-cardtitle">AI Review</span></div>',
            unsafe_allow_html=True,
        )

        review_visual = st.empty()
        review_status = st.empty()

        if st.session_state.review_phase == "complete":
            review_visual.markdown(render_review_animation(False), unsafe_allow_html=True)
            review_status.markdown(
                '<div class="pn-status">✓ Extracting sponsor requirements</div>'
                '<div class="pn-status">✓ Matching proposal categories</div>'
                '<div class="pn-status">✓ Preserving source pages</div>'
                '<div class="pn-status">✓ Readiness findings generated</div>',
                unsafe_allow_html=True,
            )
        else:
            review_visual.markdown(render_review_animation(True), unsafe_allow_html=True)
            review_status.markdown(
                '<div class="pn-status">◎ Ready to analyze the funding document</div>',
                unsafe_allow_html=True,
            )

        run_review = st.button("✨ Run Readiness Review", use_container_width=True)

        st.markdown(
            '<div class="pn-mini"><b>Human-in-the-loop</b><br>'
            'AI interprets matched source excerpts; final proposal decisions remain with people.</div>',
            unsafe_allow_html=True,
        )

responses = {
    "Personnel costs": personnel,
    "Travel": travel,
    "Equipment": equipment,
    "Participant incentives": participant_incentives,
    "Subawards or contracts": subawards,
    "Cost share or matching": cost_share,
    "Indirect costs": indirect_costs,
}

if run_review:
    st.session_state.review_phase = "running"
    review_visual.markdown(render_review_animation(True), unsafe_allow_html=True)
    review_status.markdown('<div class="pn-status">◎ Extracting sponsor requirements...</div>', unsafe_allow_html=True)
    time.sleep(0.35)

    review_status.markdown(
        '<div class="pn-status">✓ Extracting sponsor requirements</div>'
        '<div class="pn-status">◎ Matching proposal categories...</div>',
        unsafe_allow_html=True,
    )
    time.sleep(0.35)

    review_status.markdown(
        '<div class="pn-status">✓ Extracting sponsor requirements</div>'
        '<div class="pn-status">✓ Matching proposal categories</div>'
        '<div class="pn-status">◎ Preserving source pages...</div>',
        unsafe_allow_html=True,
    )
    time.sleep(0.35)

    review_status.markdown(
        '<div class="pn-status">✓ Extracting sponsor requirements</div>'
        '<div class="pn-status">✓ Matching proposal categories</div>'
        '<div class="pn-status">✓ Preserving source pages</div>'
        '<div class="pn-status">◎ Generating readiness findings...</div>',
        unsafe_allow_html=True,
    )

    results = compute_results(responses, document_pages)
    st.session_state.review_results = results
    st.session_state.review_phase = "complete"

    review_visual.markdown(render_review_animation(False), unsafe_allow_html=True)
    review_status.markdown(
        '<div class="pn-status">✓ Extracting sponsor requirements</div>'
        '<div class="pn-status">✓ Matching proposal categories</div>'
        '<div class="pn-status">✓ Preserving source pages</div>'
        '<div class="pn-status">✓ Readiness findings generated</div>',
        unsafe_allow_html=True,
    )

results = st.session_state.review_results

if results:
    # Summary
    with st.container(border=True):
        st.markdown(
            '<div class="pn-cardhead"><span class="pn-num">4</span>'
            '<span class="pn-cardtitle">Readiness Report Summary</span></div>',
            unsafe_allow_html=True,
        )
        circle_col, m1, m2, m3 = st.columns([1.0, 1.0, 1.0, 1.0], gap="small")

        with circle_col:
            st.markdown(
                f'<div class="pn-readiness" style="--pct:{results["readiness_pct"]};">'
                f'<div class="pn-readiness-inner"><strong>{results["readiness_pct"]}%</strong>'
                f'<span>Overall readiness</span></div></div>',
                unsafe_allow_html=True,
            )

        with m1:
            st.metric("Confirmed", len(results["confirmed"]))
        with m2:
            st.metric("Clarifications", len(results["clarifications"]))
        with m3:
            st.metric("Source Matches", results["source_count"])

    # Bottom three cards
    b1, b2, b3 = st.columns([1.0, 1.55, 1.0], gap="small")

    with b1:
        with st.container(border=True):
            st.markdown(
                '<div class="pn-cardhead"><span class="pn-num">4</span>'
                '<span class="pn-cardtitle">Top Items to Address</span></div>',
                unsafe_allow_html=True,
            )
            if results["clarifications"]:
                for item in results["clarifications"]:
                    st.warning(item)
            else:
                st.success("No intake items are currently marked for clarification.")
            st.markdown(
                '<div class="pn-mini"><b>Overall readiness</b><br>'
                'The v0 combines intake, source retrieval, AI interpretation, and human-review routing.</div>',
                unsafe_allow_html=True,
            )

    with b2:
        with st.container(border=True):
            st.markdown(
                '<div class="pn-cardhead"><span class="pn-num">5</span>'
                '<span class="pn-cardtitle">Detailed Findings</span></div>',
                unsafe_allow_html=True,
            )
            findings_df = pd.DataFrame(results["findings"])
            st.dataframe(findings_df, use_container_width=True, hide_index=True, height=230)

            st.markdown("#### Source Findings")
            for item, answer in responses.items():
                match = results["source_matches"][item]
                label = f"{item} — {answer if answer is not None else 'Unanswered'}"
                if match["found"]:
                    label += f" · p. {match['page']}"
                with st.expander(label):
                    if match["found"]:
                        st.markdown(
                            f"**Source:** Funding opportunity, page {match['page']}  \n"
                            f"**Matched term:** {match['keyword']}"
                        )
                        st.info(match["excerpt"])
                        st.markdown("**AI Interpretation**")
                        st.write(results["interpretations"][item])
                    else:
                        st.warning("No obvious matching language was located using the current source search.")
                        st.write(results["interpretations"][item])

    with b3:
        with st.container(border=True):
            st.markdown(
                '<div class="pn-cardhead"><span class="pn-num">6</span>'
                '<span class="pn-cardtitle">Items Requiring Human Review</span></div>',
                unsafe_allow_html=True,
            )
            if results["human_review"]:
                for item in dict.fromkeys(results["human_review"]):
                    match = results["source_matches"][item]
                    with st.expander(item):
                        st.write(
                            f"This item should be reviewed before final proposal decisions are made."
                        )
                        if match["found"]:
                            st.write(f"Relevant source language was located on page {match['page']}.")
                        st.write("Recommended next step: confirm requirements with Research Administration.")
            else:
                st.success("No items are currently routed for human review.")

            st.download_button(
                "⬇ Export Report",
                data=results["report_text"],
                file_name="proposal_navigator_readiness_report.txt",
                mime="text/plain",
                use_container_width=True,
            )

            st.markdown(
                '<div class="pn-mini"><b>Important</b><br>'
                'Proposal Navigator provides decision support based on the uploaded source and intake responses. '
                'Always follow sponsor guidance and institutional review requirements.</div>',
                unsafe_allow_html=True,
            )

    st.markdown('<div class="pn-footer">Proposal Navigator v0 • Source-grounded AI readiness prototype</div>', unsafe_allow_html=True)
else:
    st.info("Complete the guided intake and click 'Run Readiness Review' to generate the readiness dashboard.")
