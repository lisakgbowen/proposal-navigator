
import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from io import BytesIO
import re
from openai import OpenAI

st.set_page_config(
    page_title="Proposal Navigator",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="expanded"
)

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# -----------------------------
# COMPACT DASHBOARD STYLING
# -----------------------------
st.markdown("""
<style>
:root{
    --pink:#ef3b73;
    --pink2:#d92f63;
    --blush:#fff4f7;
    --cream:#fffafb;
    --ink:#2b1832;
    --muted:#756775;
    --border:#f1d3de;
    --green:#1e9b62;
    --orange:#d77a00;
    --blue:#376bb2;
}
.stApp{
    background:linear-gradient(180deg,#fff9fb 0%,#ffffff 34%);
}
.block-container{
    max-width:1500px;
    padding-top:3.2rem;
    padding-bottom:.8rem;
    padding-left:.85rem;
    padding-right:.85rem;
}
[data-testid="stSidebar"]{
    background:linear-gradient(180deg,#fff1f5,#fffafa);
    border-right:1px solid var(--border);
    min-width:165px!important;
    max-width:165px!important;
}
[data-testid="stSidebar"] .block-container{
    padding:.7rem .6rem;
}
.pn-side-title{
    font-family:Georgia,"Times New Roman",serif;
    font-size:1.48rem;
    font-weight:700;
    line-height:1;
    color:var(--ink);
}
.pn-side-sub{
    font-family:Georgia,"Times New Roman",serif;
    font-style:italic;
    color:var(--pink);
    font-size:.72rem;
    line-height:1.15;
    margin:.25rem 0 .65rem 0;
}
.pn-step{
    display:flex;
    align-items:center;
    gap:.38rem;
    margin:.34rem 0;
    font-size:.72rem;
    font-weight:600;
    color:var(--ink);
}
.pn-dot{
    width:21px;height:21px;border-radius:50%;
    display:inline-flex;align-items:center;justify-content:center;
    background:var(--pink);color:white;font-size:.68rem;font-weight:700;
    flex:0 0 21px;
}
.pn-help{
    margin-top:.8rem;
    padding:.6rem;
    border:1px solid var(--border);
    border-radius:10px;
    background:rgba(255,255,255,.72);
    font-size:.67rem;
    line-height:1.35;
    color:var(--muted);
}
.pn-help b{color:var(--ink);font-size:.78rem;}
.pn-header{
    border:1px solid var(--border);
    background:linear-gradient(135deg,#fff7fa 0%,#ffffff 74%);
    border-radius:16px;
    padding:.65rem .95rem;
    margin-bottom:.5rem;
}
.pn-title{
    font-family:Georgia,"Times New Roman",serif;
    color:var(--ink);
    font-size:2rem;
    font-weight:700;
    margin:0;
    line-height:1.05;
}
.pn-tag{
    font-family:Georgia,"Times New Roman",serif;
    font-style:italic;
    color:var(--pink);
    font-size:.92rem;
    margin-top:.12rem;
}
.pn-copy{
    color:var(--muted);
    font-size:.72rem;
    margin-top:.35rem;
}
.pn-cardhead{
    display:flex;
    align-items:center;
    gap:.42rem;
    margin-bottom:.25rem;
}
.pn-num{
    width:24px;height:24px;border-radius:50%;
    display:inline-flex;align-items:center;justify-content:center;
    background:var(--pink);color:#fff;font-weight:700;font-size:.74rem;
}
.pn-cardtitle{
    color:var(--ink);
    font-size:.9rem;
    font-weight:700;
}
[data-testid="stVerticalBlockBorderWrapper"]{
    border-color:var(--border)!important;
    border-radius:14px!important;
    background:rgba(255,255,255,.94)!important;
    box-shadow:0 3px 12px rgba(82,38,61,.035);
}
[data-testid="stVerticalBlockBorderWrapper"] > div{
    padding:.55rem .65rem!important;
}
[data-testid="stFileUploader"]{
    margin-top:-.1rem;
}
[data-testid="stFileUploader"] section{
    min-height:68px;
    padding:.35rem .55rem;
    border:1.5px dashed #ef8daf;
    background:#fff8fa;
    border-radius:10px;
}
[data-testid="stRadio"]{
    margin-bottom:-.36rem;
}
[data-testid="stRadio"] > label{
    font-size:.7rem!important;
    margin-bottom:-.12rem!important;
}
[data-testid="stRadio"] [role="radiogroup"]{
    gap:.35rem!important;
}
[data-testid="stRadio"] [role="radiogroup"] label{
    font-size:.69rem!important;
}
[data-testid="stMetric"]{
    border:1px solid var(--border);
    background:#fffafb;
    border-radius:10px;
    padding:.45rem .55rem;
    min-height:72px;
}
[data-testid="stMetricLabel"]{font-size:.68rem!important;}
[data-testid="stMetricValue"]{font-size:1.45rem!important;}
[data-testid="stExpander"]{
    border:1px solid var(--border)!important;
    border-radius:9px!important;
    overflow:hidden;
}
[data-testid="stExpander"] summary{
    min-height:31px!important;
    padding:.15rem .35rem!important;
    font-size:.7rem!important;
}
div.stButton>button{
    background:linear-gradient(90deg,var(--pink2),var(--pink));
    color:white;border:0;border-radius:9px;
    min-height:36px;font-size:.76rem;font-weight:700;
}
div.stButton>button:hover{
    color:white;border:0;background:linear-gradient(90deg,#c52758,#e9346b);
}
.stAlert{
    padding:.38rem .55rem!important;
    font-size:.69rem!important;
}
.stCaption{font-size:.64rem!important;}
div[data-testid="stDataFrame"]{font-size:.69rem!important;}
h4{font-size:.8rem!important;margin:.25rem 0!important;}
.pn-mini{
    padding:.48rem .55rem;
    border-radius:9px;
    border:1px solid var(--border);
    background:#fff5f8;
    color:var(--muted);
    font-size:.68rem;
    line-height:1.3;
}
.pn-status{
    padding:.42rem .55rem;
    border-radius:9px;
    background:#fff5f8;
    border:1px solid var(--border);
    margin:.28rem 0;
    font-size:.69rem;
}
.pn-footer{
    margin-top:.42rem;
    padding:.45rem .55rem;
    border:1px solid var(--border);
    background:#fff5f8;
    border-radius:9px;
    color:var(--muted);
    font-size:.66rem;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------
# HELPERS
# -----------------------------
def clean_text(text):
    return re.sub(r"\s+", " ", text).strip()

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
                    "excerpt": normalized_text[start:end]
                }

    return {"found": False, "page": None, "keyword": None, "excerpt": None}

def interpret_source_with_ai(category, intake_answer, match):
    if not match["found"]:
        return (
            "No relevant sponsor language was located using the current "
            "retrieval method. Human review is recommended."
        )

    prompt = f"""
You are assisting with proposal-readiness review in research administration.

Review ONLY the source excerpt provided below.
Do not invent sponsor requirements.
Do not assume that silence in the excerpt means something is allowed or prohibited.

Proposal category: {category}
User intake answer: {intake_answer}
Source page: {match['page']}

SOURCE EXCERPT:
{match['excerpt']}

Provide a concise interpretation using this format:

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
            model="gpt-5.6-luna",
            reasoning={"effort": "low"},
            input=prompt
        )
        return response.output_text
    except Exception as e:
        return f"AI interpretation unavailable: {e}"

# -----------------------------
# SIDEBAR
# -----------------------------
with st.sidebar:
    st.markdown('<div class="pn-side-title">🧭 Proposal<br>Navigator</div>', unsafe_allow_html=True)
    st.markdown('<div class="pn-side-sub">From funding opportunity to proposal ready.</div>', unsafe_allow_html=True)

    for n, label in [
        ("1","Upload Funding Opportunity"),
        ("2","Guided Intake Questions"),
        ("3","AI Review"),
        ("4","Readiness Summary"),
        ("5","Detailed Findings"),
        ("6","Human Review"),
    ]:
        st.markdown(
            f'<div class="pn-step"><span class="pn-dot">{n}</span><span>{label}</span></div>',
            unsafe_allow_html=True
        )

    st.markdown(
        '<div class="pn-help"><b>Need help?</b><br><br>'
        'Proposal Navigator supports review — it does not replace Research Administration judgment.'
        '</div>',
        unsafe_allow_html=True
    )

# -----------------------------
# HEADER
# -----------------------------
st.markdown("""
<div class="pn-header">
    <div class="pn-title">Proposal Navigator</div>
    <div class="pn-tag">From funding opportunity to proposal ready.</div>
    <div class="pn-copy">
        Upload sponsor guidance, answer a guided intake, and compare proposal details
        against source language to surface requirements, uncertainty, and items that need human review.
    </div>
</div>
""", unsafe_allow_html=True)

# -----------------------------
# TOP DASHBOARD ROW: 1 / 2 / 3
# -----------------------------
c1, c2, c3 = st.columns([1.05, 1.3, 1.0], gap="small")

document_text = ""
document_pages = []

with c1:
    with st.container(border=True):
        st.markdown(
            '<div class="pn-cardhead"><span class="pn-num">1</span>'
            '<span class="pn-cardtitle">Upload Funding Opportunity</span></div>',
            unsafe_allow_html=True
        )

        uploaded_file = st.file_uploader(
            "Upload a NOFO, RFP, or sponsor guidance document",
            type=["pdf", "txt", "docx"]
        )

        if uploaded_file is not None:
            st.success(f"Uploaded: {uploaded_file.name}")
            file_type = uploaded_file.name.lower()

            try:
                if file_type.endswith(".pdf"):
                    pdf_reader = PdfReader(uploaded_file)
                    for index, page in enumerate(pdf_reader.pages):
                        page_text = page.extract_text()
                        if page_text:
                            document_text += page_text + "\n"
                            document_pages.append((index + 1, page_text))

                elif file_type.endswith(".txt"):
                    document_text = uploaded_file.read().decode("utf-8", errors="ignore")
                    document_pages.append((1, document_text))

                elif file_type.endswith(".docx"):
                    doc = Document(BytesIO(uploaded_file.read()))
                    for paragraph in doc.paragraphs:
                        document_text += paragraph.text + "\n"
                    document_pages.append((1, document_text))

                if document_text.strip():
                    st.success("Document text extracted.")
                    if file_type.endswith(".pdf"):
                        st.caption(f"{len(document_text):,} characters • {len(document_pages)} readable pages")
                    else:
                        st.caption(f"{len(document_text):,} characters extracted")

                    with st.expander("Preview source text"):
                        st.text_area(
                            "Preview",
                            document_text[:5000],
                            height=180,
                            label_visibility="collapsed"
                        )
                else:
                    st.warning("No readable text could be extracted.")
            except Exception as e:
                st.error(f"Document reading error: {e}")

        st.markdown(
            '<div class="pn-mini"><b>What happens next?</b><br>'
            'Sponsor language is retrieved and compared with the guided intake.</div>',
            unsafe_allow_html=True
        )

with c2:
    with st.container(border=True):
        st.markdown(
            '<div class="pn-cardhead"><span class="pn-num">2</span>'
            '<span class="pn-cardtitle">Guided Intake Questions</span></div>',
            unsafe_allow_html=True
        )

        personnel = st.radio("Personnel costs?", ["Yes","No","Not sure"], horizontal=True, key="personnel")
        travel = st.radio("Travel?", ["Yes","No","Not sure"], horizontal=True, key="travel")
        equipment = st.radio("Equipment?", ["Yes","No","Not sure"], horizontal=True, key="equipment")
        participant_incentives = st.radio("Participant incentives?", ["Yes","No","Not sure"], horizontal=True, key="participant_incentives")
        subawards = st.radio("Subawards or contracts?", ["Yes","No","Not sure"], horizontal=True, key="subawards")
        cost_share = st.radio("Cost share or matching?", ["Yes","No","Not sure"], horizontal=True, key="cost_share")
        indirect_costs = st.radio("Indirect costs?", ["Yes","No","Not sure"], horizontal=True, key="indirect_costs")

with c3:
    with st.container(border=True):
        st.markdown(
            '<div class="pn-cardhead"><span class="pn-num">3</span>'
            '<span class="pn-cardtitle">AI Review</span></div>',
            unsafe_allow_html=True
        )

        st.markdown('<div class="pn-status">✓ Extracting sponsor requirements</div>', unsafe_allow_html=True)
        st.markdown('<div class="pn-status">✓ Matching proposal categories</div>', unsafe_allow_html=True)
        st.markdown('<div class="pn-status">✓ Preserving source pages</div>', unsafe_allow_html=True)
        st.markdown('<div class="pn-status">◎ Generating readiness findings</div>', unsafe_allow_html=True)

        run_review = st.button(
            "✨ Run Readiness Review",
            use_container_width=True
        )

        st.markdown(
            '<div class="pn-mini"><b>Human-in-the-loop</b><br>'
            'AI interprets matched source excerpts; final proposal decisions remain with people.</div>',
            unsafe_allow_html=True
        )

# -----------------------------
# REVIEW LOGIC
# -----------------------------
if run_review:
    responses = {
        "Personnel costs": personnel,
        "Travel": travel,
        "Equipment": equipment,
        "Participant incentives": participant_incentives,
        "Subawards or contracts": subawards,
        "Cost share or matching": cost_share,
        "Indirect costs": indirect_costs,
    }

    keyword_map = {
        "Personnel costs": ["salary","salaries","personnel","compensation","fringe"],
        "Travel": ["travel","mileage","airfare","lodging"],
        "Equipment": ["equipment","capital equipment"],
        "Participant incentives": ["participant incentive","incentive","gift card","participant support"],
        "Subawards or contracts": ["subaward","subrecipient","subcontract","contract"],
        "Cost share or matching": ["cost share","cost sharing","matching","match requirement"],
        "Indirect costs": ["indirect costs","indirect cost","indirect","f&a","facilities and administrative"],
    }

    confirmed, missing, findings, source_matches, interpretations = [], [], [], {}, {}

    with st.spinner("Analyzing funding opportunity and generating readiness findings..."):
        for item, answer in responses.items():
            match = find_source_match(document_pages, keyword_map[item])
            source_matches[item] = match

            if answer == "Yes":
                confirmed.append(item)
                findings.append({
                    "Requirement": item,
                    "Status": "Confirmed",
                    "Budget Impact": "Potential impact",
                    "Source": f"p. {match['page']}" if match["found"] else "No source",
                    "Follow-Up": "Review source" if match["found"] else "Human review"
                })
            elif answer == "No":
                findings.append({
                    "Requirement": item,
                    "Status": "Not included",
                    "Budget Impact": "No current impact",
                    "Source": f"p. {match['page']}" if match["found"] else "No source",
                    "Follow-Up": "Review if applicable" if match["found"] else "None"
                })
            else:
                missing.append(item)
                findings.append({
                    "Requirement": item,
                    "Status": "Needs clarification",
                    "Budget Impact": "Unknown",
                    "Source": f"p. {match['page']}" if match["found"] else "No source",
                    "Follow-Up": "Human review"
                })

            interpretations[item] = (
                interpret_source_with_ai(item, answer, match)
                if match["found"]
                else "No source language was located with the current retrieval method. Human review is recommended."
            )

    # -----------------------------
    # SUMMARY ROW
    # -----------------------------
    with st.container(border=True):
        st.markdown(
            '<div class="pn-cardhead"><span class="pn-num">4</span>'
            '<span class="pn-cardtitle">Readiness Report Summary</span></div>',
            unsafe_allow_html=True
        )

        m1, m2, m3, m4 = st.columns(4, gap="small")
        with m1: st.metric("Confirmed", len(confirmed))
        with m2: st.metric("Clarifications", len(missing))
        source_count = sum(1 for m in source_matches.values() if m["found"])
        with m3: st.metric("Source Matches", source_count)
        with m4: st.metric("Readiness", "Good" if len(missing)==0 else "Review Needed")

    # -----------------------------
    # BOTTOM DASHBOARD ROW: 4 / 5 / 6
    # -----------------------------
    b1, b2, b3 = st.columns([1.0, 1.55, 1.0], gap="small")

    with b1:
        with st.container(border=True):
            st.markdown(
                '<div class="pn-cardhead"><span class="pn-num">4</span>'
                '<span class="pn-cardtitle">Top Items to Address</span></div>',
                unsafe_allow_html=True
            )

            if missing:
                for item in missing:
                    st.warning(item)
            else:
                st.success("No intake items are marked uncertain.")

            st.markdown(
                '<div class="pn-mini"><b>Overall readiness</b><br>'
                'The v0 combines intake, source retrieval, and AI interpretation.</div>',
                unsafe_allow_html=True
            )

    with b2:
        with st.container(border=True):
            st.markdown(
                '<div class="pn-cardhead"><span class="pn-num">5</span>'
                '<span class="pn-cardtitle">Detailed Findings</span></div>',
                unsafe_allow_html=True
            )

            findings_df = pd.DataFrame(findings)
            st.dataframe(findings_df, use_container_width=True, hide_index=True, height=230)

            st.markdown("#### Source Findings")

            for item, answer in responses.items():
                match = source_matches[item]
                label = f"{item} — {answer}"
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
                        st.write(interpretations[item])
                    else:
                        st.warning("No obvious matching language was located.")
                        st.write(interpretations[item])

    with b3:
        with st.container(border=True):
            st.markdown(
                '<div class="pn-cardhead"><span class="pn-num">6</span>'
                '<span class="pn-cardtitle">Items Requiring Human Review</span></div>',
                unsafe_allow_html=True
            )

            if missing:
                for item in missing:
                    match = source_matches[item]
                    with st.expander(f"⚠️ {item}"):
                        st.write(
                            f"Additional information is needed before {item.lower()} "
                            "can be evaluated against sponsor requirements."
                        )
                        if match["found"]:
                            st.write(f"Relevant source language was located on page {match['page']}.")
                        st.write(
                            "Recommended next step: review the funding opportunity and "
                            "confirm the requirement with Research Administration."
                        )
            else:
                st.success("No intake items are currently marked uncertain.")

            st.markdown(
                '<div class="pn-mini"><b>Important</b><br>'
                'Proposal Navigator supports decision-making. It does not replace sponsor guidance, '
                'institutional policy, or Research Administration review.</div>',
                unsafe_allow_html=True
            )

    st.markdown(
        '<div class="pn-footer">Proposal Navigator v0 • Source-grounded AI decision-support prototype</div>',
        unsafe_allow_html=True
    )
