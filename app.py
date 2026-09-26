
import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from io import BytesIO
import re
from openai import OpenAI

# -----------------------------
# PAGE + API
# -----------------------------

st.set_page_config(
    page_title="Proposal Navigator",
    page_icon="🧭",
    layout="wide",
    initial_sidebar_state="expanded"
)

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# -----------------------------
# BRANDING / CSS
# -----------------------------

st.markdown("""
<style>
    :root {
        --pn-pink: #ef3b73;
        --pn-pink-dark: #d92f63;
        --pn-pink-soft: #fff1f6;
        --pn-blush: #fff8fa;
        --pn-ink: #2d1835;
        --pn-muted: #6f6174;
        --pn-border: #f2d7e1;
        --pn-green: #1e9b62;
        --pn-yellow: #a06a00;
        --pn-blue: #2864b8;
    }

    .stApp {
        background: linear-gradient(180deg, #fffafb 0%, #ffffff 34%);
        color: var(--pn-ink);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #fff3f7 0%, #fffafa 100%);
        border-right: 1px solid var(--pn-border);
        min-width: 175px !important;
        max-width: 175px !important;
    }

    [data-testid="stSidebar"] .block-container {
        padding: 0.8rem 0.7rem 0.7rem 0.7rem;
    }

    .block-container {
        max-width: 1380px;
        padding-top: 0.7rem;
        padding-bottom: 1rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }

    .pn-brand {
        font-family: Georgia, "Times New Roman", serif;
        font-size: 1.55rem;
        font-weight: 700;
        line-height: 1;
        color: var(--pn-ink);
        margin-bottom: 0.1rem;
    }

    .pn-tagline {
        font-family: Georgia, "Times New Roman", serif;
        font-style: italic;
        color: var(--pn-pink);
        font-size: 0.8rem;
        margin-bottom: 0.65rem;
        line-height: 1.15;
    }

    .pn-hero {
        padding: 0.75rem 1rem;
        border: 1px solid var(--pn-border);
        border-radius: 20px;
        background: linear-gradient(135deg, #fff7fa 0%, #ffffff 70%);
        box-shadow: 0 8px 24px rgba(74, 31, 53, 0.06);
        margin-bottom: 0.55rem;
    }

    .pn-hero-title {
        font-family: Georgia, "Times New Roman", serif;
        font-size: 2rem;
        font-weight: 700;
        color: var(--pn-ink);
        margin: 0;
    }

    .pn-hero-sub {
        color: var(--pn-pink);
        font-family: Georgia, "Times New Roman", serif;
        font-style: italic;
        font-size: 0.95rem;
        margin-top: 0.05rem;
    }

    .pn-hero-copy {
        color: var(--pn-muted);
        font-size: 0.8rem;
        margin-top: 0.4rem;
        max-width: 900px;
    }

    .pn-step {
        display: flex;
        align-items: center;
        gap: 0.45rem;
        margin-bottom: 0.45rem;
        font-size: 0.78rem;
        color: var(--pn-ink);
        font-weight: 600;
    }

    .pn-step-num {
        width: 23px;
        height: 23px;
        border-radius: 50%;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        background: var(--pn-pink);
        color: white;
        font-weight: 700;
        flex: 0 0 23px;
    }

    .pn-sidebar-card {
        border: 1px solid var(--pn-border);
        background: rgba(255,255,255,0.72);
        padding: 0.7rem;
        border-radius: 12px;
        margin-top: 0.7rem;
    }

    .pn-card-title {
        font-weight: 700;
        color: var(--pn-ink);
        font-size: 0.88rem;
        margin-bottom: 0.25rem;
    }

    .pn-card-copy {
        color: var(--pn-muted);
        font-size: 0.72rem;
        line-height: 1.25;
    }

    .pn-section {
        border: 1px solid var(--pn-border);
        background: white;
        border-radius: 14px;
        padding: 0.65rem 0.8rem 0.5rem 0.8rem;
        box-shadow: 0 5px 18px rgba(74, 31, 53, 0.04);
        margin-bottom: 0.45rem;
    }

    .pn-section-head {
        display: flex;
        align-items: center;
        gap: 0.45rem;
        margin-bottom: 0.35rem;
    }

    .pn-badge {
        width: 26px;
        height: 26px;
        border-radius: 50%;
        background: var(--pn-pink);
        color: white;
        font-weight: 700;
        display: inline-flex;
        align-items: center;
        justify-content: center;
    }

    .pn-section-title {
        font-size: 0.98rem;
        font-weight: 700;
        color: var(--pn-ink);
    }

    .pn-note {
        border: 1px solid #f5ccd9;
        background: var(--pn-pink-soft);
        color: #684455;
        padding: 0.55rem 0.7rem;
        border-radius: 10px;
        font-size: 0.76rem;
    }

    div.stButton > button {
        background: linear-gradient(90deg, var(--pn-pink-dark), var(--pn-pink));
        color: white;
        border: 0;
        border-radius: 10px;
        padding: 0.65rem 1.15rem;
        font-weight: 700;
    }

    div.stButton > button:hover {
        border: 0;
        color: white;
        background: linear-gradient(90deg, #c72759, #e8366c);
    }

    [data-testid="stMetric"] {
        border: 1px solid var(--pn-border);
        border-radius: 14px;
        padding: 0.8rem 1rem;
        background: #fffafa;
    }

    [data-testid="stExpander"] {
        border: 1px solid var(--pn-border);
        border-radius: 12px;
        overflow: hidden;
    }

    [data-testid="stFileUploader"] section {
        border: 1.5px dashed #ef8daf;
        background: #fff7fa;
        border-radius: 14px;
    }

    hr {
        border-color: var(--pn-border);
    }

    /* Compact Streamlit controls */
    [data-testid="stRadio"] {
        margin-bottom: -0.25rem;
    }

    [data-testid="stRadio"] > label {
        font-size: 0.76rem;
        margin-bottom: 0.05rem;
    }

    [data-testid="stRadio"] [role="radiogroup"] {
        gap: 0.55rem;
    }

    [data-testid="stRadio"] [role="radiogroup"] label {
        font-size: 0.74rem;
    }

    [data-testid="stFileUploader"] {
        margin-bottom: 0.25rem;
    }

    [data-testid="stFileUploader"] section {
        min-height: 64px;
        padding: 0.4rem 0.55rem;
    }

    [data-testid="stMetric"] {
        padding: 0.55rem 0.7rem;
        min-height: 82px;
    }

    [data-testid="stMetricLabel"] {
        font-size: 0.72rem;
    }

    [data-testid="stMetricValue"] {
        font-size: 1.7rem;
    }

    [data-testid="stExpander"] summary {
        min-height: 34px;
        padding-top: 0.25rem;
        padding-bottom: 0.25rem;
        font-size: 0.78rem;
    }

    div[data-testid="stDataFrame"] {
        font-size: 0.78rem;
    }

    .stCaption {
        font-size: 0.7rem;
    }

    div.stButton > button {
        min-height: 38px;
        font-size: 0.82rem;
        padding: 0.45rem 0.8rem;
    }

    .stAlert {
        padding-top: 0.45rem;
        padding-bottom: 0.45rem;
        font-size: 0.78rem;
    }

    h4 {
        margin-top: 0.45rem !important;
        margin-bottom: 0.35rem !important;
    }

</style>
""", unsafe_allow_html=True)

# -----------------------------
# HELPERS
# -----------------------------

def clean_text(text):
    return re.sub(r"\s+", " ", text).strip()


def find_source_match(pages, keywords):
    """
    Search page-by-page using normalized text so PDF line breaks
    and extra spaces do not prevent matches.
    """
    for page_number, page_text in pages:
        normalized_text = clean_text(page_text)
        lower_text = normalized_text.lower()

        for keyword in keywords:
            normalized_keyword = clean_text(keyword).lower()
            position = lower_text.find(normalized_keyword)

            if position != -1:
                start = max(0, position - 150)
                end = min(len(normalized_text), position + 450)
                excerpt = normalized_text[start:end]

                return {
                    "found": True,
                    "page": page_number,
                    "keyword": keyword,
                    "excerpt": excerpt
                }

    return {
        "found": False,
        "page": None,
        "keyword": None,
        "excerpt": None
    }


def interpret_source_with_ai(category, intake_answer, match):
    """
    Uses the retrieved sponsor language to generate a concise,
    source-grounded interpretation.
    """
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
    st.markdown('<div class="pn-brand">🧭 Proposal<br>Navigator</div>', unsafe_allow_html=True)
    st.markdown('<div class="pn-tagline">From funding opportunity to proposal ready.</div>', unsafe_allow_html=True)

    sidebar_steps = [
        ("1", "Upload Funding Opportunity"),
        ("2", "Guided Intake Questions"),
        ("3", "AI Review"),
        ("4", "Readiness Summary"),
        ("5", "Detailed Findings"),
        ("6", "Human Review"),
    ]

    for n, label in sidebar_steps:
        st.markdown(
            f'<div class="pn-step"><span class="pn-step-num">{n}</span><span>{label}</span></div>',
            unsafe_allow_html=True
        )

    st.markdown(
        """
        <div class="pn-sidebar-card">
            <div class="pn-card-title">Need help?</div>
            <div class="pn-card-copy">
                Proposal Navigator supports review — it does not replace Research Administration judgment.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# -----------------------------
# HERO
# -----------------------------

st.markdown(
    """
    <div class="pn-hero">
        <div class="pn-hero-title">Proposal Navigator</div>
        <div class="pn-hero-sub">From funding opportunity to proposal ready.</div>
        <div class="pn-hero-copy">
            Upload sponsor guidance, answer a guided intake, and compare proposal details
            against source language to surface requirements, uncertainty, and items that need human review.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# -----------------------------
# STEP 1 + STEP 2
# -----------------------------

left, right = st.columns([1, 1.12], gap="small")

with left:
    st.markdown(
        """
        <div class="pn-section">
            <div class="pn-section-head">
                <span class="pn-badge">1</span>
                <span class="pn-section-title">Upload Funding Opportunity</span>
            </div>
        """,
        unsafe_allow_html=True
    )

    uploaded_file = st.file_uploader(
        "Upload a NOFO, RFP, or sponsor guidance document",
        type=["pdf", "txt", "docx"],
        label_visibility="visible"
    )

    document_text = ""
    document_pages = []

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
                st.success("Document text extracted successfully.")

                c1, c2 = st.columns(2)
                with c1:
                    st.caption(f"{len(document_text):,} characters extracted")
                with c2:
                    if file_type.endswith(".pdf"):
                        st.caption(f"{len(document_pages)} readable PDF pages")

                with st.expander("Preview extracted source text"):
                    st.text_area(
                        "Document Preview",
                        document_text[:5000],
                        height=260,
                        label_visibility="collapsed"
                    )
            else:
                st.warning("The file uploaded, but no readable text could be extracted.")

        except Exception as e:
            st.error(f"Document reading error: {e}")

    st.markdown(
        """
        <div class="pn-note">
            <strong>What happens next?</strong><br>
            Proposal Navigator retrieves sponsor language and compares it with your intake responses.
        </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with right:
    st.markdown(
        """
        <div class="pn-section">
            <div class="pn-section-head">
                <span class="pn-badge">2</span>
                <span class="pn-section-title">Guided Intake Questions</span>
            </div>
        """,
        unsafe_allow_html=True
    )

    personnel = st.radio(
        "Will the project include personnel costs?",
        ["Yes", "No", "Not sure"],
        horizontal=True,
        key="personnel"
    )

    travel = st.radio(
        "Will the project include travel?",
        ["Yes", "No", "Not sure"],
        horizontal=True,
        key="travel"
    )

    equipment = st.radio(
        "Will the project include equipment?",
        ["Yes", "No", "Not sure"],
        horizontal=True,
        key="equipment"
    )

    participant_incentives = st.radio(
        "Will the project include participant incentives?",
        ["Yes", "No", "Not sure"],
        horizontal=True,
        key="participant_incentives"
    )

    subawards = st.radio(
        "Will the project include subawards or contracts?",
        ["Yes", "No", "Not sure"],
        horizontal=True,
        key="subawards"
    )

    cost_share = st.radio(
        "Will the project include cost share or matching?",
        ["Yes", "No", "Not sure"],
        horizontal=True,
        key="cost_share"
    )

    indirect_costs = st.radio(
        "Will the project include indirect costs?",
        ["Yes", "No", "Not sure"],
        horizontal=True,
        key="indirect_costs"
    )

    st.markdown("</div>", unsafe_allow_html=True)

# -----------------------------
# RUN REVIEW
# -----------------------------

st.markdown("<div style=\"height:0.15rem\"></div>", unsafe_allow_html=True)
button_col1, button_col2, button_col3 = st.columns([1, 1.2, 1])

with button_col2:
    run_review = st.button(
        "✨ Run Proposal Readiness Review",
        use_container_width=True
    )

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
        "Personnel costs": [
            "salary", "salaries", "personnel", "compensation", "fringe"
        ],
        "Travel": [
            "travel", "mileage", "airfare", "lodging"
        ],
        "Equipment": [
            "equipment", "capital equipment"
        ],
        "Participant incentives": [
            "participant incentive", "incentive", "gift card", "participant support"
        ],
        "Subawards or contracts": [
            "subaward", "subrecipient", "subcontract", "contract"
        ],
        "Cost share or matching": [
            "cost share", "cost sharing", "matching", "match requirement"
        ],
        "Indirect costs": [
            "indirect costs", "indirect cost", "indirect", "f&a",
            "facilities and administrative"
        ],
    }

    confirmed = []
    missing = []
    findings = []
    source_matches = {}
    interpretations = {}

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
                    "Source": (
                        f"Funding opportunity p. {match['page']}"
                        if match["found"]
                        else "No source located"
                    ),
                    "Follow-Up": (
                        "Review source language"
                        if match["found"]
                        else "Human review"
                    )
                })

            elif answer == "No":
                findings.append({
                    "Requirement": item,
                    "Status": "Not included",
                    "Budget Impact": "No current impact",
                    "Source": (
                        f"Funding opportunity p. {match['page']}"
                        if match["found"]
                        else "No source located"
                    ),
                    "Follow-Up": (
                        "Review if applicable"
                        if match["found"]
                        else "None"
                    )
                })

            elif answer == "Not sure":
                missing.append(item)
                findings.append({
                    "Requirement": item,
                    "Status": "Needs clarification",
                    "Budget Impact": "Unknown",
                    "Source": (
                        f"Funding opportunity p. {match['page']}"
                        if match["found"]
                        else "No source located"
                    ),
                    "Follow-Up": "Human review"
                })

            if match["found"]:
                interpretations[item] = interpret_source_with_ai(
                    item,
                    answer,
                    match
                )
            else:
                interpretations[item] = (
                    "No source language was located with the current retrieval method. "
                    "Human review is recommended."
                )

    # -----------------------------
    # COMPACT RESULTS DASHBOARD
    # -----------------------------

    st.markdown(
        """
        <div class="pn-section">
            <div class="pn-section-head">
                <span class="pn-badge">4</span>
                <span class="pn-section-title">Readiness Report Summary</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    m1, m2, m3, m4 = st.columns(4, gap="small")
    with m1:
        st.metric("Confirmed", len(confirmed))
    with m2:
        st.metric("Clarifications", len(missing))
    with m3:
        source_count = sum(1 for m in source_matches.values() if m["found"])
        st.metric("Source Matches", source_count)
    with m4:
        status = "Good" if len(missing) == 0 else "Review Needed"
        st.metric("Readiness", status)

    results_left, results_mid, results_right = st.columns([1.05, 1.55, 1.05], gap="small")

    with results_left:
        st.markdown(
            """
            <div class="pn-section">
                <div class="pn-section-head">
                    <span class="pn-badge">4</span>
                    <span class="pn-section-title">Top Items to Address</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        if missing:
            for item in missing:
                st.warning(item)
        else:
            st.success("No intake items are marked uncertain.")

        st.markdown(
            """
            <div class="pn-note">
                <strong>Overall readiness</strong><br>
                The v0 uses intake responses, source matching, and AI interpretation
                to surface proposal-readiness issues.
            </div>
            """,
            unsafe_allow_html=True
        )

    with results_mid:
        st.markdown(
            """
            <div class="pn-section">
                <div class="pn-section-head">
                    <span class="pn-badge">5</span>
                    <span class="pn-section-title">Detailed Findings</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        findings_df = pd.DataFrame(findings)
        st.dataframe(
            findings_df,
            use_container_width=True,
            hide_index=True,
            height=250
        )

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
                    st.caption(
                        "AI interpretation is grounded in the retrieved source excerpt. "
                        "Final proposal decisions still require human review."
                    )
                else:
                    st.warning(
                        "No obvious matching language was located using the current source search."
                    )
                    st.write(interpretations[item])

    with results_right:
        st.markdown(
            """
            <div class="pn-section">
                <div class="pn-section-head">
                    <span class="pn-badge">6</span>
                    <span class="pn-section-title">Human Review</span>
                </div>
            </div>
            """,
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
                        st.write(
                            f"Relevant source language was located on page {match['page']}."
                        )

                    st.write(
                        "Recommended next step: review the funding opportunity and "
                        "confirm the requirement with Research Administration."
                    )
        else:
            st.success("No intake items are currently marked as uncertain.")

        st.markdown(
            """
            <div class="pn-note">
                <strong>Important</strong><br>
                Proposal Navigator supports decision-making. It does not replace sponsor guidance,
                institutional policy, or Research Administration review.
            </div>
            """,
            unsafe_allow_html=True
        )
