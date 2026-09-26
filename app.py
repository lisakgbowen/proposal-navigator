import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from io import BytesIO
import re
import time
import hashlib
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
    min-width:190px !important;
    max-width:190px !important;
    background:linear-gradient(180deg,#fff2f6,#fffafa);
    border-right:1px solid var(--border);
}
[data-testid="stSidebar"] .block-container{
    padding:.8rem .65rem;
}

.pn-side-compass{
    font-size:6.25rem;
    line-height:.88;
    margin-bottom:.18rem;
    text-align:center;
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
    padding:.48rem .62rem !important;
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
[data-testid="stRadio"]{
    margin-top:-.28rem;
    margin-bottom:-.62rem;
}
[data-testid="stRadio"] > label{
    font-size:.7rem !important;
    margin-bottom:-.26rem !important;
    line-height:1.05 !important;
}
[data-testid="stRadio"] [role="radiogroup"]{
    gap:.16rem !important;
    justify-content:flex-end;
    flex-wrap:nowrap !important;
    white-space:nowrap !important;
    margin-top:-.12rem;
}
[data-testid="stRadio"] [role="radiogroup"] label{
    font-size:.66rem !important;
    border:1px solid #efb8cb;
    border-radius:7px;
    padding:.05rem .25rem !important;
    background:#fffafb;
    min-width:auto !important;
    width:auto !important;
    justify-content:center;
    white-space:nowrap !important;
}
[data-testid="stRadio"] [role="radiogroup"] label:has(input:checked){
    background:#ef3b73;
    color:white !important;
    border-color:#ef3b73;
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
.stCaption{
    font-size:.64rem !important;
    margin-top:-.22rem !important;
    margin-bottom:-.08rem !important;
}
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
.pn-loader .dragonfly-svg{
    position:absolute;
    top:-4px;
    left:50%;
    transform:translateX(-50%);
    width:40px;
    height:40px;
    display:flex;
    align-items:center;
    justify-content:center;
    filter:drop-shadow(0 1px 2px rgba(0,0,0,.08));
}
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

[data-testid="stProgress"]{
    margin-top:.18rem;
    margin-bottom:.05rem;
}
[data-testid="stProgress"] > div > div{
    height:.42rem !important;
}


.pn-qtext{
    font-size:.69rem;
    font-weight:600;
    color:var(--ink);
    line-height:1.05;
    padding-top:.18rem;
    white-space:nowrap;
}
[data-testid="stHorizontalBlock"]{
    gap:.32rem !important;
}


[data-testid="stRadio"] [role="radiogroup"] > label > div:first-child{
    margin-right:.12rem !important;
}


.pn-info-strip{
    margin-top:.7rem;
    border:1px solid var(--border);
    border-radius:14px;
    background:linear-gradient(90deg,#fff3f7 0%,#fffafa 16%,#ffffff 100%);
    display:grid;
    grid-template-columns:1.1fr 1.35fr 1.55fr 1.6fr 1.55fr;
    gap:0;
    overflow:hidden;
}
.pn-info-brand,
.pn-info-cell{
    padding:.72rem .78rem;
    min-width:0;
}
.pn-info-cell{
    border-left:1px solid #f3dde5;
}
.pn-info-brand{
    display:flex;
    align-items:center;
    gap:.55rem;
}
.pn-info-dragonfly{
    font-size:1.75rem;
    line-height:1;
    transform:rotate(-12deg);
}
.pn-info-tagline{
    font-size:.66rem;
    color:var(--muted);
    line-height:1.32;
}
.pn-info-title{
    font-size:.69rem;
    font-weight:700;
    color:var(--ink);
    margin-bottom:.26rem;
}
.pn-info-text{
    font-size:.64rem;
    color:var(--muted);
    line-height:1.38;
}
.pn-info-list{
    margin:0;
    padding-left:.95rem;
    font-size:.64rem;
    color:var(--muted);
    line-height:1.38;
}
.pn-info-icon{
    font-size:1.08rem;
    margin-right:.25rem;
}
@media (max-width: 900px){
    .pn-info-strip{
        grid-template-columns:1fr 1fr;
    }
    .pn-info-cell{
        border-left:0;
        border-top:1px solid #f3dde5;
    }
}

</style>
""", unsafe_allow_html=True)

# -----------------------------
# HELPERS
# -----------------------------
def reset_intake():
    for key in [
        "personnel",
        "travel",
        "equipment",
        "participant_incentives",
        "gift_cards",
        "research_project",
        "subawards",
        "cost_share",
        "indirect_costs",
    ]:
        st.session_state[key] = None

    st.session_state.review_phase = "ready"
    st.session_state.review_results = None


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


@st.cache_data(show_spinner=False)
def extract_document_cached(file_name, file_bytes):
    """
    Parse the uploaded funding document once and cache the result.
    Streamlit reruns the script whenever an intake answer changes; caching
    prevents the PDF/DOCX from being re-read on every click.
    """
    document_text = ""
    document_pages = []

    if not file_bytes:
        return document_text, document_pages, None

    try:
        lower_name = file_name.lower()

        if lower_name.endswith(".pdf"):
            reader = PdfReader(BytesIO(file_bytes))
            for i, page in enumerate(reader.pages, start=1):
                page_text = page.extract_text() or ""
                if page_text.strip():
                    document_text += page_text + "\n"
                    document_pages.append((i, page_text))

        elif lower_name.endswith(".txt"):
            document_text = file_bytes.decode("utf-8", errors="ignore")
            document_pages.append((1, document_text))

        elif lower_name.endswith(".docx"):
            doc = Document(BytesIO(file_bytes))
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


def get_document_data(uploaded_file):
    """
    Converts the upload to immutable bytes and uses the cached parser above.
    """
    if uploaded_file is None:
        return "", [], None

    file_bytes = uploaded_file.getvalue()
    return extract_document_cached(uploaded_file.name, file_bytes)


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
            input=prompt,
        )
        return response.output_text
    except Exception as e:
        return f"AI interpretation unavailable: {e}"


def compute_results(responses, document_pages, gift_cards=None, research_project=None):
    """
    v0 decision-support logic:
    - Sponsor rules come from the uploaded NOFO/source.
    - ETSU/Research Administration routing is handled separately and explicitly.
    - Required human review is not treated as a failure; unresolved information is.
    """

    keyword_map = {
        "Personnel costs": [
            "salary", "salaries", "personnel", "compensation", "fringe",
            "personnel costs", "salary limitation"
        ],
        "Travel": [
            "travel", "mileage", "airfare", "lodging", "transportation",
            "per diem"
        ],
        "Equipment": [
            "equipment", "capital equipment", "equipment costs"
        ],
        "Participant incentives": [
            "participant incentive", "incentive", "gift card",
            "participant support", "participant costs"
        ],
        "Subawards or contracts": [
            "subaward", "subrecipient", "subcontract", "contract",
            "consultant"
        ],
        "Cost share or matching": [
            "cost share", "cost sharing", "matching", "match requirement"
        ],
        "Indirect costs": [
            "indirect costs", "indirect cost", "f&a",
            "facilities and administrative", "modified total direct cost"
        ],
    }

    findings = []
    source_matches = {}
    interpretations = {}
    confirmed = []
    clarifications = []
    human_review = []

    def add_review(item, reason):
        human_review.append({"item": item, "reason": reason})

    def add_finding(item, answer, match, budget_impact, institutional_guidance, follow_up):
        if answer == "Yes":
            status = "Included"
            confirmed.append(item)
        elif answer == "No":
            status = "Not included"
        elif answer == "Not sure":
            status = "Needs clarification"
            clarifications.append(item)
        else:
            status = "Unanswered"
            clarifications.append(item)

        findings.append({
            "Requirement": item,
            "Status": status,
            "Budget Impact": budget_impact,
            "Sponsor Source": f"p. {match['page']}" if match["found"] else "Not located",
            "Institutional Guidance": institutional_guidance,
            "Follow-Up": follow_up,
        })

    for item, answer in responses.items():
        match = find_source_match(document_pages, keyword_map[item])
        source_matches[item] = match

        # AI interprets sponsor language only. Institutional rules are deterministic below.
        interpretations[item] = interpret_source_with_ai(item, answer, match)

        if answer is None:
            add_review(item, "The intake question has not been answered.")
            add_finding(
                item, answer, match,
                "Unknown",
                "Complete the intake before final review.",
                "Complete intake"
            )
            continue

        if answer == "Not sure":
            add_review(item, "The PI marked this item Not sure and additional clarification is needed.")

        # -------------------------
        # PERSONNEL
        # -------------------------
        if item == "Personnel costs":
            if answer == "Yes":
                add_review(
                    item,
                    "Research Administration must review personnel costs to determine salary, effort, and fringe."
                )
                if not match["found"]:
                    add_review(
                        item,
                        "Personnel costs are planned, but applicable salary/personnel language was not located in the NOFO."
                    )
                add_finding(
                    item, answer, match,
                    "Salary, effort, and fringe may affect the budget.",
                    "Research Administration must verify salary, effort, and fringe. Personnel costs should also be supported by the NOFO.",
                    "Confirm sponsor personnel rules and obtain Research Administration review."
                )
            elif answer == "No":
                add_finding(
                    item, answer, match,
                    "No current personnel budget impact.",
                    "No personnel review is triggered unless personnel are later added.",
                    "None"
                )
            else:
                add_finding(
                    item, answer, match,
                    "Unknown",
                    "Clarify whether personnel costs will be included.",
                    "Human review"
                )

        # -------------------------
        # TRAVEL
        # -------------------------
        elif item == "Travel":
            if answer == "Yes":
                if not match["found"]:
                    add_review(
                        item,
                        "Travel is planned, but travel allowability or limits were not located in the NOFO."
                    )
                add_finding(
                    item, answer, match,
                    "Travel may affect direct costs.",
                    "The NOFO should allow travel and identify applicable allowable travel costs, restrictions, or limits.",
                    "Verify sponsor travel allowability and restrictions."
                )
            elif answer == "No":
                add_finding(
                    item, answer, match,
                    "No current travel budget impact.",
                    "No travel review is triggered unless travel is later added.",
                    "None"
                )
            else:
                add_finding(
                    item, answer, match,
                    "Unknown",
                    "Clarify whether travel will be included and review sponsor travel rules.",
                    "Human review"
                )

        # -------------------------
        # EQUIPMENT
        # -------------------------
        elif item == "Equipment":
            if answer == "Yes":
                if not match["found"]:
                    add_review(
                        item,
                        "Equipment is planned, but equipment allowability was not located in the NOFO."
                    )
                add_finding(
                    item, answer, match,
                    "Equipment may increase direct costs but is excluded from the indirect-cost base.",
                    "Equipment is not subject to indirect costs. Sponsor allowability and any equipment definitions or limits should be verified.",
                    "Verify sponsor equipment rules and exclude equipment from the indirect-cost base."
                )
            elif answer == "No":
                add_finding(
                    item, answer, match,
                    "No current equipment budget impact.",
                    "No equipment-specific indirect-cost adjustment is needed.",
                    "None"
                )
            else:
                add_finding(
                    item, answer, match,
                    "Unknown",
                    "Clarify whether equipment will be included.",
                    "Human review"
                )

        # -------------------------
        # PARTICIPANT INCENTIVES
        # -------------------------
        elif item == "Participant incentives":
            if answer == "Yes":
                if not match["found"]:
                    add_review(
                        item,
                        "Participant incentives are planned, but applicable participant/incentive language was not located in the NOFO."
                    )

                gift_note = "No gift-card response provided."
                if gift_cards == "Yes":
                    add_review(
                        "Gift cards",
                        "Research Administration handles gift cards. Sponsor allowability alone is not sufficient; ETSU policy must also be satisfied."
                    )
                    if research_project == "No":
                        add_review(
                            "Gift cards",
                            "The project is identified as not research. Under the stated ETSU rule, gift cards are not allowable even if the sponsor permits them."
                        )
                        gift_note = "Gift cards selected; project marked not research. Gift cards are not allowable under the stated ETSU rule."
                    elif research_project == "Not sure" or research_project is None:
                        add_review(
                            "Gift cards",
                            "Research status must be confirmed before gift-card allowability can be determined."
                        )
                        gift_note = "Gift cards selected; research status still requires confirmation."
                    else:
                        gift_note = "Gift cards selected; Research Administration must verify ETSU gift-card requirements for the research project."
                elif gift_cards == "No":
                    gift_note = "No gift cards planned."
                elif gift_cards == "Not sure":
                    add_review(
                        "Gift cards",
                        "The PI is unsure whether gift cards will be used. This must be resolved before budget finalization."
                    )
                    gift_note = "Gift-card use is unresolved."

                add_finding(
                    item, answer, match,
                    "Participant costs/incentives may affect direct costs and are excluded from the indirect-cost base.",
                    f"Participant costs are not subject to indirect costs. {gift_note}",
                    "Verify sponsor incentive allowability and resolve any gift-card requirements."
                )
            elif answer == "No":
                add_finding(
                    item, answer, match,
                    "No current participant incentive budget impact.",
                    "No participant incentive or gift-card review is triggered unless incentives are later added.",
                    "None"
                )
            else:
                add_finding(
                    item, answer, match,
                    "Unknown",
                    "Clarify whether participant incentives will be included.",
                    "Human review"
                )

        # -------------------------
        # SUBAWARDS / CONTRACTS
        # -------------------------
        elif item == "Subawards or contracts":
            if answer == "Yes":
                add_review(
                    item,
                    "Research Administration must review proposed contracts/subawards to determine eligibility, appropriate mechanism, and required documentation."
                )
                add_finding(
                    item, answer, match,
                    "May affect direct costs, indirect-cost treatment, and required documentation.",
                    "Contracts and subawards require Research Administration review before submission.",
                    "Send proposed subaward/contract information to Research Administration."
                )
            elif answer == "No":
                add_finding(
                    item, answer, match,
                    "No current subaward/contract budget impact.",
                    "No subaward/contract review is triggered unless one is later added.",
                    "None"
                )
            else:
                add_finding(
                    item, answer, match,
                    "Unknown",
                    "Clarify whether an outside entity will be included and what relationship is proposed.",
                    "Human review"
                )

        # -------------------------
        # COST SHARE
        # -------------------------
        elif item == "Cost share or matching":
            if answer == "Yes":
                add_review(
                    item,
                    "Research Administration must review cost share/matching requirements, commitments, allowability, and approvals."
                )
                add_finding(
                    item, answer, match,
                    "May create institutional financial or effort commitments.",
                    "Cost share/matching must be reviewed by Research Administration before submission.",
                    "Confirm sponsor requirement and obtain Research Administration review/approval."
                )
            elif answer == "No":
                add_finding(
                    item, answer, match,
                    "No current cost-share budget impact.",
                    "No cost-share review is triggered unless a commitment is later added.",
                    "None"
                )
            else:
                add_finding(
                    item, answer, match,
                    "Unknown",
                    "Clarify whether cost share or matching is required or proposed.",
                    "Human review"
                )

        # -------------------------
        # INDIRECT COSTS
        # -------------------------
        elif item == "Indirect costs":
            if answer == "Yes":
                add_review(
                    item,
                    "Research Administration must compare any NOFO F&A limitation with the university's negotiated indirect cost rate and apply the appropriate rate/base."
                )
                if not match["found"]:
                    add_review(
                        item,
                        "Indirect costs are planned, but a sponsor F&A/indirect-cost rate or limitation was not located in the NOFO."
                    )
                add_finding(
                    item, answer, match,
                    "F&A/indirect costs affect the total budget and applicable base.",
                    "Determine any NOFO limit and compare it with the university's negotiated indirect cost rate. Apply exclusions such as equipment and participant costs where applicable.",
                    "Confirm sponsor limit, negotiated rate, and applicable indirect-cost base with Research Administration."
                )
            elif answer == "No":
                add_finding(
                    item, answer, match,
                    "No indirect costs are currently proposed.",
                    "Verify that omitting indirect costs is permitted or appropriate for the opportunity.",
                    "Review if sponsor or university requirements indicate F&A should be included."
                )
            else:
                add_finding(
                    item, answer, match,
                    "Unknown",
                    "Clarify whether indirect costs will be included and what rate/base applies.",
                    "Human review"
                )

    # -------------------------
    # READINESS SCORE
    # -------------------------
    base_answers = list(responses.values())
    required_followups = []
    if responses.get("Participant incentives") == "Yes":
        required_followups.append(gift_cards)
        if gift_cards == "Yes":
            required_followups.append(research_project)

    completion_values = base_answers + required_followups
    completion_total = len(completion_values)
    completion_answered = sum(v is not None for v in completion_values)
    completion_score = completion_answered / completion_total if completion_total else 1.0

    relevant_items = [
        item for item, answer in responses.items()
        if answer in ("Yes", "Not sure")
    ]
    if relevant_items:
        sponsor_match_count = sum(
            1 for item in relevant_items
            if source_matches[item]["found"]
        )
        source_score = sponsor_match_count / len(relevant_items)
    else:
        sponsor_match_count = 0
        source_score = 1.0

    unresolved_count = sum(
        1 for v in completion_values
        if v in (None, "Not sure")
    )
    clarity_score = 1 - (unresolved_count / completion_total if completion_total else 0)

    # Mandatory RA review does NOT lower readiness by itself.
    readiness_pct = round(
        (completion_score * 50) +
        (source_score * 30) +
        (clarity_score * 20)
    )
    readiness_pct = max(0, min(100, readiness_pct))

    # Deduplicate human review items while preserving separate reasons.
    deduped_review = []
    seen = set()
    for entry in human_review:
        key = (entry["item"], entry["reason"])
        if key not in seen:
            seen.add(key)
            deduped_review.append(entry)

    report_lines = [
        "PROPOSAL NAVIGATOR — READINESS REPORT",
        "",
        f"Overall readiness: {readiness_pct}%",
        f"Intake completion: {completion_answered} of {completion_total}",
        f"Relevant sponsor source matches: {sponsor_match_count} of {len(relevant_items)}",
        f"Items requiring Research Administration / human review: {len(deduped_review)}",
        "",
        "NOTE ON READINESS",
        "Required Research Administration review is a normal workflow step and does not automatically reduce the readiness score. Unanswered or unresolved items and missing sponsor language reduce readiness.",
        "",
        "DETAILED FINDINGS",
    ]

    for row in findings:
        report_lines.extend([
            "",
            f"Requirement: {row['Requirement']}",
            f"Status: {row['Status']}",
            f"Budget Impact: {row['Budget Impact']}",
            f"Sponsor Source: {row['Sponsor Source']}",
            f"Institutional Guidance: {row['Institutional Guidance']}",
            f"Follow-Up: {row['Follow-Up']}",
        ])

    report_lines.extend(["", "HUMAN REVIEW"])
    if deduped_review:
        for entry in deduped_review:
            report_lines.extend([
                "",
                f"{entry['item']}: {entry['reason']}",
            ])
    else:
        report_lines.append("No human-review items were identified.")

    report_lines.extend(["", "SOURCE INTERPRETATIONS"])
    for item, interpretation in interpretations.items():
        match = source_matches[item]
        report_lines.extend([
            "",
            item,
            f"Source page: {match['page'] if match['found'] else 'Not located'}",
            interpretation,
        ])

    report_lines.extend([
        "",
        "IMPORTANT",
        "Proposal Navigator is a decision-support prototype. Sponsor language and institutional requirements must be confirmed before final proposal submission.",
    ])

    return {
        "confirmed": confirmed,
        "clarifications": clarifications,
        "findings": findings,
        "source_matches": source_matches,
        "interpretations": interpretations,
        "human_review": deduped_review,
        "answered_count": completion_answered,
        "completion_total": completion_total,
        "source_count": sponsor_match_count,
        "relevant_source_total": len(relevant_items),
        "readiness_pct": readiness_pct,
        "report_text": "\\n".join(report_lines),
    }


def render_review_animation(running=True):
    if running:
        return """
        <div class="pn-review-wrap">
            <div class="pn-loader">
                <div class="orbit-ring"></div>
                <div class="compass">🧭</div>
                <div class="orbiter">
                    <div class="dragonfly-svg">
                        <svg viewBox="0 0 120 120" width="40" height="40" aria-hidden="true">
                            <ellipse cx="38" cy="38" rx="23" ry="10"
                                fill="rgba(255,255,255,0.95)"
                                stroke="#ef3b73" stroke-width="4"
                                transform="rotate(-22 38 38)"/>
                            <ellipse cx="82" cy="38" rx="23" ry="10"
                                fill="rgba(255,255,255,0.95)"
                                stroke="#ef3b73" stroke-width="4"
                                transform="rotate(22 82 38)"/>
                            <ellipse cx="38" cy="68" rx="20" ry="9"
                                fill="rgba(255,255,255,0.95)"
                                stroke="#ef3b73" stroke-width="4"
                                transform="rotate(18 38 68)"/>
                            <ellipse cx="82" cy="68" rx="20" ry="9"
                                fill="rgba(255,255,255,0.95)"
                                stroke="#ef3b73" stroke-width="4"
                                transform="rotate(-18 82 68)"/>
                            <rect x="55" y="20" width="10" height="62" rx="5" fill="#ef3b73"/>
                            <circle cx="60" cy="15" r="8" fill="#ef3b73"/>
                            <rect x="57" y="80" width="6" height="24" rx="3" fill="#ef3b73"/>
                            <line x1="56" y1="9" x2="49" y2="3" stroke="#ef3b73" stroke-width="3" stroke-linecap="round"/>
                            <line x1="64" y1="9" x2="71" y2="3" stroke="#ef3b73" stroke-width="3" stroke-linecap="round"/>
                        </svg>
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
        """
        <div class="pn-help">
            <b>Need help?</b><br><br>
            Proposal Navigator supports review — it does not replace Research Administration judgment.
            <br><br>
            <b>Contact your Research Support Team for assistance.</b>
            <br><br>
            <a href="mailto:conresearch@etsu.edu"
               style="
                   display:block;
                   text-align:center;
                   background:#ef3b73;
                   color:white;
                   text-decoration:none;
                   padding:8px 10px;
                   border-radius:9px;
                   font-weight:700;
               ">
               Contact Support
            </a>
        </div>
        """,
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
col1, col2, col3 = st.columns([1.05, 1.15, 1.0], gap="small")

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

        document_text, document_pages, document_meta = get_document_data(uploaded_file)

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

        # Progress sits directly beneath the title.
        base_keys = [
            "personnel",
            "travel",
            "equipment",
            "participant_incentives",
            "subawards",
            "cost_share",
            "indirect_costs",
        ]

        progress_keys = list(base_keys)
        if st.session_state.get("participant_incentives") == "Yes":
            progress_keys.append("gift_cards")
            if st.session_state.get("gift_cards") == "Yes":
                progress_keys.append("research_project")

        answered_count = sum(
            st.session_state.get(key) is not None
            for key in progress_keys
        )
        total_questions = len(progress_keys)

        st.progress(answered_count / total_questions if total_questions else 0)
        st.caption(f"{answered_count} of {total_questions} answered")

        st.markdown(
            '<div style="font-size:.64rem;color:#77697a;margin:-.12rem 0 .18rem 0;">'
            'Yes, No, and Not sure all count as responses. '
            'Not sure identifies an item that still needs clarification.'
            '</div>',
            unsafe_allow_html=True,
        )

        def intake_row(label, key):
            q_left, q_right = st.columns([1.0, 1.75], gap="small")
            with q_left:
                st.markdown(f'<div class="pn-qtext">{label}</div>', unsafe_allow_html=True)
            with q_right:
                return st.radio(
                    label,
                    ["Yes", "No", "Not sure"],
                    horizontal=True,
                    index=None,
                    key=key,
                    label_visibility="collapsed",
                )

        personnel = intake_row("Personnel costs?", "personnel")
        travel = intake_row("Travel?", "travel")
        equipment = intake_row("Equipment?", "equipment")
        participant_incentives = intake_row("Participant incentives?", "participant_incentives")

        gift_cards = None
        research_project = None

        if participant_incentives == "Yes":
            gift_cards = intake_row("Will gift cards be used?", "gift_cards")

            if gift_cards == "Yes":
                research_project = intake_row(
                    "Is the project considered research?",
                    "research_project"
                )

        subawards = intake_row("Subawards or contracts?", "subawards")
        cost_share = intake_row("Cost share or matching?", "cost_share")
        indirect_costs = intake_row("Indirect costs?", "indirect_costs")

        st.button(
            "Reset Intake",
            on_click=reset_intake,
            use_container_width=True,
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
            review_visual.markdown(
                '<div class="pn-review-wrap">'
                '<div class="compass" style="position:relative;left:auto;top:auto;transform:none;'
                'width:66px;height:66px;border:2px solid #f0c4d5;border-radius:50%;'
                'background:linear-gradient(180deg,#fff7fa,#fff);display:flex;align-items:center;'
                'justify-content:center;font-size:1.55rem;margin:.1rem auto .3rem auto;">🧭</div>'
                '</div>',
                unsafe_allow_html=True
            )
            review_status.markdown(
                '<div class="pn-status">◎ Ready to analyze the funding document</div>',
                unsafe_allow_html=True,
            )

        run_review = st.button("✨ Run Readiness Review", use_container_width=True)

        st.markdown(
            '<div class="pn-mini"><b>Human-in-the-loop</b><br>'
            'AI interprets sponsor language. Institutional routing rules identify items that require Research Administration review.</div>',
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

    results = compute_results(responses, document_pages, gift_cards, research_project)
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
            st.metric("Included", len(results["confirmed"]))
        with m2:
            st.metric("Clarifications", len(results["clarifications"]))
        with m3:
            st.metric(
                "Sponsor Matches",
                f'{results["source_count"]}/{results["relevant_source_total"]}'
            )

        st.caption(
            "Required Research Administration review is a normal workflow step and does not automatically lower readiness."
        )

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
                for item in dict.fromkeys(results["clarifications"]):
                    st.warning(item)
            else:
                st.success("No intake items are currently marked unresolved.")

            if results["human_review"]:
                st.caption(
                    f'{len(results["human_review"])} Research Administration / human-review action(s) identified.'
                )
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
                grouped_review = {}
                for entry in results["human_review"]:
                    grouped_review.setdefault(entry["item"], []).append(entry["reason"])

                for item, reasons in grouped_review.items():
                    with st.expander(f"⚠️ {item}"):
                        for reason in reasons:
                            st.write(f"• {reason}")

                        match = results["source_matches"].get(item)
                        if match and match["found"]:
                            st.caption(f"Sponsor language located on page {match['page']}.")

                        st.markdown("**Next step:** Research Administration / human review.")
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

else:
    st.info("Complete the guided intake and click 'Run Readiness Review' to generate the readiness dashboard.")


# -----------------------------
# BOTTOM INFORMATION STRIP
# -----------------------------
footer_html = """
<div class="pn-info-strip">
<div class="pn-info-brand">
<div class="pn-info-dragonfly">✧</div>
<div class="pn-info-tagline">
<b style="color:#ef3b73;">Guiding your way.</b><br>
Helping your research take flight.
</div>
</div>

<div class="pn-info-cell">
<div class="pn-info-title"><span class="pn-info-icon">♙</span>Who is this for?</div>
<div class="pn-info-text">
Principal Investigators,<br>
Research Staff, and<br>
Pre-Award Teams
</div>
</div>

<div class="pn-info-cell">
<div class="pn-info-title"><span class="pn-info-icon">🎯</span>What does it do?</div>
<div class="pn-info-text">
Helps you understand budget rules,<br>
identify gaps early, and prepare<br>
stronger proposals.
</div>
</div>

<div class="pn-info-cell">
<div class="pn-info-title"><span class="pn-info-icon">✓</span>How does it help?</div>
<ul class="pn-info-list">
<li>Saves time</li>
<li>Reduces errors</li>
<li>Supports compliance</li>
<li>Flags items needing expert review</li>
</ul>
</div>

<div class="pn-info-cell">
<div class="pn-info-title"><span class="pn-info-icon">💡</span>Important</div>
<div class="pn-info-text">
AI provides guidance based on the NOFO and institutional rules.
Always follow up with your Research Administration team for final approval.
</div>
</div>
</div>
"""

st.markdown(footer_html, unsafe_allow_html=True)

