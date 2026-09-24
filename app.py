import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from io import BytesIO
import re

st.set_page_config(
    page_title="Proposal Navigator",
    page_icon="🧭",
    layout="wide"
)

st.title("Proposal Navigator")
st.subheader("From funding opportunity to proposal ready.")

st.write(
    "This prototype helps users review a funding opportunity, "
    "identify proposal requirements, and flag items that may need human review."
)


# -----------------------------
# Helper function
# -----------------------------

def clean_text(text):
    return re.sub(r"\s+", " ", text).strip()


def find_source_match(pages, keywords):
    """
    Search page-by-page and return a page number
    plus a cleaner excerpt around the first keyword match.
    """

    for page_number, page_text in pages:

        lower_text = page_text.lower()

        for keyword in keywords:

            position = lower_text.find(keyword.lower())

            if position != -1:

                start = max(0, position - 150)
                end = min(len(page_text), position + 450)

                excerpt = clean_text(page_text[start:end])

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


# -----------------------------
# STEP 1
# -----------------------------

st.divider()

st.header("Step 1: Upload Funding Opportunity")

uploaded_file = st.file_uploader(
    "Upload a NOFO, RFP, or sponsor guidance document",
    type=["pdf", "txt", "docx"]
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

                    document_pages.append(
                        (index + 1, page_text)
                    )

        elif file_type.endswith(".txt"):

            document_text = uploaded_file.read().decode(
                "utf-8",
                errors="ignore"
            )

            document_pages.append(
                (1, document_text)
            )

        elif file_type.endswith(".docx"):

            doc = Document(
                BytesIO(uploaded_file.read())
            )

            for paragraph in doc.paragraphs:
                document_text += paragraph.text + "\n"

            document_pages.append(
                (1, document_text)
            )

        if document_text.strip():

            st.success("Document text extracted successfully.")

            with st.expander(
                "Preview extracted document text"
            ):

                st.text_area(
                    "Document Preview",
                    document_text[:5000],
                    height=300
                )

            st.caption(
                f"Approximately {len(document_text):,} "
                "characters extracted."
            )

            if file_type.endswith(".pdf"):
                st.caption(
                    f"{len(document_pages)} readable PDF pages detected."
                )

        else:

            st.warning(
                "The file uploaded successfully, but no readable "
                "text could be extracted."
            )

    except Exception as e:

        st.error(
            f"Document reading error: {e}"
        )


# -----------------------------
# STEP 2
# -----------------------------

st.divider()

st.header("Step 2: Guided Intake Questions")

personnel = st.radio(
    "Will the project include personnel costs?",
    ["Yes", "No", "Not sure"],
    horizontal=True
)

travel = st.radio(
    "Will the project include travel?",
    ["Yes", "No", "Not sure"],
    horizontal=True
)

equipment = st.radio(
    "Will the project include equipment?",
    ["Yes", "No", "Not sure"],
    horizontal=True
)

participant_incentives = st.radio(
    "Will the project include participant incentives?",
    ["Yes", "No", "Not sure"],
    horizontal=True
)

subawards = st.radio(
    "Will the project include subawards or contracts?",
    ["Yes", "No", "Not sure"],
    horizontal=True
)

cost_share = st.radio(
    "Will the project include cost share or matching?",
    ["Yes", "No", "Not sure"],
    horizontal=True
)

indirect_costs = st.radio(
    "Will the project include indirect costs?",
    ["Yes", "No", "Not sure"],
    horizontal=True
)


# -----------------------------
# STEP 3
# -----------------------------

st.divider()

if st.button("Run Readiness Review"):

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
            "salary",
            "salaries",
            "personnel",
            "compensation",
            "fringe"
        ],

        "Travel": [
            "travel",
            "mileage",
            "airfare",
            "lodging"
        ],

        "Equipment": [
            "equipment",
            "capital equipment"
        ],

        "Participant incentives": [
            "participant incentive",
            "incentive",
            "gift card",
            "participant support"
        ],

        "Subawards or contracts": [
            "subaward",
            "subrecipient",
            "subcontract",
            "contract"
        ],

        "Cost share or matching": [
            "cost share",
            "cost sharing",
            "matching",
            "match requirement"
        ],

        "Indirect costs": [
            "indirect costs",
            "indirect cost",
            "f&a",
            "facilities and administrative"
        ],
    }

    confirmed = []
    missing = []
    findings = []
    source_matches = {}

    for item, answer in responses.items():

        match = find_source_match(
            document_pages,
            keyword_map[item]
        )

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


    st.header("Step 3: Readiness Report Summary")

    col1, col2, col3 = st.columns(3)

    with col1:

        st.metric(
            "Confirmed Areas",
            len(confirmed)
        )

    with col2:

        st.metric(
            "Needs Clarification",
            len(missing)
        )

    with col3:

        status = (
            "Good"
            if len(missing) == 0
            else "Review Needed"
        )

        st.metric(
            "Readiness Status",
            status
        )


    # -----------------------------
    # Detailed Findings
    # -----------------------------

    st.subheader("Detailed Findings")

    findings_df = pd.DataFrame(findings)

    st.dataframe(
        findings_df,
        use_container_width=True,
        hide_index=True
    )


    # -----------------------------
    # Source Findings
    # -----------------------------

    st.subheader(
        "Funding Opportunity Source Findings"
    )

    for item, answer in responses.items():

        match = source_matches[item]

        with st.expander(
            f"{item} — {answer}"
        ):

            if match["found"]:

                st.write(
                    f"**Source:** Funding opportunity, "
                    f"page {match['page']}"
                )

                st.write(
                    f"**Matched term:** {match['keyword']}"
                )

                st.write(
                    "**Relevant source excerpt:**"
                )

                st.info(
                    match["excerpt"]
                )

            else:

                st.warning(
                    "No obvious matching language was located "
                    "using the current source search."
                )

            st.caption(
                "The source text has been retrieved from the "
                "uploaded funding opportunity but has not yet "
                "been interpreted by generative AI."
            )


    # -----------------------------
    # Human Review
    # -----------------------------

    st.subheader(
        "Items Requiring Human Review"
    )

    if missing:

        for item in missing:

            with st.expander(item):

                st.write(
                    f"Additional information is needed before "
                    f"{item.lower()} can be evaluated against "
                    "sponsor requirements."
                )

                match = source_matches[item]

                if match["found"]:

                    st.write(
                        f"Relevant source language was located "
                        f"on page {match['page']}."
                    )

                st.write(
                    "Recommended next step: review the funding "
                    "opportunity and confirm the requirement "
                    "with Research Administration."
                )

    else:

        st.success(
            "No intake items currently require clarification."
        )


    st.info(
        "This v0 connects user intake with page-level source "
        "language from the uploaded funding opportunity. "
        "Generative AI interpretation will be added next."
    )
