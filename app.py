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

def find_source_excerpt(text, keywords, window=300):
    """
    Finds the first occurrence of any keyword and returns
    a short excerpt around it.
    """
    if not text:
        return "No funding opportunity text available."

    lower_text = text.lower()

    for keyword in keywords:
        position = lower_text.find(keyword.lower())

        if position != -1:
            start = max(0, position - window)
            end = min(len(text), position + window)

            excerpt = text[start:end]
            excerpt = re.sub(r"\s+", " ", excerpt).strip()

            return excerpt

    return "No matching language located in the uploaded document."


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

if uploaded_file is not None:
    st.success(f"Uploaded: {uploaded_file.name}")

    file_type = uploaded_file.name.lower()

    try:
        if file_type.endswith(".pdf"):
            pdf_reader = PdfReader(uploaded_file)

            for page in pdf_reader.pages:
                page_text = page.extract_text()

                if page_text:
                    document_text += page_text + "\n"

        elif file_type.endswith(".txt"):
            document_text = uploaded_file.read().decode(
                "utf-8",
                errors="ignore"
            )

        elif file_type.endswith(".docx"):
            doc = Document(BytesIO(uploaded_file.read()))

            for paragraph in doc.paragraphs:
                document_text += paragraph.text + "\n"

        if document_text.strip():
            st.success("Document text extracted successfully.")

            with st.expander("Preview extracted document text"):
                st.text_area(
                    "Document Preview",
                    document_text[:5000],
                    height=300
                )

            st.caption(
                f"Approximately {len(document_text):,} characters extracted."
            )

        else:
            st.warning(
                "The file uploaded successfully, but no readable text "
                "could be extracted."
            )

    except Exception as e:
        st.error(f"Document reading error: {e}")


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
            "indirect cost",
            "indirect costs",
            "f&a",
            "facilities and administrative"
        ],
    }

    confirmed = []
    missing = []
    findings = []

    for item, answer in responses.items():

        excerpt = find_source_excerpt(
            document_text,
            keyword_map[item]
        )

        if answer == "Yes":
            confirmed.append(item)

            status = "Confirmed"

            if "No matching language" in excerpt:
                follow_up = "Human review"
            else:
                follow_up = "Review source language"

            findings.append({
                "Requirement": item,
                "Status": status,
                "Budget Impact": "Potential impact",
                "Source Match": (
                    "Found in document"
                    if "No matching language" not in excerpt
                    else "Not located"
                ),
                "Follow-Up": follow_up
            })

        elif answer == "No":

            findings.append({
                "Requirement": item,
                "Status": "Not included",
                "Budget Impact": "No current impact",
                "Source Match": (
                    "Found in document"
                    if "No matching language" not in excerpt
                    else "Not located"
                ),
                "Follow-Up": "None"
            })

        elif answer == "Not sure":
            missing.append(item)

            findings.append({
                "Requirement": item,
                "Status": "Needs clarification",
                "Budget Impact": "Unknown",
                "Source Match": (
                    "Found in document"
                    if "No matching language" not in excerpt
                    else "Not located"
                ),
                "Follow-Up": "Human review"
            })

    st.header("Step 3: Readiness Report Summary")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Confirmed Areas", len(confirmed))

    with col2:
        st.metric("Needs Clarification", len(missing))

    with col3:
        status = "Good" if len(missing) == 0 else "Review Needed"
        st.metric("Readiness Status", status)

    st.subheader("Detailed Findings")

    findings_df = pd.DataFrame(findings)

    st.dataframe(
        findings_df,
        use_container_width=True,
        hide_index=True
    )


    # -----------------------------
    # SOURCE FINDINGS
    # -----------------------------

    st.subheader("Funding Opportunity Source Findings")

    for item, answer in responses.items():

        excerpt = find_source_excerpt(
            document_text,
            keyword_map[item]
        )

        with st.expander(f"{item} — {answer}"):

            if "No matching language" in excerpt:
                st.warning(
                    "No obvious matching language was located "
                    "using the current keyword search."
                )

            else:
                st.write("Relevant source excerpt:")
                st.info(excerpt)

            st.caption(
                "This source match is based on keyword retrieval only. "
                "It has not yet been interpreted by AI."
            )


    # -----------------------------
    # HUMAN REVIEW
    # -----------------------------

    st.subheader("Items Requiring Human Review")

    if missing:

        for item in missing:

            with st.expander(item):

                st.write(
                    f"Additional information is needed before "
                    f"{item.lower()} can be evaluated against "
                    "sponsor requirements."
                )

                st.write(
                    "Recommended next step: review the funding opportunity "
                    "and confirm the requirement with Research Administration."
                )

    else:
        st.success("No intake items currently require clarification.")

    st.info(
        "This v0 now connects user intake with source language from the "
        "uploaded funding opportunity. AI interpretation will be added next."
    )
