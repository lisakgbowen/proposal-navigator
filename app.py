import streamlit as st
import pandas as pd
from pypdf import PdfReader
from docx import Document
from io import BytesIO

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

    confirmed = []
    missing = []
    findings = []

    for item, answer in responses.items():

        if answer == "Yes":
            confirmed.append(item)

            findings.append({
                "Requirement": item,
                "Status": "Confirmed",
                "Budget Impact": "Potential impact",
                "Source": "User intake",
                "Follow-Up": "Review sponsor guidance"
            })

        elif answer == "No":

            findings.append({
                "Requirement": item,
                "Status": "Not included",
                "Budget Impact": "No current impact",
                "Source": "User intake",
                "Follow-Up": "None"
            })

        elif answer == "Not sure":
            missing.append(item)

            findings.append({
                "Requirement": item,
                "Status": "Needs clarification",
                "Budget Impact": "Unknown",
                "Source": "User intake",
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

    st.subheader("Items Requiring Human Review")

    if missing:
        for item in missing:

            with st.expander(item):

                st.write(
                    f"Additional information is needed before {item.lower()} "
                    "can be evaluated against sponsor requirements."
                )

                st.write(
                    "Recommended next step: review the funding opportunity "
                    "and confirm the requirement with Research Administration."
                )

    else:
        st.success("No intake items currently require clarification.")

    if document_text:
        st.success(
            "The uploaded funding opportunity has been read and is available "
            "for the AI review step."
        )

    st.info(
        "This v0 now demonstrates document upload, text extraction, "
        "structured intake, readiness logic, detailed findings, "
        "and human-review routing."
    )
