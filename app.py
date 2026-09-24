import streamlit as st

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

if uploaded_file is not None:
    st.success(f"Uploaded: {uploaded_file.name}")

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

    for item, answer in responses.items():
        if answer == "Yes":
            confirmed.append(item)
        elif answer == "Not sure":
            missing.append(item)

    st.header("Step 3: Readiness Review")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Confirmed Areas", len(confirmed))

    with col2:
        st.metric("Needs Clarification", len(missing))

    with col3:
        if len(missing) == 0:
            st.metric("Readiness Status", "Good")
        else:
            st.metric("Readiness Status", "Review Needed")

    if confirmed:
        st.subheader("Confirmed Proposal Areas")
        for item in confirmed:
            st.success(item)

    if missing:
        st.subheader("Items Requiring Clarification")
        for item in missing:
            st.warning(f"{item} — additional information is needed.")

    if not missing:
        st.success(
            "No intake questions were marked as uncertain. "
            "The proposal can proceed to detailed funding opportunity review."
        )

    st.info(
        "This v0 demonstrates structured intake and readiness logic. "
        "Funding opportunity interpretation and source-based AI findings "
        "will be added in the next development step."
    )
