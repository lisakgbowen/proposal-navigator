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

st.caption(
    "For this v0, we are first testing the upload and intake workflow "
    "before connecting the full AI review."
)
