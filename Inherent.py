import streamlit as st
import pandas as pd
from io import BytesIO

# Function to load the uploaded Excel file
def load_data(uploaded_file):
    inherent_data = pd.read_excel(uploaded_file, sheet_name="Inherent Risk Assessment")
    mitigation_data = pd.read_excel(uploaded_file, sheet_name="Mitigation Question Bank")
    return inherent_data, mitigation_data

# Stage 1: Upload the Excel file and display inherent risk questions
st.title("Third Party Risk Management (TPRM) - Inherent Risk Assessment")

# Upload the Excel file
uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")
if uploaded_file is not None:
    # Load the data from the uploaded file
    inherent_data, mitigation_data = load_data(uploaded_file)
    
    # Stage 1: Ask Inherent Risk Questions (Yes/No)
    st.header("Inherent Risk Questions")

    responses = []
    for index, row in inherent_data.iterrows():
        question = row['Question']
        response = st.radio(f"{question}", ("Yes", "No"))
        responses.append(response)

    # Generate Mitigation Questions based on Inherent Risk Responses
    mitigation_questions_to_ask = []
    for idx, response in enumerate(responses):
        if response == "Yes":  # Only ask mitigation questions for "Yes" responses
            inherent_question_id = inherent_data.iloc[idx]["ID"]
            mitigation_questions_for_domain = mitigation_data[mitigation_data['Triggering Question ID'] == inherent_question_id]
            mitigation_questions_to_ask.append(mitigation_questions_for_domain)

    # Flatten the list of questions to display
    mitigation_questions_to_ask = pd.concat(mitigation_questions_to_ask)
    
    # Prepare the Excel file for download with the mitigation questions
    @st.cache_data
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Mitigation Questions")
        processed_data = output.getvalue()
        return processed_data

    # Display a download button for the mitigation questions Excel file
    if mitigation_questions_to_ask.empty:
        st.warning("No mitigation questions required based on your responses.")
    else:
        st.download_button(
            label="Download Mitigation Questions",
            data=to_excel(mitigation_questions_to_ask),
            file_name="Mitigation_Questions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
