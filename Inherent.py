import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

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
    
    # Stage 1: Ask Inherent Risk Questions (Yes/No and Positive/Negative)
    st.header("Inherent Risk Questions")

    responses = []
    scores = []

    for index, row in inherent_data.iterrows():
        question = row['Question']
        response = st.radio(f"{question} (Yes/No)", ("Yes", "No"))
        sentiment = st.radio(f"Sentiment for '{question}' (Positive/Negative)", ("Positive", "Negative"))

        # Apply scoring logic based on the response and sentiment
        if response == "Yes" and sentiment == "Positive":
            score = 3
        elif response == "No" and sentiment == "Positive":
            score = 0
        elif response == "No" and sentiment == "Negative":
            score = 3
        elif response == "Yes" and sentiment == "Negative":
            score = 0

        responses.append(response)
        scores.append(score)

    # Generate Mitigation Questions based on Inherent Risk Responses
    mitigation_questions_to_ask = []
    for idx, response in enumerate(responses):
        if response == "Yes":  # Only ask mitigation questions for "Yes" responses
            inherent_question_id = inherent_data.iloc[idx]["ID"]
            mitigation_questions_for_domain = mitigation_data[mitigation_data['Triggering Question ID'] == inherent_question_id]
            mitigation_questions_to_ask.append(mitigation_questions_for_domain)

    # Flatten the list of questions to display
    mitigation_questions_to_ask = pd.concat(mitigation_questions_to_ask)
    
    # Add columns for Supplier Response and Score
    mitigation_questions_to_ask['Supplier Response'] = 'Pending'
    mitigation_questions_to_ask['Score'] = 0

    # Prepare the Excel file for download with the mitigation questions and dropdown options
    @st.cache_data
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Mitigation Questions")
            workbook = writer.book
            sheet = workbook["Mitigation Questions"]
            
            # Add dropdown for Supplier Response ("Yes", "No")
            dv_response = DataValidation(type="list", formula1='"Yes,No"', showDropDown=True)
            sheet.add_data_validation(dv_response)
            dv_response.range = f"C2:C{len(df) + 1}"  # Supplier Response column
            
            # Add dropdown for Score (0, 1, 2, 3)
            dv_score = DataValidation(type="list", formula1='"0,1,2,3"', showDropDown=True)
            sheet.add_data_validation(dv_score)
            dv_score.range = f"D2:D{len(df) + 1}"  # Score column
            
        processed_data = output.getvalue()
        return processed_data

    # Display a download button for the mitigation questions Excel file
    if mitigation_questions_to_ask.empty:
        st.warning("No mitigation questions required based on your responses.")
    else:
        st.download_button(
            label="Download Mitigation Questions with Response & Score Columns",
            data=to_excel(mitigation_questions_to_ask),
            file_name="Mitigation_Questions_with_Scores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Display the scores for Inherent Risk Questions
    st.subheader("Inherent Risk Scores")
    score_df = pd.DataFrame({
        "Inherent Risk Question": inherent_data["Question"],
        "Response": responses,
        "Sentiment": ["Positive" if r == "Positive" else "Negative" for r in responses],
        "Score": scores
    })
    st.dataframe(score_df)

