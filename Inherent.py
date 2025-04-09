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
    mitigation
