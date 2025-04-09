import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO

# Load Inherent Risk Questions and Mitigation Question Bank
@st.cache
def load_data():
    inherent_data = pd.read_excel("path_to_your_file/TPRM_With_RAG_Colors.xlsx", sheet_name="Inherent Risk Assessment")
    mitigation_data = pd.read_excel("path_to_your_file/TPRM_With_RAG_Colors.xlsx", sheet_name="Mitigation Question Bank")
    return inherent_data, mitigation_data

# Section 1: Display Inherent Risk Questions
st.title("Third Party Risk Management (TPRM) - Inherent Risk Assessment")

# Load the questions from the Excel file
inherent_data, mitigation_data = load_data()

# Ask Inherent Risk Questions (Yes/No)
responses = []
for index, row in inherent_data.iterrows():
    question = row['Question']
    response = st.radio(f"{question}", ("Yes", "No"))
    responses.append(response)

# Section 2: Show Mitigation Questions based on Inherent Risk Responses
st.header("Mitigation Questions")

# Filter the mitigation questions based on responses
mitigation_questions_to_ask = []
for idx, response in enumerate(responses):
    if response == "Yes":  # Only ask mitigation questions for "Yes" responses
        inherent_question_id = inherent_data.iloc[idx]["ID"]
        domain = inherent_data.iloc[idx]["Risk Domain"]
        mitigation_questions_for_domain = mitigation_data[mitigation_data['Triggering Question ID'] == inherent_question_id]
        mitigation_questions_to_ask.append(mitigation_questions_for_domain)

# Flatten the list of questions to display
mitigation_questions_to_ask = pd.concat(mitigation_questions_to_ask)

# Let user answer mitigation questions and provide scores
mitigation_scores = []
for idx, row in mitigation_questions_to_ask.iterrows():
    question = row["Mitigation Question"]
    response = st.radio(f"{question}", ("Yes", "No"))
    score = st.selectbox(f"Score for {question}", (0, 1, 2, 3))
    mitigation_scores.append({"Question": question, "Response": response, "Score": score})

# Section 3: Summary & Download
st.header("Summary & Download")

# Calculate the summary based on Inherent Risk and Mitigation Scores
inherent_scores = [3 if response == "Yes" else 0 for response in responses]
mitigation_summary = pd.DataFrame(mitigation_scores)

# Prepare final data for download (in Excel format)
final_df = pd.DataFrame({
    "Inherent Risk Question": inherent_data["Question"],
    "Inherent Risk Score": inherent_scores,
    "Mitigation Question": mitigation_summary["Question"],
    "Mitigation Score": mitigation_summary["Score"]
})

# Convert DataFrame to Excel for download
@st.cache
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Summary")
    processed_data = output.getvalue()
    return processed_data

# Provide download link
download_button = st.download_button(
    label="Download Summary",
    data=to_excel(final_df),
    file_name="TPRM_Summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

