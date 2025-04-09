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
    
    # Add columns for Supplier Response and Score
    mitigation_questions_to_ask['Supplier Response'] = 'Pending'
    mitigation_questions_to_ask['Score'] = 0
    
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
            label="Download Mitigation Questions with Response & Score Columns",
            data=to_excel(mitigation_questions_to_ask),
            file_name="Mitigation_Questions_with_Scores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
Explanation of Stage 1 Changes:
After displaying the inherent risk questions, the app filters out the relevant mitigation questions (based on the responses).

The mitigation questions table now includes:

Supplier Response column (set to "Pending" by default).

Score column (set to 0 by default).

Download the Excel File: The user can download the Excel file, which includes the mitigation questions, supplier response placeholders, and score placeholders for later input.

Stage 2 Update: Allow Users to Upload Scored Mitigation Questions
In Stage 2, users will upload the Excel file they edited, which now includes their supplier responses and scores. We'll process the uploaded data to display it.

python
Copy
# Stage 2: Upload Scored Mitigation Questions
st.title("Stage 2: Upload Scored Mitigation Questions")

# Upload the scored mitigation questions file
uploaded_scored_file = st.file_uploader("Upload the file with scored mitigation questions", type="xlsx")
if uploaded_scored_file is not None:
    # Load the data from the uploaded scored mitigation questions file
    scored_data = pd.read_excel(uploaded_scored_file, sheet_name="Mitigation Questions")
    
    # Check if the necessary columns exist
    if "Supplier Response" not in scored_data.columns or "Score" not in scored_data.columns:
        st.error("The uploaded file is missing the required columns ('Supplier Response' and 'Score').")
    else:
        st.write("Scored Mitigation Questions uploaded successfully.")
        st.dataframe(scored_data.head())  # Display the first few rows of the uploaded data
Explanation of Stage 2:
Upload Scored File: After the user uploads their file with the responses and scores, the app checks if the required columns are present.

If the file is correct, it displays a preview of the uploaded data.

Stage 3 Update: Summarize the Scoring and Provide Suggested Actions
In Stage 3, we’ll summarize the Inherent Risk Scores and Mitigation Scores, and provide suggested actions based on the scores.

python
Copy
# Stage 3: Summarize Inherent and Mitigation Scores
st.title("Stage 3: Summarize Inherent and Mitigation Scores")

# Example calculation of scores (Replace with actual scoring logic)
inherent_scores = [3 if response == "Yes" else 0 for response in responses]
mitigation_scores = scored_data["Score"].tolist()  # Assuming mitigation scores are available

# Summarize the final scores
summary_df = pd.DataFrame({
    "Inherent Risk Question": inherent_data["Question"],
    "Inherent Risk Score": inherent_scores,
    "Mitigation Question": scored_data["Question"],
    "Mitigation Score": mitigation_scores
})

# Suggested actions based on scoring
summary_df["Suggested Action"] = summary_df.apply(lambda row: "Escalate to Risk Owner" if row["Mitigation Score"] < 2 else "Monitor and Proceed", axis=1)

# Display the summary
st.write("Summary of Inherent and Mitigation Scores:")
st.dataframe(summary_df)

# Provide a download button for the final summary
@st.cache_data
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Summary")
    processed_data = output.getvalue()
    return processed_data

# Download summary as Excel
st.download_button(
    label="Download Final Summary",
    data=to_excel(summary_df),
    file_name="Final_Summary.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
