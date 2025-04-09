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

# Function to apply the scoring logic based on Supplier Response and sentiment
def apply_scoring_logic(row):
    if row['Supplier Response'] == "Yes" and row['Sentiment'] == "Positive":
        return 3
    elif row['Supplier Response'] == "No" and row['Sentiment'] == "Positive":
        return 0
    elif row['Supplier Response'] == "No" and row['Sentiment'] == "Negative":
        return 3
    elif row['Supplier Response'] == "Yes" and row['Sentiment'] == "Negative":
        return 0
    return 0  # Default case

# Function to generate the Excel file for download
@st.cache_data
def to_excel(df, sheet_name="Mitigation Questions"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        sheet = workbook[sheet_name]
        
        # Add dropdown for Supplier Response ("Yes", "No")
        dv_response = DataValidation(type="list", formula1='"Yes,No"', showDropDown=True)
        sheet.add_data_validation(dv_response)
        dv_response.range = f"C2:C{len(df) + 1}"  # Supplier Response column
        
        # Add dropdown for Score (0, 1, 2, 3)
        dv_score = DataValidation(type="list", formula1='"0,1,2,3"', showDropDown=True)
        sheet.add_data_validation(dv_score)
        dv_score.range = f"D2:D{len(df) + 1}"  # Score column

        # Set column widths for readability
        sheet.column_dimensions['C'].width = 15  # Supplier Response column
        sheet.column_dimensions['D'].width = 10  # Score column

    processed_data = output.getvalue()
    return processed_data

# Stage 1 and Stage 2 combined
st.title("Third Party Risk Management (TPRM) - Inherent Risk Assessment & Scoring")

# Step 1: Upload the Inherent Risk Assessment file
uploaded_file = st.file_uploader("Upload your Inherent Risk Assessment Excel file", type="xlsx")
if uploaded_file is not None:
    # Load the data from the uploaded file
    inherent_data, mitigation_data = load_data(uploaded_file)
    
    # Stage 1: Display Inherent Risk Questions and allow for 'Yes' or 'No' responses
    st.header("Inherent Risk Questions")

    responses = []
    for index, row in inherent_data.iterrows():
        question = row['Question']
        response = st.radio(f"{question} (Yes/No)", ("Yes", "No"))
        responses.append(response)

    # Generate Mitigation Questions based on Inherent Risk Responses (only for 'Yes')
    mitigation_questions_to_ask = []
    for idx, response in enumerate(responses):
        if response == "Yes":
            inherent_question_id = inherent_data.iloc[idx]["ID"]
            mitigation_questions_for_domain = mitigation_data[mitigation_data['Triggering Question ID'] == inherent_question_id]
            mitigation_questions_to_ask.append(mitigation_questions_for_domain)

    # Flatten the list of mitigation questions
    mitigation_questions_to_ask = pd.concat(mitigation_questions_to_ask)
    
    # Add Supplier Response and Score columns (empty for now)
    mitigation_questions_to_ask['Supplier Response'] = 'Pending'
    mitigation_questions_to_ask['Score'] = 0

    # Allow the user to download the mitigation questions
    if mitigation_questions_to_ask.empty:
        st.warning("No mitigation questions required based on your responses.")
    else:
        st.download_button(
            label="Download Mitigation Questions for Scoring",
            data=to_excel(mitigation_questions_to_ask),
            file_name="Mitigation_Questions_for_Scoring.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Step 2: Upload the downloaded Mitigation Questions for Scoring file
    st.header("Stage 2: Upload the Mitigation Questions for Scoring File")

    uploaded_scored_file = st.file_uploader("Upload the Mitigation Questions for Scoring file", type="xlsx")
    if uploaded_scored_file is not None:
        # Load the uploaded file
        scored_data = pd.read_excel(uploaded_scored_file, sheet_name="Mitigation Questions")
        
        # Check if the necessary columns exist
        if "Supplier Response" not in scored_data.columns or "Score" not in scored_data.columns:
            st.error("The uploaded file is missing the required columns ('Supplier Response' and 'Score').")
        else:
            st.write("Mitigation Questions uploaded successfully.")
            st.dataframe(scored_data.head())  # Display the first few rows of the uploaded data

            # Stage 2: Allow the user to score the questions inside the app
            for index, row in scored_data.iterrows():
                supplier_response = st.selectbox(f"Supplier Response for {row['Mitigation Question']}", ["Yes", "No"], key=f"response_{index}")
                sentiment = st.selectbox(f"Sentiment for {row['Mitigation Question']}", ["Positive", "Negative"], key=f"sentiment_{index}")

                # Apply scoring logic based on the Supplier Response and Sentiment
                scored_data.at[index, "Supplier Response"] = supplier_response
                scored_data.at[index, "Sentiment"] = sentiment
                scored_data.at[index, "Score"] = apply_scoring_logic(scored_data.iloc[index])

            # Display the scored mitigation questions
            st.subheader("Scored Mitigation Questions")
            st.dataframe(scored_data)

            # Option to download the file with supplier responses and scores
            st.download_button(
                label="Download Scored Mitigation Questions",
                data=to_excel(scored_data, sheet_name="Scored Mitigation Questions"),
                file_name="Scored_Mitigation_Questions.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
