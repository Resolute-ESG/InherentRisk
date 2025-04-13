import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation

# Function to load the uploaded Excel file
def load_data(uploaded_file):
    try:
        # Load the user-uploaded file
        inherent_data = pd.read_excel(uploaded_file, sheet_name="Inherent Risk Assessment")
        mitigation_data = pd.read_excel(uploaded_file, sheet_name="Mitigation Question Bank")
    except Exception as e:
        st.error(f"Error loading the file: {e}")
        return None, None
    
    # Clean up column names by stripping spaces
    inherent_data.columns = inherent_data.columns.str.strip()
    mitigation_data.columns = mitigation_data.columns.str.strip()
    
    return inherent_data, mitigation_data

# Function to apply the original scoring logic
def apply_scoring_logic(supplier_response, sentiment):
    """ This function applies the scoring logic based on Supplier Response and Sentiment """
    if supplier_response == "Yes" and sentiment == "Positive":
        return 3
    elif supplier_response == "No" and sentiment == "Positive":
        return 0
    elif supplier_response == "No" and sentiment == "Negative":
        return 3
    elif supplier_response == "Yes" and sentiment == "Negative":
        return 0
    return 0  # Default case if no conditions match

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

# Stage 1: Upload the file
st.title("Third Party Risk Management (TPRM) - Inherent Risk Assessment & Scoring")

uploaded_file = st.file_uploader("Upload your Inherent Risk Assessment Excel file", type="xlsx")
if uploaded_file is not None:
    # Load the uploaded file
    inherent_data, mitigation_data = load_data(uploaded_file)
    
    if inherent_data is not None and mitigation_data is not None:
        # Display Inherent Risk Questions
        st.header("Inherent Risk Questions")
        st.write(inherent_data[['ID', 'Risk Domain', 'Question']])

        # Store the inherent data in session state
        if 'responses' not in st.session_state:
            st.session_state['responses'] = []

        if 'scores' not in st.session_state:
            st.session_state['scores'] = []

        # Generate Mitigation Questions based on Inherent Risk responses (only for 'Yes')
        responses = []
        for index, row in inherent_data.iterrows():
            response = st.radio(f"{row['Question']} (Yes/No)", ("Yes", "No"), key=f"response_{index}")
            responses.append(response)

        # Store responses in session state for future steps
        st.session_state['responses'] = responses

        # Filter the mitigation questions based on 'Yes' answers
        mitigation_questions_to_ask = []
        for idx, response in enumerate(responses):
            if response == "Yes":
                inherent_question_id = inherent_data.iloc[idx]["ID"]
                mitigation_questions_for_domain = mitigation_data[mitigation_data['Triggering Question ID'] == inherent_question_id]
                mitigation_questions_to_ask.append(mitigation_questions_for_domain)  # Fixed the parenthesis issue here
        
        # Flatten the list of mitigation questions
        mitigation_questions_to_ask = pd.concat(mitigation_questions_to_ask)

        # Add Supplier Response and Score columns (empty for now)
        mitigation_questions_to_ask['Supplier Response'] = 'Pending'
        mitigation_questions_to_ask['Sentiment'] = 'Pending'  # Add Sentiment column
        mitigation_questions_to_ask['Score'] = 0  # Add Score column here

        # Allow the user to download the mitigation questions
        if not mitigation_questions_to_ask.empty:
            st.download_button(
                label="Download Mitigation Questions for Scoring",
                data=to_excel(mitigation_questions_to_ask),
                file_name="Mitigation_Questions_for_Scoring.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No mitigation questions required based on your responses.")

# Step 2: Upload the downloaded Mitigation Questions for Scoring file
st.header("Stage 2: Upload the Mitigation Questions for Scoring File")

uploaded_scored_file = st.file_uploader("Upload the Mitigation Questions for Scoring file", type="xlsx")
scored_data = None
if uploaded_scored_file is not None:
    # Load the uploaded file and list sheet names
    try:
        # List all sheet names
        xls = pd.ExcelFile(uploaded_scored_file)
        sheet_names = xls.sheet_names
        st.write(f"Sheets available in the uploaded file: {sheet_names}")
        
        # Check if the "Mitigation Questions" sheet exists
        if "Mitigation Questions" in sheet_names:
            scored_data = pd.read_excel(uploaded_scored_file, sheet_name="Mitigation Questions")
        else:
            st.error("The 'Mitigation Questions' sheet is not found in the uploaded file. Please check the sheet name.")
            scored_data = None
        
    except Exception as e:
        st.error(f"Error loading the file: {e}")
    
    if scored_data is not None:
        # Check if the necessary columns exist, and if not, add them
        if "Supplier Response" not in scored_data.columns:
            scored_data['Supplier Response'] = 'Pending'  # Add if missing
        if "Sentiment" not in scored_data.columns:
            scored_data['Sentiment'] = 'Pending'  # Add if missing
        if "Score" not in scored_data.columns:
            scored_data['Score'] = 0  # Add if missing
        
        # Display the first few rows of the uploaded data
        st.write("Mitigation Questions uploaded successfully.")
        st.dataframe(scored_data.head())  # Display the first few rows of the uploaded data

        # Stage 2: Allow the user to score the questions inside the app
        for index, row in scored_data.iterrows():
            # Handle missing or invalid Supplier Response and Sentiment
            supplier_response = row['Supplier Response'] if row['Supplier Response'] in ["Yes", "No"] else "Yes"  # Default to "Yes" if invalid
            sentiment = row['Sentiment'] if row['Sentiment'] in ["Positive", "Negative"] else "Positive"  # Default to "Positive" if invalid
            
            # Use a more robust key that combines index and the question to ensure uniqueness
            key_response = f"response_{index}_{row['Mitigation Question']}"
            key_sentiment = f"sentiment_{index}_{row['Mitigation Question']}"
            key_score = f"score_{index}_{row['Mitigation Question']}"

            # Allow user to modify Supplier Response and Score in the app
            supplier_response = st.selectbox(
                f"Supplier Response for {row['Mitigation Question']}", 
                ["Yes", "No"], 
                key=key_response, 
                index=["Yes", "No"].index(supplier_response)
            )
            
            sentiment = st.selectbox(
                f"Sentiment for {row['Mitigation Question']}", 
                ["Positive", "Negative"], 
                key=key_sentiment, 
                index=["Positive", "Negative"].index(sentiment)
            )
            
            if row['Mitigation Type'] == 'Auto Scored':
                # Apply auto scoring logic if it's an auto-scored question
                score = apply_scoring_logic(supplier_response, sentiment)
            else:
                # Allow user to manually score for user-scored questions
                score = st.selectbox(f"Score for {row['Mitigation Question']}", [0, 1, 2, 3], key=key_score)

            # Update the Supplier Response, Sentiment, and Score in the data
            scored_data.at[index, "Supplier Response"] = supplier_response
            scored_data.at[index, "Sentiment"] = sentiment
            scored_data.at[index, "Score"] = score

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

# Step 3: Show Summary and Suggested Actions, only if the uploaded scored data exists
if scored_data is not None and 'Score' in scored_data.columns:
    st.header("Stage 3: Summary and Suggested Actions")
    
    # Apply the original logic for final scores based on Supplier Response and Score
    scored_data['Final Score'] = scored_data.apply(lambda row: apply_scoring_logic(row['Supplier Response'], row['Sentiment']) if row['Mitigation Type'] == 'Auto Scored' else row['Score'], axis=1)

    # Display summary of scores and recommended actions
    st.subheader("Final Scoring Summary")
    st.dataframe(scored_data)

    # Suggested actions based on scores
    st.subheader("Suggested Actions")
    for index, row in scored_data.iterrows():
        if row['Final Score'] == 0:
            st.warning(f"Question: {row['Mitigation Question']} - **Escalation Recommended**")
        elif row['Final Score'] == 3:
            st.success(f"Question: {row['Mitigation Question']} - **No Escalation Needed**")
        else:
            st.info(f"Question: {row['Mitigation Question']} - **Please review mitigation**")
