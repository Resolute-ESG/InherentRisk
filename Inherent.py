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
        
        # Generate Mitigation Questions based on Inherent Risk responses (only for 'Yes')
        responses = []
        for index, row in inherent_data.iterrows():
            response = st.radio(f"{row['Question']} (Yes/No)", ("Yes", "No"), key=f"response_{index}")
            responses.append(response)
        
        # Filter the mitigation questions based on 'Yes' answers
        mitigation_questions_to_ask = []
        for idx, response in enumerate(responses):
            if response == "Yes":
                inherent_question_id = inherent_data.iloc[idx]["ID"]
                mitigation_questions_for_domain = mitigation_data[mitigation_data['Triggering Question ID'] == inherent_question_id]
                mitigation_questions_to_ask.append(mitigation_questions_for_domain
