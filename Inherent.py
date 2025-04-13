# Step 2: Upload the downloaded Mitigation Questions for Scoring file
st.header("Stage 2: Upload the Mitigation Questions for Scoring File")

uploaded_scored_file = st.file_uploader("Upload the Mitigation Questions for Scoring file", type="xlsx")
if uploaded_scored_file is not None:
    # Load the uploaded file
    scored_data = pd.read_excel(uploaded_scored_file, sheet_name="Mitigation Questions")
    
    # Check if the necessary columns exist, and if not, add them
    if "Supplier Response" not in scored_data.columns:
        scored_data['Supplier Response'] = 'Pending'  # Add if missing
    if "Sentiment" not in scored_data.columns:
        scored_data['Sentiment'] = 'Pending'  # Add if missing
    if "Score" not in scored_data.columns:
        scored_data['Score'] = 0  # Add if missing
    
    # Debugging: Print column names to verify
    st.write("Columns in uploaded data:", scored_data.columns)

    # Display the first few rows of the uploaded data
    st.write("Mitigation Questions uploaded successfully.")
    st.dataframe(scored_data.head())  # Display the first few rows of the uploaded data

    # Stage 2: Allow the user to score the questions inside the app
    for index, row in scored_data.iterrows():
        # Handle missing or invalid Supplier Response and Sentiment
        supplier_response = row['Supplier Response'] if row['Supplier Response'] in ["Yes", "No"] else "Yes"  # Default to "Yes" if invalid
        sentiment = row['Sentiment'] if row['Sentiment'] in ["Positive", "Negative"] else "Positive"  # Default to "Positive" if invalid
        
        # Allow user to modify Supplier Response and Score in the app
        supplier_response = st.selectbox(f"Supplier Response for {row['Mitigation Question']}", ["Yes", "No"], key=f"response_{index}", index=["Yes", "No"].index(supplier_response))
        sentiment = st.selectbox(f"Sentiment for {row['Mitigation Question']}", ["Positive", "Negative"], key=f"sentiment_{index}", index=["Positive", "Negative"].index(sentiment))
        score = st.selectbox(f"Score for {row['Mitigation Question']}", [0, 1, 2, 3], key=f"score_{index}", index=[0, 1, 2, 3].index(row['Score']))

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

# Step 3: Show Summary and Suggested Actions
st.header("Stage 3: Summary and Suggested Actions")

if 'Score' in scored_data.columns:
    # Apply the original logic for final scores based on Supplier Response and Score
    scored_data['Final Score'] = scored_data.apply(apply_scoring_logic, axis=1)

    # Display summary of scores and recommended actions
    st.subheader("Final Scoring Summary")
    st.dataframe(scored_data)

    # Suggested actions based on scores
    if scored_data['Final Score'].max() < 3:
        st.warning("Some mitigation questions have a low score. It is recommended to escalate.")
    else:
        st.success("Mitigation scores are adequate. No escalation required.")
