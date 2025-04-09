# Third Party Risk Management (TPRM) Dashboard

This **Streamlit** app is designed for Third Party Risk Management (TPRM). It allows users to assess inherent risks, capture mitigation responses, and calculate scores for supplier risk. The tool helps organizations track and manage their third-party suppliers to ensure compliance and identify areas of concern in a risk-conscious way.

## Features

### Section 1: Inherent Risk Assessment
- **Inherent Risk Questions**: The app prompts the user with a set of inherent risk questions related to third-party suppliers. The user answers with "Yes" or "No".
- Based on these responses, the **Mitigation Questions** section is dynamically adjusted.

### Section 2: Mitigation Questions
- **Mitigation Questions**: The app displays mitigation questions relevant to the responses given in the **Inherent Risk Assessment** section.
- The user can then provide supplier responses and score them on a scale from 0 to 3.

### Section 3: Summary & Download
- The app **calculates** scores for inherent risks and mitigation actions.
- Users can download an **Excel file** that summarizes all of their responses, scores, and suggested actions.

### Downloadable Summary
- **Excel Summary**: The app allows users to download the results of their assessments in an easy-to-read Excel format.

### Suggested Actions
- The app generates **suggested actions** based on risk scores:
  - **Escalate to Risk Owner & Reassess Supplier** for high-risk scenarios or weak mitigations.
  - **Monitor Closely** for moderate-risk scenarios.
  - **Proceed with Standard Controls** for low-risk scenarios.
