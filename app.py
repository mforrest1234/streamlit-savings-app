import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- Streamlit UI ---
st.title("💰 Financial Savings Calculator + Google Slides Export")

# Inputs
num_employees = st.number_input("Number of Employees", min_value=1, value=50)
avg_salary = st.number_input("Average Salary per Employee ($)", min_value=0.0, value=60000.0, step=1000.0)
improvement_pct = st.slider("Improvement (%)", min_value=0, max_value=100, value=10)

# Calculation
total_savings = num_employees * avg_salary * (improvement_pct / 100)

st.subheader("📈 Estimated Total Savings")
st.metric(label="Total Savings", value=f"${total_savings:,.2f}")

# --- Google Slides Export ---
st.markdown("## 📤 Export to Google Slides")

if st.button("Generate Slide Deck"):
    # Load service account credentials
    SERVICE_ACCOUNT_FILE = 'service_account.json'
    SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive']
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    # Build APIs
    slides_service = build('slides', 'v1', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

    # === Replace with your actual Google Slides template ID ===
    TEMPLATE_ID = "YOUR_TEMPLATE_PRESENTATION_ID"

    # Copy the template
    copy_title = "Financial Savings Report"
    copied_file = drive_service.files().copy(fileId=TEMPLATE_ID, body={'name': copy_title}).execute()
    new_presentation_id = copied_file['id']

    # Define replacements
    requests = [
        {'replaceAllText': {
            'containsText': {'text': '{{num_employees}}'},
            'replaceText': str(num_employees)
        }},
        {'replaceAllText': {
            'containsText': {'text': '{{avg_salary}}'},
            'replaceText': f"${avg_salary:,.2f}"
        }},
        {'replaceAllText': {
            'containsText': {'text': '{{improvement_pct}}'},
            'replaceText': f"{improvement_pct}%"
        }},
        {'replaceAllText': {
            'containsText': {'text': '{{total_savings}}'},
            'replaceText': f"${total_savings:,.2f}"
        }},
    ]

    # Update the slide
    slides_service.presentations().batchUpdate(
        presentationId=new_presentation_id,
        body={'requests': requests}
    ).execute()

    # Share link
    link = f"https://docs.google.com/presentation/d/1ycgNhIgPTGJBLXk656dlYt758YCoWtCqhMYcznPG4VU/edit?usp=sharing"
    st.success("✅ Slide deck created!")
    st.markdown(f"[👉 View Slides]({link})", unsafe_allow_html=True)
