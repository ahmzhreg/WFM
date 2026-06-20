import streamlit as st
import pandas as pd
import os
from datetime import datetime

# --- System Setup & Storage Directories ---
st.set_page_config(page_title="WFM Assessment Portal", page_icon="📊", layout="wide")

# Create folders to store Excel uploads if they don't exist
if not os.path.exists("uploaded_tests"):
    os.makedirs("uploaded_tests")

# --- Initialize Session State ---
if 'submitted' not in st.session_state:
    st.session_state.submitted = False

# --- Header & Candidate Info ---
st.title("📊 Boutiqaat WFM Team Leader Assessment")
st.markdown("""
Welcome to the technical evaluation. This assessment measures your proficiency in capacity planning, schedule efficiency, and real-time operations according to SWPP and COPC standards.
* **Time limit:** 90 minutes
* **Required:** Please complete all multiple-choice questions, operational scenarios, and upload your completed Excel test.
---
""")

col1, col2 = st.columns(2)
with col1:
    candidate_name = st.text_input("Full Name *")
with col2:
    candidate_email = st.text_input("Email Address *")

# --- Assessment Tabs ---
tab1, tab2, tab3 = st.tabs(["📚 Section 1: Theory & KPI Mastery", "🧮 Section 2: Excel Test Upload", "📝 Section 3: Intraday Operations"])

with tab1:
    st.header("Section 1: WFM Theory")
    st.write("Select the most appropriate answer for each scenario based on standard Contact Center practices.")
    
    q1 = st.radio("1. You observe Agent Schedule Adherence is 94%, but Occupancy has spiked to 92%. What is the most immediate risk?",
        options=["A) High schedule shrinkage and increased non-productive time.", "B) Increased agent fatigue, potential burnout, and longer handle times.", "C) A drop in Schedule Adherence.", "D) Overstaffing leading to budget overruns."], index=None)
    
    q2 = st.radio("2. Which forecasting methodology is most appropriate for a promotional event heavily influenced by marketing spend?",
        options=["A) Simple Moving Average.", "B) Naive Forecasting.", "C) Multiple Regression Analysis with scenario modeling.", "D) Erlang C Volume Distribution."], index=None)
    
    q3 = st.radio("3. According to standard Boutiqaat KPIs, what is the maximum acceptable intraday variance for Average Handle Time (AHT)?",
        options=["A) ≤ ±5%", "B) ≤ ±10%", "C) ≤ ±15%", "D) ≤ ±2%"], index=None)

    q4 = st.radio("4. What is the standard formula for calculating Schedule Efficiency?",
        options=["A) (Staffed FTE / Required FTE) x 100", "B) (Productive Time / Logged Time) x 100", "C) (Forecasted Volume / Actual Volume) x 100", "D) (Available Time / Handle Time) x 100"], index=None)

with tab2:
    st.header("Section 2: Practical Data Analysis")
    st.markdown("""
    Please complete the **WFM_TL_Excel_Test.xlsx** provided to you by the recruitment team. 
    Ensure you have completed all sheets (Forecast Accuracy, Erlang & FTE, Schedule Efficiency, Intraday Decision, and Reporting Pivot) before uploading.
    """)
    
    st.info("💡 **Instructions:** Save your completed file in the format `Firstname_Lastname_WFM_Test.xlsx` before uploading.")
    
    uploaded_excel = st.file_uploader("Upload your completed Excel test here *", type=["xlsx", "xls", "csv"])

with tab3:
    st.header("Section 3: Operations Floor & Leadership")
    
    crisis_response = st.text_area("Intraday Crisis Simulation:\nIt is 2:00 PM. Inbound voice volume is 40% above forecast due to an early marketing launch. Abandonment rate is at 12%. You have an outbound team currently handling non-urgent back-office tickets. Outline your immediate corrective actions within the next 10 minutes:", height=150)
    
    fatigue_response = st.text_area("Fatigue Management & Fairness:\nDescribe how you would balance cost-efficiency with agent well-being during a high-volume seasonal peak while maintaining transparent shift-bidding and leave allocation:", height=150)

st.divider()

# --- Submission & Data Saving Logic ---
if st.button("Submit Assessment", type="primary"):
    if not candidate_name or not candidate_email or not uploaded_excel:
        st.error("⚠️ Please fill in your Name, Email, and upload the Excel test before submitting.")
    else:
        # 1. Save the uploaded Excel file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = candidate_name.replace(" ", "_")
        file_path = f"uploaded_tests/{safe_name}_{timestamp}_{uploaded_excel.name}"
        
        with open(file_path, "wb") as f:
            f.write(uploaded_excel.getbuffer())
        
        # 2. Auto-Score the MCQs
        score = 0
        if q1 and q1.startswith("B"): score += 10
        if q2 and q2.startswith("C"): score += 10
        if q3 and q3.startswith("A"): score += 10
        if q4 and q4.startswith("A"): score += 10
        
        # 3. Save response data to a CSV
        submission_data = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Candidate Name": candidate_name,
            "Email": candidate_email,
            "MCQ Score (Max 40)": score,
            "Excel File Path": file_path,
            "Q1 Answer": q1,
            "Q2 Answer": q2,
            "Q3 Answer": q3,
            "Q4 Answer": q4,
            "Crisis Response": crisis_response,
            "Fatigue Response": fatigue_response
        }
        
        df = pd.DataFrame([submission_data])
        csv_file = "candidate_submissions.csv"
        
        # Append to CSV (create if it doesn't exist)
        if not os.path.isfile(csv_file):
            df.to_csv(csv_file, index=False)
        else:
            df.to_csv(csv_file, mode='a', header=False, index=False)
        
        # 4. Display Success Message
        st.session_state.submitted = True
        st.success("✅ Assessment Submitted Successfully!")
        st.write("Thank you for your time. The recruitment team will review your Excel upload and operational responses.")
