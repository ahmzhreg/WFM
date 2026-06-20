import streamlit as st
import pandas as pd
import os
from datetime import datetime

# --- System Setup ---
st.set_page_config(page_title="WFM Assessment Portal", page_icon="📊", layout="wide")

if not os.path.exists("uploaded_tests"):
    os.makedirs("uploaded_tests")

if 'submitted' not in st.session_state:
    st.session_state.submitted = False

# --- Header & Candidate Info ---
st.title("📊 Boutiqaat WFM Team Leader Assessment")
st.markdown("""
Welcome to the technical evaluation. This assessment measures your proficiency in capacity planning, schedule efficiency, and real-time operations according to SWPP and COPC standards.
* **Time limit:** 90 minutes
* **Required:** Complete all Theory MCQs, operational scenarios, and upload your completed Excel Data Test.
---
""")

col1, col2 = st.columns(2)
with col1:
    candidate_name = st.text_input("Full Name *")
with col2:
    candidate_email = st.text_input("Email Address *")

# --- Assessment Tabs ---
tab1, tab2, tab3 = st.tabs(["📚 Section 1: Theory & KPI Mastery", "🧮 Section 2: Excel Test Upload", "📝 Section 3: Operations Floor"])

# --- SECTION 1: MCQs ---
with tab1:
    st.header("Section 1: WFM Theory & Standards")
    st.write("Select the most appropriate answer for each scenario.")
    
    st.subheader("Forecasting & Capacity")
    q1 = st.radio("1. Which forecasting methodology is most appropriate for a Boutiqaat promotional event heavily influenced by marketing spend and new-product launches?",
        options=["A) Simple Moving Average.", "B) Naive Forecasting.", "C) Multiple Regression Analysis with scenario modeling.", "D) Erlang C Volume Distribution."], index=None)
    
    q2 = st.radio("2. According to standard COPC/SWPP KPIs, what is the maximum acceptable intraday variance for Average Handle Time (AHT)?",
        options=["A) ≤ ±2%", "B) ≤ ±5%", "C) ≤ ±10%", "D) ≤ ±15%"], index=None)

    q3 = st.radio("3. A core assumption of the basic Erlang C formula is that:",
        options=["A) Callers never abandon the queue, regardless of wait time.", "B) Agents can handle multiple chats simultaneously.", "C) Handle times are perfectly consistent.", "D) Shrinkage is automatically accounted for in the raw calculation."], index=None)

    st.subheader("Scheduling & Optimization")
    q4 = st.radio("4. You observe Agent Schedule Adherence is 94% (Target ≥90%), but Occupancy has spiked to 92% (Target 75-85%). What is the most immediate risk?",
        options=["A) High schedule shrinkage.", "B) Increased agent fatigue, potential burnout, and extended handle times.", "C) A drop in Schedule Conformance.", "D) Overstaffing leading to budget overruns."], index=None)

    q5 = st.radio("5. What is the standard formula for calculating Schedule Efficiency?",
        options=["A) (Productive Time / Logged Time) x 100", "B) (Forecasted Volume / Actual Volume) x 100", "C) (Staffed FTE / Required FTE) x 100", "D) (Available Time / Handle Time) x 100"], index=None)

    q6 = st.radio("6. Which of the following is considered 'Unplanned Shrinkage'?",
        options=["A) 1-on-1 Coaching Sessions", "B) System Outage Downtime", "C) Annual Leave/Vacation", "D) Scheduled Team Meetings"], index=None)

    st.subheader("Real-Time & Analytics")
    q7 = st.radio("7. What is the industry standard target for Intraday Response Time (median time to detect and act on intraday deviations)?",
        options=["A) ≤ 5 minutes", "B) ≤ 15 minutes", "C) ≤ 30 minutes", "D) ≤ 60 minutes"], index=None)

    q8 = st.radio("8. Abandonment Rate is climbing, but Service Level is perfectly meeting the 80/20 target. What is the most likely cause?",
        options=["A) AHT is significantly lower than forecasted.", "B) Callers have an exceptionally low tolerance/patience threshold today.", "C) The IVR routing is highly efficient.", "D) Agents are spending too much time in After Call Work (ACW)."], index=None)

    q9 = st.radio("9. Which metric provides the most accurate correlation between planning quality and Customer Experience (CX)?",
        options=["A) Shrinkage Variance", "B) Schedule Adherence", "C) Plan-Driven CSAT/FCR Impact", "D) Agent Occupancy"], index=None)

    q10 = st.radio("10. When applying PDCA (Plan-Do-Check-Act) to WFM Continuous Improvement, analyzing root causes for forecast misses falls under which phase?",
        options=["A) Plan", "B) Do", "C) Check", "D) Act"], index=None)


# --- SECTION 2: EXCEL UPLOAD ---
with tab2:
    st.header("Section 2: Practical Data Analysis")
    st.markdown("""
    Please complete the **WFM_TL_Excel_Test.xlsx** provided to you by the recruitment team. 
    Ensure you have completed all sheets before uploading:
    1. Forecast Accuracy
    2. Erlang & FTE Calculation
    3. Schedule Efficiency
    4. Intraday Decision Logic
    5. Reporting Pivot & Dashboard
    """)
    st.info("💡 **Instructions:** Save your completed file in the format `Firstname_Lastname_WFM_Test.xlsx` before uploading.")
    uploaded_excel = st.file_uploader("Upload your completed Excel test here *", type=["xlsx", "xls", "csv"])


# --- SECTION 3: OPEN-ENDED ---
with tab3:
    st.header("Section 3: Operations Floor & Leadership")
    st.write("Provide structured, professional responses detailing your operational logic.")
    
    q11 = st.text_area("1. Intraday Crisis Simulation:\nIt is 2:00 PM. Inbound voice volume is 40% above forecast due to an early marketing launch. Abandonment rate is at 12% (Target ≤ 5%). Intraday response SLA is breaching. You have an outbound/email team currently handling non-urgent back-office tickets. Outline your immediate corrective actions within the next 10 minutes:", height=150)
    
    q12 = st.text_area("2. Agent Welfare & Fatigue Management:\nDescribe how you would balance cost-efficiency with agent well-being during a high-volume seasonal peak (e.g., Ramadan or White Friday). How do you manage shift-bidding, overtime, and leave requests while maintaining transparency and fairness?", height=150)
    
    q13 = st.text_area("3. Continuous Improvement & Automation:\nAs the functional owner of the WFM platform, you notice the team spends 3 hours a day manually consolidating reports in Excel. Walk through your strategy to automate this reporting cadence and ensure data integrity.", height=150)

st.divider()

# --- SUBMISSION LOGIC ---
if st.button("Submit Assessment", type="primary"):
    if not candidate_name or not candidate_email or not uploaded_excel:
        st.error("⚠️ Please fill in your Name, Email, and upload the Excel test before submitting.")
    else:
        # Save Excel File
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = candidate_name.replace(" ", "_")
        file_path = f"uploaded_tests/{safe_name}_{timestamp}_{uploaded_excel.name}"
        
        with open(file_path, "wb") as f:
            f.write(uploaded_excel.getbuffer())
        
        # Auto-Score MCQs (10 points each)
        score = 0
        if q1 and q1.startswith("C"): score += 10
        if q2 and q2.startswith("B"): score += 10
        if q3 and q3.startswith("A"): score += 10
        if q4 and q4.startswith("B"): score += 10
        if q5 and q5.startswith("C"): score += 10
        if q6 and q6.startswith("B"): score += 10
        if q7 and q7.startswith("B"): score += 10
        if q8 and q8.startswith("B"): score += 10
        if q9 and q9.startswith("C"): score += 10
        if q10 and q10.startswith("C"): score += 10
        
        # Compile Data
        submission_data = {
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Candidate Name": candidate_name,
            "Email": candidate_email,
            "MCQ Score (Max 100)": score,
            "Excel File Path": file_path,
            "Q1": q1, "Q2": q2, "Q3": q3, "Q4": q4, "Q5": q5,
            "Q6": q6, "Q7": q7, "Q8": q8, "Q9": q9, "Q10": q10,
            "Crisis Response": q11,
            "Fatigue Response": q12,
            "Automation Response": q13
        }
        
        # Save to CSV
        df = pd.DataFrame([submission_data])
        csv_file = "candidate_submissions.csv"
        if not os.path.isfile(csv_file):
            df.to_csv(csv_file, index=False)
        else:
            df.to_csv(csv_file, mode='a', header=False, index=False)
        
        st.session_state.submitted = True
        st.success("✅ Assessment Submitted Successfully!")
        st.write(f"**Your MCQ Theory Score:** {score} / 100")
        st.write("Thank you for your time. The recruitment team will review your Excel upload and operational responses shortly.")
