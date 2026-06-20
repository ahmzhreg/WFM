import streamlit as st
import pandas as pd
import os
from datetime import datetime
import smtplib
from email.message import EmailMessage

# --- System Setup ---
st.set_page_config(page_title="WFM Assessment Portal", page_icon="📊", layout="wide")

if not os.path.exists("uploaded_tests"):
    os.makedirs("uploaded_tests")

# --- Initialize Session State (The App's Memory) ---
# This ensures data isn't lost when the user moves between steps
if 'step' not in st.session_state:
    st.session_state.step = 1
    
session_keys = ["name", "email", "q1", "q2", "q3", "q4", "q5", "q6", "q7", "q8", "q9", "q10",
                "excel_data", "excel_filename", "q11", "q12", "q13"]
for key in session_keys:
    if key not in st.session_state:
        st.session_state[key] = None if key.startswith("q") and int(key[1:]) <= 10 else ""

# --- Helper Functions for Navigation & State ---
def next_step():
    st.session_state.step += 1

def prev_step():
    st.session_state.step -= 1

def get_idx(options, val):
    return options.index(val) if val in options else None

# --- Header & Progress Bar ---
st.title("📊 Boutiqaat WFM Team Leader Assessment")

if st.session_state.step < 4:
    st.progress(st.session_state.step / 3.0, text=f"Step {st.session_state.step} of 3")
    st.markdown("---")

# ==========================================
# STEP 1: THEORY & KPI MASTERY
# ==========================================
if st.session_state.step == 1:
    st.header("Step 1: Theory & Standards")
    st.write("Please fill out your details and complete the multiple-choice section.")
    
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.name = st.text_input("Full Name *", value=st.session_state.name)
    with col2:
        st.session_state.email = st.text_input("Email Address *", value=st.session_state.email)

    st.subheader("Forecasting & Capacity")
    opts_q1 = ["A) Simple Moving Average.", "B) Naive Forecasting.", "C) Multiple Regression Analysis with scenario modeling.", "D) Erlang C Volume Distribution."]
    st.session_state.q1 = st.radio("1. Which forecasting methodology is most appropriate for a Boutiqaat promotional event heavily influenced by marketing spend?", options=opts_q1, index=get_idx(opts_q1, st.session_state.q1))
    
    opts_q2 = ["A) ≤ ±2%", "B) ≤ ±5%", "C) ≤ ±10%", "D) ≤ ±15%"]
    st.session_state.q2 = st.radio("2. According to standard COPC/SWPP KPIs, what is the maximum acceptable intraday variance for Average Handle Time (AHT)?", options=opts_q2, index=get_idx(opts_q2, st.session_state.q2))

    opts_q3 = ["A) Callers never abandon the queue, regardless of wait time.", "B) Agents can handle multiple chats simultaneously.", "C) Handle times are perfectly consistent.", "D) Shrinkage is automatically accounted for in the raw calculation."]
    st.session_state.q3 = st.radio("3. A core assumption of the basic Erlang C formula is that:", options=opts_q3, index=get_idx(opts_q3, st.session_state.q3))

    st.subheader("Scheduling & Optimization")
    opts_q4 = ["A) High schedule shrinkage.", "B) Increased agent fatigue, potential burnout, and extended handle times.", "C) A drop in Schedule Conformance.", "D) Overstaffing leading to budget overruns."]
    st.session_state.q4 = st.radio("4. You observe Agent Schedule Adherence is 94% (Target ≥90%), but Occupancy has spiked to 92% (Target 75-85%). What is the most immediate risk?", options=opts_q4, index=get_idx(opts_q4, st.session_state.q4))

    opts_q5 = ["A) (Productive Time / Logged Time) x 100", "B) (Forecasted Volume / Actual Volume) x 100", "C) (Staffed FTE / Required FTE) x 100", "D) (Available Time / Handle Time) x 100"]
    st.session_state.q5 = st.radio("5. What is the standard formula for calculating Schedule Efficiency?", options=opts_q5, index=get_idx(opts_q5, st.session_state.q5))

    opts_q6 = ["A) 1-on-1 Coaching Sessions", "B) System Outage Downtime", "C) Annual Leave/Vacation", "D) Scheduled Team Meetings"]
    st.session_state.q6 = st.radio("6. Which of the following is considered 'Unplanned Shrinkage'?", options=opts_q6, index=get_idx(opts_q6, st.session_state.q6))

    st.subheader("Real-Time & Analytics")
    opts_q7 = ["A) ≤ 5 minutes", "B) ≤ 15 minutes", "C) ≤ 30 minutes", "D) ≤ 60 minutes"]
    st.session_state.q7 = st.radio("7. What is the industry standard target for Intraday Response Time?", options=opts_q7, index=get_idx(opts_q7, st.session_state.q7))

    opts_q8 = ["A) AHT is significantly lower than forecasted.", "B) Callers have an exceptionally low tolerance/patience threshold today.", "C) The IVR routing is highly efficient.", "D) Agents are spending too much time in After Call Work (ACW)."]
    st.session_state.q8 = st.radio("8. Abandonment Rate is climbing, but Service Level is perfectly meeting the 80/20 target. What is the most likely cause?", options=opts_q8, index=get_idx(opts_q8, st.session_state.q8))

    opts_q9 = ["A) Shrinkage Variance", "B) Schedule Adherence", "C) Plan-Driven CSAT/FCR Impact", "D) Agent Occupancy"]
    st.session_state.q9 = st.radio("9. Which metric provides the most accurate correlation between planning quality and Customer Experience (CX)?", options=opts_q9, index=get_idx(opts_q9, st.session_state.q9))

    opts_q10 = ["A) Plan", "B) Do", "C) Check", "D) Act"]
    st.session_state.q10 = st.radio("10. When applying PDCA to WFM Continuous Improvement, analyzing root causes for forecast misses falls under which phase?", options=opts_q10, index=get_idx(opts_q10, st.session_state.q10))

    st.divider()
    
    # Validation: Check if all Step 1 fields have a value
    step1_complete = all([
        st.session_state.name, st.session_state.email, st.session_state.q1, st.session_state.q2,
        st.session_state.q3, st.session_state.q4, st.session_state.q5, st.session_state.q6,
        st.session_state.q7, st.session_state.q8, st.session_state.q9, st.session_state.q10
    ])
    
    if not step1_complete:
        st.warning("⚠️ Please answer all questions and fill in your name/email to unlock the Next button.")

    st.button("Next: Section 2 ➡️", disabled=not step1_complete, on_click=next_step, type="primary")


# ==========================================
# STEP 2: EXCEL UPLOAD
# ==========================================
elif st.session_state.step == 2:
    st.header("Step 2: Practical Data Analysis")
    st.markdown("""
    Please download the **WFM Data Test** using the link below. Ensure you have completed all sheets before uploading your final version:
    1. Forecast Accuracy
    2. Erlang & FTE Calculation
    3. Schedule Efficiency
    4. Intraday Decision Logic
    5. Reporting Pivot & Dashboard
    """)
    st.markdown("### [📥 Click Here to Download the WFM_TL_Excel_Test.xlsx](https://docs.google.com/spreadsheets/d/1OCexYljty2ZZZzzgS8iTP8HssByQFQll/export?format=xlsx)")
    st.info("💡 **Instructions:** Once completed, save your file in the format `Firstname_Lastname_WFM_Test.xlsx` before uploading.")
    
    # Show success message if they already uploaded it and clicked back
    if st.session_state.excel_data:
        st.success(f"✅ File successfully saved: **{st.session_state.excel_filename}**. You may proceed, or upload a new file to replace it.")

    uploaded_excel = st.file_uploader("Upload your completed Excel test here *", type=["xlsx", "xls", "csv"])
    
    # Save the file immediately to memory so it isn't lost if they go to Step 3
    if uploaded_excel is not None:
        st.session_state.excel_data = uploaded_excel.getvalue()
        st.session_state.excel_filename = uploaded_excel.name

    st.divider()
    
    # Validation: Check if file is in memory
    step2_complete = bool(st.session_state.excel_data)
    
    if not step2_complete:
        st.warning("⚠️ Please upload your completed Excel test to unlock the Next button.")

    col1, col2 = st.columns([1, 5])
    with col1:
        st.button("⬅️ Back", on_click=prev_step)
    with col2:
        st.button("Next: Section 3 ➡️", disabled=not step2_complete, on_click=next_step, type="primary")


# ==========================================
# STEP 3: OPERATIONS FLOOR & SUBMISSION
# ==========================================
elif st.session_state.step == 3:
    st.header("Step 3: Operations Floor & Leadership")
    st.write("Provide structured, professional responses detailing your operational logic.")
    
    st.session_state.q11 = st.text_area("1. Intraday Crisis Simulation:\nIt is 2:00 PM. Inbound voice volume is 40% above forecast due to an early marketing launch. Abandonment rate is at 12%. SLA is breaching. You have an outbound/email team currently handling non-urgent back-office tickets. Outline your immediate corrective actions within the next 10 minutes:", value=st.session_state.q11, height=150)
    
    st.session_state.q12 = st.text_area("2. Agent Welfare & Fatigue Management:\nDescribe how you would balance cost-efficiency with agent well-being during a high-volume seasonal peak. How do you manage shift-bidding, overtime, and leave requests while maintaining transparency and fairness?", value=st.session_state.q12, height=150)
    
    st.session_state.q13 = st.text_area("3. Continuous Improvement & Automation:\nAs the functional owner of the WFM platform, you notice the team spends 3 hours a day manually consolidating reports in Excel. Walk through your strategy to automate this reporting cadence and ensure data integrity.", value=st.session_state.q13, height=150)

    st.divider()

    # Validation: Check if text boxes are filled
    step3_complete = bool(st.session_state.q11.strip() and st.session_state.q12.strip() and st.session_state.q13.strip())
    
    if not step3_complete:
        st.warning("⚠️ Please provide answers for all three operational scenarios to unlock the Submit button.")

    col1, col2 = st.columns([1, 5])
    with col1:
        st.button("⬅️ Back", on_click=prev_step)
    with col2:
        if st.button("🚀 Submit Final Assessment", disabled=not step3_complete, type="primary"):
            
            st.info("Processing submission and securely sending data... Please wait.")
            
            # 1. Auto-Score MCQs
            score = 0
            if st.session_state.q1.startswith("C"): score += 10
            if st.session_state.q2.startswith("B"): score += 10
            if st.session_state.q3.startswith("A"): score += 10
            if st.session_state.q4.startswith("B"): score += 10
            if st.session_state.q5.startswith("C"): score += 10
            if st.session_state.q6.startswith("B"): score += 10
            if st.session_state.q7.startswith("B"): score += 10
            if st.session_state.q8.startswith("B"): score += 10
            if st.session_state.q9.startswith("C"): score += 10
            if st.session_state.q10.startswith("C"): score += 10
            
            st.session_state.final_score = score

            # 2. Email Delivery Engine
            try:
                msg = EmailMessage()
                msg['Subject'] = f"🚨 New WFM Assessment: {st.session_state.name} - Score: {score}/100"
                msg['From'] = st.secrets["email_user"]
                msg['To'] = st.secrets["email_receiver"]
                
                email_body = f"""
                Candidate Assessment Completed!
                
                Name: {st.session_state.name}
                Email: {st.session_state.email}
                Theory Score: {score}/100
                
                --- Intraday Crisis Response ---
                {st.session_state.q11}
                
                --- Fatigue Management Response ---
                {st.session_state.q12}
                
                --- Automation Response ---
                {st.session_state.q13}
                """
                msg.set_content(email_body)
                
                # Attach Excel stored in session memory
                msg.add_attachment(
                    st.session_state.excel_data,
                    maintype='application',
                    subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    filename=f"{st.session_state.name.replace(' ', '_')}_WFM_Test.xlsx"
                )
                
                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(st.secrets["email_user"], st.secrets["email_password"])
                    smtp.send_message(msg)
                    
                # Move to Success Screen
                st.session_state.step = 4
                st.rerun()
                
            except Exception as e:
                st.error(f"An error occurred sending the email. Ensure Streamlit Secrets are set. Error details: {e}")

# ==========================================
# STEP 4: SUCCESS / COMPLETION SCREEN
# ==========================================
elif st.session_state.step == 4:
    st.success("✅ Assessment Submitted Successfully!")
    st.balloons()
    st.markdown(f"### **Your MCQ Theory Score:** {st.session_state.final_score} / 100")
    st.write("Thank you for your time. The Boutiqaat recruitment team has received your test and will review your Excel upload and operational responses shortly. You may now close this window.")
