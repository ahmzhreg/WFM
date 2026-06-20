import streamlit as st
import pandas as pd
import os
import smtplib
from email.message import EmailMessage
import openpyxl
from io import BytesIO

# --- System Setup ---
st.set_page_config(page_title="WFM Assessment Portal", page_icon="📊", layout="wide")

if not os.path.exists("uploaded_tests"):
    os.makedirs("uploaded_tests")

# --- Initialize Session State ---
if 'step' not in st.session_state:
    st.session_state.step = 1
    
session_keys = [
    "name", "email",
    "q1", "q2", "q3", "q4", "q5", "q6", "q7", "q8", "q9", "q10", "q11", "q12", "q13", "q14", "q15",
    "b1", "b2", "b3", "b4", "b5",
    "excel_data", "excel_filename", "final_theory", "final_excel"
]
for key in session_keys:
    if key not in st.session_state:
        st.session_state[key] = None if key.startswith("q") else ""
        if key in ["final_theory", "final_excel"]:
            st.session_state[key] = 0

def next_step(): st.session_state.step += 1
def prev_step(): st.session_state.step -= 1
def get_idx(options, val): return options.index(val) if val in options else None

# --- Header ---
st.title("📊 Boutiqaat WFM Team Leader Assessment")
if st.session_state.step < 4:
    st.progress(st.session_state.step / 3.0, text=f"Step {st.session_state.step} of 3")
    st.markdown("---")

# ==========================================
# STEP 1: THEORY & KPI MASTERY (15 MCQs)
# ==========================================
if st.session_state.step == 1:
    st.header("Step 1: Theory & Standards")
    st.write("Please complete the theoretical multiple-choice section (COPC & SWPP Aligned).")
    
    col1, col2 = st.columns(2)
    with col1: st.session_state.name = st.text_input("Full Name *", value=st.session_state.name)
    with col2: st.session_state.email = st.text_input("Email Address *", value=st.session_state.email)

    opts_q1 = ["A) 13.0% — within target", "B) 15.0% — outside target", "C) 15.0% — within target", "D) 18.0% — outside target"]
    st.session_state.q1 = st.radio("1. A voice queue had 1,200 calls forecasted for an interval and 1,380 actually arrived. What is the absolute forecast variance, and does it meet the JD intraday target (≤ ±10%)?", opts_q1, index=get_idx(opts_q1, st.session_state.q1))

    opts_q2 = ["A) R × (1 − S)", "B) R / (1 − S)", "C) R + S", "D) R / S"]
    st.session_state.q2 = st.radio("2. Which formula correctly converts required agents (R) into required FTE given a shrinkage rate S?", opts_q2, index=get_idx(opts_q2, st.session_state.q2))

    opts_q3 = ["A) Adherence and conformance are synonyms.", "B) Adherence measures total hours worked vs scheduled; conformance measures activity timing.", "C) Adherence measures activity timing vs schedule; conformance measures total time worked regardless of timing.", "D) Conformance applies to email only; adherence applies to voice."]
    st.session_state.q3 = st.radio("3. Schedule adherence and conformance are often confused. Which statement is correct?", opts_q3, index=get_idx(opts_q3, st.session_state.q3))

    opts_q4 = ["A) Erlang B", "B) Erlang C", "C) Erlang A", "D) Erlang X with no retry input"]
    st.session_state.q4 = st.radio("4. Which Erlang variant explicitly models customer abandonment behaviour?", opts_q4, index=get_idx(opts_q4, st.session_state.q4))

    opts_q5 = ["A) Hours worked / hours paid", "B) Staffed FTE / required FTE, measured at the interval level and aggregated", "C) Service level / target service level", "D) Adherence percentage of the top quartile of agents"]
    st.session_state.q5 = st.radio("5. JD target for Schedule Efficiency is ≥ 95%. Which is the best operating definition?", opts_q5, index=get_idx(opts_q5, st.session_state.q5))

    opts_q6 = ["A) Improved SL with no side effects", "B) Lower AHT due to focus", "C) Higher attrition and AHT drift due to agent burnout", "D) Lower shrinkage"]
    st.session_state.q6 = st.radio("6. Occupancy is sustained at 92% across a quarter. What is the most likely operational consequence?", opts_q6, index=get_idx(opts_q6, st.session_state.q6))

    opts_q7 = ["A) Sizing an interval given AHT and SL target", "B) Long-range volume forecasting that must separate trend, seasonality, and campaign effects", "C) Calculating intraday queue depth in real time", "D) Determining required agents to meet 80/20 SL"]
    st.session_state.q7 = st.radio("7. Which scenario is the strongest case for using time-series decomposition over Erlang C for planning?", opts_q7, index=get_idx(opts_q7, st.session_state.q7))

    opts_q8 = ["A) Trigger overtime for the next shift", "B) Re-skill or pull agents from a lower-priority queue, then notify shift owner", "C) Send a global broadcast to all staff", "D) Recalculate the forecast"]
    st.session_state.q8 = st.radio("8. A 30-minute interval has an SL of 64% (target 80%). Which intraday action is generally first, before escalation?", opts_q8, index=get_idx(opts_q8, st.session_state.q8))

    opts_q9 = ["A) Pure year-over-year multiplier", "B) Erlang C with a 3x arrival rate", "C) Time-series baseline + campaign uplift model, validated against last campaign's actuals", "D) Same staffing as previous Friday plus 25% buffer"]
    st.session_state.q9 = st.radio("9. Boutiqaat is running a White Friday campaign with predicted 3x volume. Which forecasting approach is most defensible?", opts_q9, index=get_idx(opts_q9, st.session_state.q9))

    opts_q10 = ["A) Coaching and 1:1s", "B) Training", "C) Talk time on customer contacts", "D) Unplanned absence"]
    st.session_state.q10 = st.radio("10. Which is NOT typically counted in shrinkage?", opts_q10, index=get_idx(opts_q10, st.session_state.q10))

    opts_q11 = ["A) Forecast must always use Erlang C", "B) Documented methodology, tracked accuracy, and root-cause review of misses", "C) Forecasts must be locked 12 weeks ahead", "D) Only one channel may be forecasted at a time"]
    st.session_state.q11 = st.radio("11. The COPC CX Standard requires which of the following for forecasting?", opts_q11, index=get_idx(opts_q11, st.session_state.q11))

    opts_q12 = ["A) 20 agents", "B) 40 agents", "C) 60 agents", "D) 80 agents"]
    st.session_state.q12 = st.radio("12. Chat channel: average concurrency = 2.0, AHT (per chat) = 8 minutes, forecast = 600 chats in a 60-min interval. Roughly how many concurrent chat agents are needed (ignoring shrinkage and SL buffer)?", opts_q12, index=get_idx(opts_q12, st.session_state.q12))

    opts_q13 = ["A) Forecast accuracy and shrinkage", "B) SL, longest wait, and agent state mix", "C) NPS and CSAT", "D) Attrition and FCR"]
    st.session_state.q13 = st.radio("13. Which is the best metric pairing to monitor in real-time on a voice queue?", opts_q13, index=get_idx(opts_q13, st.session_state.q13))

    opts_q14 = ["A) Discipline the agents immediately", "B) Investigate via QA tools for system or contact-type root cause before action", "C) Force-end ACW for all 4", "D) Ignore until end-of-day"]
    st.session_state.q14 = st.radio("14. A Real-Time Officer flags that 4 agents are in After-Call Work (ACW) for >5 minutes during a backlog. What is the right first response?", opts_q14, index=get_idx(opts_q14, st.session_state.q14))

    opts_q15 = ["A) Reducing AHT", "B) Improving fairness perception, engagement, and adherence to chosen shifts", "C) Reducing forecast error", "D) Replacing the need for capacity planning"]
    st.session_state.q15 = st.radio("15. A scheduling policy that lets agents bid for shifts based on tenure has the primary advantage of:", opts_q15, index=get_idx(opts_q15, st.session_state.q15))

    st.divider()
    step1_complete = all([st.session_state.name, st.session_state.email] + [st.session_state[f"q{i}"] for i in range(1, 16)])
    if not step1_complete: st.warning("⚠️ Please answer all 15 questions and provide your details.")
    st.button("Next: Section 2 ➡️", disabled=not step1_complete, on_click=next_step, type="primary")

# ==========================================
# STEP 2: EXCEL UPLOAD
# ==========================================
elif st.session_state.step == 2:
    st.header("Step 2: Practical Data Analysis")
    st.markdown("### [📥 Click Here to Download the WFM_TL_Excel_Test.xlsx](https://docs.google.com/spreadsheets/d/1OCexYljty2ZZZzzgS8iTP8HssByQFQll/export?format=xlsx)")
    
    if st.session_state.excel_data:
        st.success(f"✅ File successfully saved: **{st.session_state.excel_filename}**")

    uploaded_excel = st.file_uploader("Upload your completed Excel test here *", type=["xlsx", "xls"])
    if uploaded_excel:
        st.session_state.excel_data = uploaded_excel.getvalue()
        st.session_state.excel_filename = uploaded_excel.name

    st.divider()
    step2_complete = bool(st.session_state.excel_data)
    
    col1, col2 = st.columns([1, 5])
    with col1: st.button("⬅️ Back", on_click=prev_step)
    with col2: st.button("Next: Section 3 ➡️", disabled=not step2_complete, on_click=next_step, type="primary")

# ==========================================
# STEP 3: OPEN-ENDED SCENARIOS
# ==========================================
elif st.session_state.step == 3:
    st.header("Step 3: Operations Floor & Leadership Scenarios")
    
    st.session_state.b1 = st.text_area("B1. Describe your end-to-end weekly forecasting cycle for a multi-channel contact centre. What inputs do you use, what outputs do you produce, and how do you track accuracy?", value=st.session_state.b1, height=150)
    st.session_state.b2 = st.text_area("B2. Walk through how you would design a schedule that improves Schedule Efficiency from 88% to ≥95% without increasing headcount. What levers would you use?", value=st.session_state.b2, height=150)
    st.session_state.b3 = st.text_area("B3. Describe an intraday escalation matrix you would implement for Boutiqaat. Include thresholds, triggers, owners, and a measurable response-time target.", value=st.session_state.b3, height=150)
    st.session_state.b4 = st.text_area("B4. How would you measure and improve forecast accuracy when a sudden influencer launch creates a same-day demand spike that historical data cannot predict?", value=st.session_state.b4, height=150)
    st.session_state.b5 = st.text_area("B5. You lead a team. One Officer is technically strong but defensive when feedback is given. How do you coach this person while protecting team performance?", value=st.session_state.b5, height=150)

    st.divider()
    step3_complete = all(st.session_state[f"b{i}"].strip() for i in range(1, 6))

    col1, col2 = st.columns([1, 5])
    with col1: st.button("⬅️ Back", on_click=prev_step)
    with col2:
        if st.button("🚀 Submit Final Assessment", disabled=not step3_complete, type="primary"):
            st.info("Grading assessment and securely sending data... Please wait.")
            
            # --- 1. THEORY SCORE (MCQ Key) ---
            t_score = 0
            correct_keys = ["B","B","C","C","B","C","B","B","C","C","B","B","B","B","B"]
            for i, ans in enumerate(correct_keys, 1):
                if st.session_state[f"q{i}"].startswith(ans): t_score += 1
            st.session_state.final_theory = round((t_score / 15) * 100, 1)

            # --- 2. EXCEL AUTO-SCORE ENGINE ---
            excel_score = 0
            
            # ---> UPDATE YOUR EXACT CELL COORDINATES HERE <---
            CELL_MAP = {
                "vol_mape": "B17",    # Cell containing ~ 0.0483
                "aht_mape": "C17",    # Cell containing ~ 0.0620
                "workload": "B12",    # Cell containing ~ 373.3
                "req_agents": "B13",  # Cell containing ~ 390
                "fte": "B14",         # Cell containing ~ 542
                "efficiency": "B20"   # Cell containing ~ 0.914
            }
            
            def check_tol(val, expected, tol):
                try: return abs(float(val) - expected) <= tol
                except: return False

            try:
                wb = openpyxl.load_workbook(BytesIO(st.session_state.excel_data), data_only=True)
                
                # Task 1: Forecast Accuracy (20 pts)
                ws1 = wb["2. Forecast Accuracy"]
                if check_tol(ws1[CELL_MAP["vol_mape"]].value, 0.0483, 0.003): excel_score += 10
                if check_tol(ws1[CELL_MAP["aht_mape"]].value, 0.0620, 0.003): excel_score += 10

                # Task 2: Erlang & FTE (25 pts)
                ws2 = wb["3. Erlang & FTE"]
                if check_tol(ws2[CELL_MAP["workload"]].value, 373.3, 2.0): excel_score += 5
                if check_tol(ws2[CELL_MAP["req_agents"]].value, 390, 5): excel_score += 10
                if check_tol(ws2[CELL_MAP["fte"]].value, 542, 6): excel_score += 10

                # Task 3: Schedule Efficiency (Partial auto-grade)
                ws3 = wb["4. Schedule Efficiency"]
                if check_tol(ws3[CELL_MAP["efficiency"]].value, 0.914, 0.005): excel_score += 10
                # Assuming 10 points for manual validation of worst-3 intervals to avoid string-matching failure
                
                # Assigning placeholder max points for tasks 4 & 5 assuming formatting checks require manual review
                excel_score += 35

            except Exception as e:
                st.warning("Note: Excel file structure altered by candidate. Manual grading required.")

            st.session_state.final_excel = excel_score

            # --- 3. HARD GATE DECISION LOGIC & EMAIL ---
            passed_gate = excel_score >= 70
            gate_status = "✅ PASSED EXCEL GATE" if passed_gate else "🚨 FAILED EXCEL GATE (Auto-Decline)"

            try:
                msg = EmailMessage()
                msg['Subject'] = f"WFM Assessment: {st.session_state.name} | {gate_status}"
                msg['From'] = st.secrets["email_user"]
                msg['To'] = st.secrets["email_receiver"]
                
                email_body = f"""
                Candidate Assessment Completed!
                
                Name: {st.session_state.name}
                Email: {st.session_state.email}
                
                --- AUTO-SCORES ---
                Theory Score (Sec A): {st.session_state.final_theory}% ({t_score}/15)
                Excel Score (Sec D): {excel_score}% (100 Max)
                
                --- DECISION GATE ---
                Excel Gate Requirement: ≥ 70%
                Candidate Status: {gate_status}
                
                --- MANUAL GRADING REQUIRED (Open-Ended) ---
                B1. Forecasting: {st.session_state.b1}
                B2. Efficiency: {st.session_state.b2}
                B3. Escalation: {st.session_state.b3}
                B4. Spike Demand: {st.session_state.b4}
                B5. Coaching: {st.session_state.b5}
                
                *Attach candidate's Excel file and run final scores through your Scoring Calculator.*
                """
                msg.set_content(email_body)
                
                msg.add_attachment(
                    st.session_state.excel_data,
                    maintype='application',
                    subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    filename=f"{st.session_state.name.replace(' ', '_')}_WFM_Test.xlsx"
                )
                
                with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(st.secrets["email_user"], st.secrets["email_password"])
                    smtp.send_message(msg)
                    
                st.session_state.step = 4
                st.rerun()
                
            except Exception as e:
                st.error("Email delivery failed. Please check your Secret keys.")

# ==========================================
# STEP 4: SUCCESS / COMPLETION SCREEN
# ==========================================
elif st.session_state.step == 4:
    st.success("✅ Assessment Submitted Successfully!")
    st.balloons()
    
    col1, col2 = st.columns(2)
    with col1: st.metric(label="📚 Theory Score", value=f"{st.session_state.final_theory}%")
    with col2: st.metric(label="🧮 Excel Practical Score", value=f"{st.session_state.final_excel}%")
        
    st.markdown("---")
    st.write("Thank you for your time. The Boutiqaat recruitment team has received your submission and will review your analytical models and leadership scenarios shortly.")
