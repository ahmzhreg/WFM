import streamlit as st
import pandas as pd

st.set_page_config(page_title="WFM Assessment Portal", page_icon="📊", layout="centered")

if 'score' not in st.session_state:
    st.session_state.score = 0
if 'submitted' not in st.session_state:
    st.session_state.submitted = False

st.title("📊 WFM Team Leader Assessment")
st.markdown("""
Welcome to the technical assessment. 
* **Time limit:** 90 minutes
* **Instructions:** Complete all tabs below and hit submit on the final page.
---
""")

tab1, tab2, tab3 = st.tabs(["📚 Section 1: Theory", "🧮 Section 2: Data Lab", "📝 Section 3: Operations Floor"])

with tab1:
    st.header("Section 1: WFM Theory")
    q1 = st.radio("1. You observe Agent Schedule Adherence is 94%, but Occupancy has spiked to 92%. What is the most immediate risk?",
        options=["A) High schedule shrinkage and increased non-productive time.", "B) Increased agent fatigue, potential burnout, and longer handle times.", "C) A drop in Schedule Adherence.", "D) Overstaffing leading to budget overruns."], index=None)
    q2 = st.radio("2. Which forecasting methodology is most appropriate for a promotional event influenced by marketing spend?",
        options=["A) Simple Moving Average.", "B) Naive Forecasting.", "C) Multiple Regression Analysis with scenario modeling.", "D) Erlang C Volume Distribution."], index=None)
    q3 = st.radio("3. According to standard KPIs, what is the maximum acceptable intraday variance for AHT?",
        options=["A) ≤ ±5%", "B) ≤ ±10%", "C) ≤ ±15%", "D) ≤ ±2%"], index=None)

with tab2:
    st.header("Section 2: Capacity Planning Lab")
    st.markdown("**Scenario:** 450 calls | 240s AHT | SLA: 80/20 | 30% Shrinkage")
    erlang_workload = st.number_input("1. What is the raw workload in Erlang hours?", min_value=0, step=1)
    scheduled_hc = st.number_input("2. If raw requirement is 64 agents, what is the scheduled headcount to cover 30% shrinkage?", min_value=0, step=1)

with tab3:
    st.header("Section 3: Operations Floor")
    crisis_response = st.text_area("Intraday Crisis (2:00 PM, vol +40%, Abandonment 12%): Outline immediate corrective actions.", height=150)
    fatigue_response = st.text_area("Fatigue Management: How do you balance cost-efficiency with well-being during Ramadan?", height=150)
    st.divider()
    
    if st.button("Submit Assessment", type="primary"):
        st.session_state.submitted = True
        score = 0
        if q1 and q1.startswith("B"): score += 10
        if q2 and q2.startswith("C"): score += 10
        if q3 and q3.startswith("A"): score += 10
        if erlang_workload == 60: score += 20
        if scheduled_hc == 92: score += 20
        
        st.success("Submitted Successfully!")
        st.write("### Auto-Scored Results")
        st.metric(label="Quantitative Score (Max 70 pts)", value=f"{score} / 70")