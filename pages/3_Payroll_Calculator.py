import sys, os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import pandas as pd
from datetime import datetime

try:
    from Utils import render_top_nav, read_any_table, download_button_for_df, compute_fulltime_payroll, compute_gig_worker_billing
except ModuleNotFoundError:
    from utils import render_top_nav, read_any_table, download_button_for_df, compute_fulltime_payroll, compute_gig_worker_billing

render_top_nav("Payroll Calculator")

st.title("🧾 Payroll Calculator")
st.caption("Compute monthly gross, deductions (PF/ESIC), and net pay for full-time staff and gig workers.")

st.markdown("#### Select Employee Type")
employee_type = st.radio("Employee Category", options=["Full-Time", "Gig/Contract"], horizontal=True)

if employee_type == "Full-Time":
    st.markdown("#### Upload Full-Time Employee Data")
    st.caption("Columns: Employee ID, Name, Basic, HRA (optional), Other Allowances (optional)")
    
    uploaded = st.file_uploader("Upload Full-Time Employee Data (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="fulltime")
    
    if uploaded:
        df = read_any_table(uploaded)
        st.session_state["fulltime_data"] = df
        st.success(f"Loaded {len(df)} employees.")
    
    if "fulltime_data" in st.session_state and not st.session_state["fulltime_data"].empty:
        df = st.session_state["fulltime_data"]
        
        st.markdown("#### Map Columns")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            basic_col = st.selectbox("Basic Column", options=df.columns)
        with col2:
            hra_col = st.selectbox("HRA Column (optional)", options=["-"] + list(df.columns))
        with col3:
            other_col = st.selectbox("Other Allowances (optional)", options=["-"] + list(df.columns))
        with col4:
            st.write("")
        
        st.markdown("#### Configure Statutory Rates & Thresholds")
        
        st.markdown("**PF Settings**")
        pf_col1, pf_col2, pf_col3 = st.columns(3)
        with pf_col1:
            pf_emp_pct = st.number_input("PF Employee %", value=12.0, min_value=0.0, max_value=50.0)
        with pf_col2:
            pf_emp_pct = st.number_input("PF Employer %", value=12.0, min_value=0.0, max_value=50.0)
        with pf_col3:
            pf_wage_cap = st.number_input("PF Wage Cap (₹)", value=15000.0, min_value=0.0)
        
        st.markdown("**ESIC Settings**")
        esic_col1, esic_col2, esic_col3 = st.columns(3)
        with esic_col1:
            esic_threshold = st.number_input("ESIC Threshold (₹)", value=21000.0, min_value=0.0)
        with esic_col2:
            esic_emp_pct = st.number_input("ESIC Employee %", value=0.75, min_value=0.0, max_value=5.0)
        with esic_col3:
            esic_emp_pct = st.number_input("ESIC Employer %", value=3.25, min_value=0.0, max_value=10.0)
        
        if st.button("Calculate Payroll"):
            try:
                result = compute_fulltime_payroll(
                    df,
                    basic_col=basic_col,
                    hra_col=None if hra_col == "-" else hra_col,
                    other_allow_col=None if other_col == "-" else other_col,
                    pf_wage_cap=pf_wage_cap,
                    pf_employee_pct=pf_emp_pct,
                    pf_employer_pct=pf_emp_pct,
                    esic_threshold=esic_threshold,
                    esic_employee_pct=esic_emp_pct,
                    esic_employer_pct=esic_emp_pct,
                )
                st.session_state["payroll_result"] = result
                st.success("✅ Payroll calculated successfully!")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
        
        if "payroll_result" in st.session_state:
            result = st.session_state["payroll_result"]
            
            st.markdown("#### Payroll Results")
            st.dataframe(result, use_container_width=True, height=400)
            
            # Summary
            st.markdown("#### Cost Summary")
            sum_col1, sum_col2, sum_col3, sum_col4 = st.columns(4)
            sum_col1.metric("Total Gross", f"₹{result['Gross Salary'].sum():,.0f}")
            sum_col2.metric("Total PF Cost", f"₹{(result['PF (Employee)'] + result['PF (Employer)']).sum():,.0f}")
            sum_col3.metric("Total ESIC Cost", f"₹{(result['ESIC (Employee)'] + result['ESIC (Employer)']).sum():,.0f}")
            sum_col4.metric("Total Net Pay", f"₹{result['Net Pay (Employee Take-home)'].sum():,.0f}")
            
            download_button_for_df(result, "⬇️ Download Payroll", f"fulltime_payroll_{datetime.now().strftime('%Y%m%d')}.xlsx")

else:  # Gig/Contract
    st.markdown("#### Upload Gig Worker Data")
    st.caption("Columns: Worker ID, Name, Payment Amount")
    
    uploaded = st.file_uploader("Upload Gig Worker Data (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="gig")
    
    if uploaded:
        df = read_any_table(uploaded)
        st.session_state["gig_data"] = df
        st.success(f"Loaded {len(df)} gig workers.")
    
    if "gig_data" in st.session_state and not st.session_state["gig_data"].empty:
        df = st.session_state["gig_data"]
        
        st.markdown("#### Map Columns")
        amount_col = st.selectbox("Payment Amount Column", options=df.columns)
        
        st.markdown("#### TDS Settings")
        tds_pct = st.number_input("TDS % (to be added to billing)", value=1.0, min_value=0.0, max_value=50.0)
        
        if st.button("Calculate Billing"):
            try:
                result = compute_gig_worker_billing(df, amount_col, tds_pct)
                st.session_state["gig_result"] = result
                st.success("✅ Billing calculated successfully!")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
        
        if "gig_result" in st.session_state:
            result = st.session_state["gig_result"]
            
            st.markdown("#### Billing Results")
            st.dataframe(result, use_container_width=True, height=400)
            
            # Summary
            st.markdown("#### Cost Summary")
            sum_col1, sum_col2, sum_col3 = st.columns(3)
            sum_col1.metric("Total Payment", f"₹{result[amount_col].sum():,.0f}")
            sum_col2.metric("Total TDS", f"₹{result['TDS Amount'].sum():,.0f}")
            sum_col3.metric("Total Billing", f"₹{result['Monthly Billing Amount'].sum():,.0f}")
            
            download_button_for_df(result, "⬇️ Download Billing", f"gig_billing_{datetime.now().strftime('%Y%m%d')}.xlsx")
