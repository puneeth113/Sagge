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
        col1, col2, col3 = st.columns(3)
        with col1:
            basic_col = st.selectbox("Basic Column", options=df.columns, key="basic_col")
        with col2:
            hra_col = st.selectbox("HRA Column (optional)", options=["-"] + list(df.columns), key="hra_col")
        with col3:
            other_col = st.selectbox("Other Allowances (optional)", options=["-"] + list(df.columns), key="other_col")
        
        st.markdown("#### Configure Statutory Rates & Thresholds")
        
        st.markdown("**PF Settings**")
        pf_col1, pf_col2, pf_col3 = st.columns(3)
        with pf_col1:
            pf_emp_pct = st.number_input("PF Employee %", value=12.0, min_value=0.0, max_value=50.0, key="pf_emp")
        with pf_col2:
            pf_employer_pct = st.number_input("PF Employer %", value=12.0, min_value=0.0, max_value=50.0, key="pf_emp_rate")
        with pf_col3:
            pf_wage_cap = st.number_input("PF Wage Cap (₹)", value=15000.0, min_value=0.0, key="pf_cap")
        
        st.markdown("**ESIC Settings**")
        esic_col1, esic_col2, esic_col3 = st.columns(3)
        with esic_col1:
            esic_threshold = st.number_input("ESIC Threshold (₹)", value=21000.0, min_value=0.0, key="esic_thresh")
        with esic_col2:
            esic_emp_pct = st.number_input("ESIC Employee %", value=0.75, min_value=0.0, max_value=5.0, key="esic_emp")
        with esic_col3:
            esic_employer_pct = st.number_input("ESIC Employer %", value=3.25, min_value=0.0, max_value=10.0, key="esic_emp_rate")
        
        st.markdown("**Apply PF Wage Cap**")
        apply_pf_cap = st.checkbox("Apply ₹15,000 wage cap for PF", value=True)
        
        if st.button("Calculate Payroll", key="calc_fulltime"):
            try:
                result = compute_fulltime_payroll(
                    df,
                    basic_col=basic_col,
                    hra_col=None if hra_col == "-" else hra_col,
                    other_allow_col=None if other_col == "-" else other_col,
                    pf_wage_cap=pf_wage_cap,
                    apply_pf_cap=apply_pf_cap,
                    pf_employee_pct=pf_emp_pct,
                    pf_employer_pct=pf_employer_pct,
                    esic_threshold=esic_threshold,
                    esic_employee_pct=esic_emp_pct,
                    esic_employer_pct=esic_employer_pct,
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
            st.markdown("#### Employee Deductions Summary")
            sum_col1, sum_col2, sum_col3, sum_col4 = st.columns(4)
            sum_col1.metric("Total Gross", f"₹{result['Gross Salary'].sum():,.0f}")
            sum_col2.metric("Total Employee PF", f"₹{result['PF (Employee)'].sum():,.0f}")
            sum_col3.metric("Total Employee ESIC", f"₹{result['ESIC (Employee)'].sum():,.0f}")
            sum_col4.metric("Total Net Pay", f"₹{result['Net Pay (Employee Take-home)'].sum():,.0f}")
            
            st.markdown("#### Employer Cost Summary")
            emp_col1, emp_col2, emp_col3 = st.columns(3)
            emp_col1.metric("Total Employer PF", f"₹{result['PF (Employer)'].sum():,.0f}")
            emp_col2.metric("Total Employer ESIC", f"₹{result['ESIC (Employer)'].sum():,.0f}")
            emp_col3.metric("Total Employer Cost", f"₹{result['Total Employer Cost (CTC add-on)'].sum():,.0f}")
            
            st.markdown("#### Cost Breakdown")
            cost_chart_data = pd.DataFrame({
                "Employee": ["Gross Salary", "PF (Employee)", "ESIC (Employee)", "Net Pay"],
                "Amount": [
                    result['Gross Salary'].sum(),
                    result['PF (Employee)'].sum(),
                    result['ESIC (Employee)'].sum(),
                    result['Net Pay (Employee Take-home)'].sum(),
                ]
            })
            st.bar_chart(cost_chart_data.set_index('Employee'))
            
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
        amount_col = st.selectbox("Payment Amount Column", options=df.columns, key="amount_col")
        
        st.markdown("#### TDS Settings")
        col1, col2 = st.columns(2)
        with col1:
            tds_pct = st.number_input("TDS % (to be added to billing)", value=1.0, min_value=0.0, max_value=50.0, key="tds_pct")
        with col2:
            st.info(f"💡 1% TDS adds: ₹{1000 * 0.01:.0f} per ₹1,000 payment")
        
        if st.button("Calculate Billing", key="calc_gig"):
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
            st.markdown("#### Billing Summary")
            sum_col1, sum_col2, sum_col3 = st.columns(3)
            sum_col1.metric("Total Payment", f"₹{result[amount_col].sum():,.0f}")
            sum_col2.metric("Total TDS", f"₹{result['TDS Amount'].sum():,.0f}")
            sum_col3.metric("Total Billing", f"₹{result['Monthly Billing Amount'].sum():,.0f}")
            
            # Visualization
            st.markdown("#### Billing Breakdown")
            billing_chart = pd.DataFrame({
                "Category": ["Payment", "TDS"],
                "Amount": [
                    result[amount_col].sum(),
                    result['TDS Amount'].sum()
                ]
            })
            st.bar_chart(billing_chart.set_index('Category'))
            
            download_button_for_df(result, "⬇️ Download Billing", f"gig_billing_{datetime.now().strftime('%Y%m%d')}.xlsx")
