import sys, os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import pandas as pd
from datetime import datetime

try:
    from Utils import render_top_nav, read_any_table, download_button_for_df
except ModuleNotFoundError:
    from utils import render_top_nav, read_any_table, download_button_for_df

render_top_nav("Paysheet Operations")

st.title("📋 Paysheet Operations")
st.caption("Calculate component-wise paysheet and manage bulk employee payroll with detailed deductions.")

st.markdown("#### Upload Payroll Data")
st.caption("Columns: Employee ID, Name, Basic, HRA, Other Allowances, CTC, etc.")

uploaded = st.file_uploader("Upload Payroll Data (.xlsx or .csv)", type=["xlsx", "xls", "csv"])

if uploaded:
    df = read_any_table(uploaded)
    st.session_state["paysheet_data"] = df
    st.success(f"Loaded {len(df)} employee records.")

if "paysheet_data" in st.session_state and not st.session_state["paysheet_data"].empty:
    df = st.session_state["paysheet_data"]
    
    st.markdown("#### Data Preview")
    st.dataframe(df, use_container_width=True, height=250)
    
    st.markdown("#### Configure Paysheet Calculation")
    
    # Column selection
    st.markdown("**Select Salary Components**")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        basic_col = st.selectbox("Basic Column", options=df.columns, key="basic")
    with col2:
        hra_col = st.selectbox("HRA Column (optional)", options=["-"] + list(df.columns), key="hra")
    with col3:
        other_col = st.selectbox("Other Allowances (optional)", options=["-"] + list(df.columns), key="other")
    with col4:
        st.write("")
    
    # Deduction configuration
    st.markdown("**Configure Deductions**")
    ded_col1, ded_col2, ded_col3 = st.columns(3)
    with ded_col1:
        pf_pct = st.number_input("PF % (Employee)", value=12.0, min_value=0.0, max_value=50.0, key="pf_pct")
    with ded_col2:
        esic_pct = st.number_input("ESIC % (Employee)", value=0.75, min_value=0.0, max_value=5.0, key="esic_pct")
    with ded_col3:
        esic_threshold = st.number_input("ESIC Threshold (₹)", value=21000.0, min_value=0.0, key="esic_threshold")
    
    st.markdown("**Optional Deductions**")
    opt_col1, opt_col2, opt_col3 = st.columns(3)
    with opt_col1:
        add_tds = st.checkbox("Add TDS Calculation", value=False)
        if add_tds:
            tds_pct = st.number_input("TDS %", value=5.0, min_value=0.0, max_value=30.0, key="tds_pct_paysheet")
        else:
            tds_pct = 0.0
    with opt_col2:
        add_pt = st.checkbox("Add Professional Tax", value=False)
        if add_pt:
            pt_amount = st.number_input("Professional Tax Amount (₹)", value=0.0, min_value=0.0, key="pt_amount")
        else:
            pt_amount = 0.0
    with opt_col3:
        add_other_ded = st.checkbox("Other Deductions", value=False)
        if add_other_ded:
            other_ded_col = st.selectbox("Other Deductions Column", options=["-"] + list(df.columns), key="other_ded_col")
        else:
            other_ded_col = "-"
    
    st.markdown("**Employer Contribution**")
    cont_col1, cont_col2 = st.columns(2)
    with cont_col1:
        pf_emp_pct = st.number_input("PF % (Employer)", value=12.0, min_value=0.0, max_value=50.0, key="pf_emp_pct")
    with cont_col2:
        esic_emp_pct = st.number_input("ESIC % (Employer)", value=3.25, min_value=0.0, max_value=10.0, key="esic_emp_pct")
    
    if st.button("Generate Component-wise Paysheet"):
        try:
            # Create result dataframe
            result = df.copy()
            
            # Calculate Gross
            result['Basic'] = result[basic_col]
            result['HRA'] = result[hra_col] if hra_col != "-" else 0.0
            result['Other Allowances'] = result[other_col] if other_col != "-" else 0.0
            result['Gross Salary'] = result['Basic'] + result['HRA'] + result['Other Allowances']
            
            # Calculate Employee Deductions
            result['PF (Employee)'] = result['Basic'] * pf_pct / 100
            
            # ESIC only if gross <= threshold
            result['ESIC Applicable'] = result['Gross Salary'] <= esic_threshold
            result['ESIC (Employee)'] = result.apply(
                lambda row: row['Gross Salary'] * esic_pct / 100 if row['ESIC Applicable'] else 0.0,
                axis=1
            )
            
            # TDS
            result['TDS'] = result['Gross Salary'] * tds_pct / 100 if add_tds else 0.0
            
            # Professional Tax
            result['Professional Tax'] = pt_amount if add_pt else 0.0
            
            # Other Deductions
            result['Other Deductions'] = result[other_ded_col] if other_ded_col != "-" else 0.0
            
            # Total Deductions
            result['Total Deductions'] = (
                result['PF (Employee)'] + 
                result['ESIC (Employee)'] + 
                result['TDS'] + 
                result['Professional Tax'] + 
                result['Other Deductions']
            )
            
            # Net Pay
            result['Net Pay'] = result['Gross Salary'] - result['Total Deductions']
            
            # Employer Contributions
            result['PF (Employer)'] = result['Basic'] * pf_emp_pct / 100
            result['ESIC (Employer)'] = result.apply(
                lambda row: row['Gross Salary'] * esic_emp_pct / 100 if row['ESIC Applicable'] else 0.0,
                axis=1
            )
            result['Total Employer Cost'] = result['PF (Employer)'] + result['ESIC (Employer)']
            
            st.session_state["paysheet_result"] = result
            st.success("✅ Paysheet generated successfully!")
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
    
    if "paysheet_result" in st.session_state:
        result = st.session_state["paysheet_result"]
        
        # Show relevant columns
        display_cols = [
            col for col in ['Basic', 'HRA', 'Other Allowances', 'Gross Salary', 
                           'PF (Employee)', 'ESIC (Employee)', 'TDS', 'Professional Tax', 
                           'Other Deductions', 'Total Deductions', 'Net Pay'] 
            if col in result.columns
        ]
        
        st.markdown("#### Paysheet")
        st.dataframe(result[display_cols], use_container_width=True, height=400)
        
        # Employee Summary
        st.markdown("#### Employee Summary")
        emp_col1, emp_col2, emp_col3, emp_col4, emp_col5 = st.columns(5)
        emp_col1.metric("Total Gross", f"₹{result['Gross Salary'].sum():,.0f}")
        emp_col2.metric("Total PF (E)", f"₹{result['PF (Employee)'].sum():,.0f}")
        emp_col3.metric("Total ESIC (E)", f"₹{result['ESIC (Employee)'].sum():,.0f}")
        emp_col4.metric("Total Deductions", f"₹{result['Total Deductions'].sum():,.0f}")
        emp_col5.metric("Total Net Pay", f"₹{result['Net Pay'].sum():,.0f}")
        
        # Employer Summary
        st.markdown("#### Employer Cost Summary")
        emp_sum1, emp_sum2, emp_sum3 = st.columns(3)
        emp_sum1.metric("Total PF (Employer)", f"₹{result['PF (Employer)'].sum():,.0f}")
        emp_sum2.metric("Total ESIC (Employer)", f"₹{result['ESIC (Employer)'].sum():,.0f}")
        emp_sum3.metric("Total Employer Cost", f"₹{result['Total Employer Cost'].sum():,.0f}")
        
        # Overall Cost
        st.markdown("#### Overall Cost Analysis")
        overall_col1, overall_col2, overall_col3 = st.columns(3)
        total_ctc = result['Gross Salary'].sum() + result['Total Employer Cost'].sum()
        overall_col1.metric("Total Payroll (Employees)", f"₹{result['Net Pay'].sum():,.0f}")
        overall_col2.metric("Employer Contribution", f"₹{result['Total Employer Cost'].sum():,.0f}")
        overall_col3.metric("Total CTC Cost", f"₹{total_ctc:,.0f}")
        
        # Charts
        st.markdown("#### Deduction Breakdown")
        chart_cols = ['PF (Employee)', 'ESIC (Employee)', 'TDS', 'Professional Tax', 'Other Deductions']
        chart_data = pd.DataFrame({
            'Deduction': chart_cols,
            'Amount': [result[col].sum() if col in result.columns else 0 for col in chart_cols]
        })
        st.bar_chart(chart_data.set_index('Deduction'))
        
        # Full data download
        st.markdown("#### Download")
        download_cols = [
            col for col in result.columns 
            if col not in ['ESIC Applicable']
        ]
        download_button_for_df(result[download_cols], "⬇️ Download Complete Paysheet", f"paysheet_{datetime.now().strftime('%Y%m%d')}.xlsx")
else:
    st.info("Upload payroll data to get started.")
