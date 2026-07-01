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
st.caption("Calculate component-wise paysheet and manage bulk employee payroll.")

st.markdown("#### Upload Payroll Data")
st.caption("Columns: Employee ID, Name, Basic, HRA, Other Allowances, etc.")

uploaded = st.file_uploader("Upload Payroll Data (.xlsx or .csv)", type=["xlsx", "xls", "csv"])

if uploaded:
    df = read_any_table(uploaded)
    st.session_state["paysheet_data"] = df
    st.success(f"Loaded {len(df)} employee records.")

if "paysheet_data" in st.session_state and not st.session_state["paysheet_data"].empty:
    df = st.session_state["paysheet_data"]
    
    st.markdown("#### Data Preview")
    st.dataframe(df, use_container_width=True, height=300)
    
    st.markdown("#### Generate Paysheet")
    if st.button("Generate Component-wise Paysheet"):
        # Calculate gross
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        
        result = df.copy()
        result['Gross'] = df[numeric_cols].sum(axis=1) if numeric_cols else 0
        
        # Add statutory deductions (sample)
        result['PF (12%)'] = result['Gross'] * 0.12
        result['ESIC (0.75%)'] = result['Gross'] * 0.0075
        result['Net Pay'] = result['Gross'] - result['PF (12%)'] - result['ESIC (0.75%)'] 
        
        st.session_state["paysheet_result"] = result
        st.success("✅ Paysheet generated!")
    
    if "paysheet_result" in st.session_state:
        result = st.session_state["paysheet_result"]
        
        st.markdown("#### Paysheet")
        st.dataframe(result, use_container_width=True, height=400)
        
        # Summary
        st.markdown("#### Summary")
        sum_col1, sum_col2, sum_col3, sum_col4 = st.columns(4)
        sum_col1.metric("Total Gross", f"₹{result['Gross'].sum():,.0f}")
        sum_col2.metric("Total PF", f"₹{result['PF (12%)'].sum():,.0f}")
        sum_col3.metric("Total ESIC", f"₹{result['ESIC (0.75%)'].sum():,.0f}")
        sum_col4.metric("Total Net", f"₹{result['Net Pay'].sum():,.0f}")
        
        download_button_for_df(result, "⬇️ Download Paysheet", f"paysheet_{datetime.now().strftime('%Y%m%d')}.xlsx")
else:
    st.info("Upload payroll data to get started.")
