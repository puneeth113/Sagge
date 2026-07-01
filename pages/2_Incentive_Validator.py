import sys, os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import pandas as pd
from datetime import datetime

try:
    from Utils import render_top_nav, read_any_table, download_button_for_df, sample_incentive_template, compute_incentives, INCENTIVE_TEMPLATE_COLUMNS
except ModuleNotFoundError:
    from utils import render_top_nav, read_any_table, download_button_for_df, sample_incentive_template, compute_incentives, INCENTIVE_TEMPLATE_COLUMNS

render_top_nav("Incentive Validator")

st.title("💰 Incentive Validator")
st.caption("Bulk-upload branch incentive data, validate, and compute incentives.")

st.markdown("#### Download Sample Template")
sample_df = sample_incentive_template()
if st.button("📋 Download Template"):
    from Utils import to_excel_bytes
    try:
        from utils import to_excel_bytes
    except:
        pass
    
    buffer = to_excel_bytes({"Incentive Data": sample_df})
    st.download_button(
        label="Download Incentive Template",
        data=buffer,
        file_name="incentive_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("#### Upload Incentive Data")
st.caption(f"Required columns: {', '.join(INCENTIVE_TEMPLATE_COLUMNS)}")
uploaded = st.file_uploader("Upload Branch Incentive Data (.xlsx or .csv)", type=["xlsx", "xls", "csv"])

if uploaded:
    df = read_any_table(uploaded)
    st.session_state["incentive_data"] = df
    st.success(f"Loaded {len(df)} branch records.")

if "incentive_data" in st.session_state and not st.session_state["incentive_data"].empty:
    df = st.session_state["incentive_data"]
    
    # Validate columns
    missing_cols = [col for col in INCENTIVE_TEMPLATE_COLUMNS if col not in df.columns]
    if missing_cols:
        st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
    else:
        st.markdown("#### Configure Rates")
        col1, col2 = st.columns(2)
        with col1:
            rate_washroom = st.number_input("Rate per Washroom Maid (₹)", value=500.0, step=10.0, min_value=0.0)
        with col2:
            rate_bus = st.number_input("Rate per Bus Maid (₹)", value=500.0, step=10.0, min_value=0.0)
        
        st.markdown("#### Current Data")
        st.dataframe(df, use_container_width=True, height=250)
        
        if st.button("Compute Incentives"):
            try:
                result = compute_incentives(df, rate_washroom, rate_bus)
                st.session_state["incentive_result"] = result
                st.success("✅ Incentives computed successfully!")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
        
        if "incentive_result" in st.session_state:
            result = st.session_state["incentive_result"]
            
            st.markdown("#### Results")
            st.dataframe(result, use_container_width=True, height=400)
            
            # Summary metrics
            st.markdown("#### Summary")
            mcol1, mcol2, mcol3 = st.columns(3)
            mcol1.metric("Total Branches", len(result))
            mcol2.metric("Total Incentive", f"₹{result['Total Incentive'].sum():,.0f}")
            mcol3.metric("Avg per Branch", f"₹{result['Total Incentive'].mean():,.0f}")
            
            # Charts
            st.markdown("#### Visualizations")
            chart_col1, chart_col2 = st.columns(2)
            with chart_col1:
                st.bar_chart(result.set_index('Branch Name')['Total Incentive'])
            with chart_col2:
                st.bar_chart(result.set_index('Branch Name')[['Calculated Washroom Incentive', 'Calculated Bus Incentive']])
            
            # Download
            download_button_for_df(result, "⬇️ Download Incentive Results", f"incentive_validation_{datetime.now().strftime('%Y%m%d')}.xlsx")
else:
    st.info("Upload incentive data to get started.")
