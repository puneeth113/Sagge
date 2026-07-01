import sys, os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import pandas as pd
from datetime import datetime

try:
    from Utils import render_top_nav, read_any_table, download_button_for_df, count_absences, categorize_absences, DEFAULT_ABSENCE_THRESHOLDS, sample_attendance_muster
except ModuleNotFoundError:
    from utils import render_top_nav, read_any_table, download_button_for_df, count_absences, categorize_absences, DEFAULT_ABSENCE_THRESHOLDS, sample_attendance_muster

render_top_nav("Long Absence Tracker")

st.title("📅 Long Absence Tracker")
st.caption("Upload an attendance muster, auto-count absence days, and categorize employees into risk buckets.")

st.markdown("#### Upload Attendance Muster")
st.caption("Expected format: First column(s) for employee ID/name, remaining columns are day-wise attendance marks.")

uploaded = st.file_uploader("Upload Attendance Muster (.xlsx or .csv)", type=["xlsx", "xls", "csv"])

if uploaded:
    df = read_any_table(uploaded)
    st.session_state["attendance_muster"] = df
    st.success(f"Loaded muster with {len(df)} employees and {len(df.columns)} columns.")

if "attendance_muster" in st.session_state and not st.session_state["attendance_muster"].empty:
    df = st.session_state["attendance_muster"]
    
    st.markdown("#### Configure ID Columns")
    st.caption("Select which columns identify each employee (usually Employee ID, Name, etc.)")
    id_cols = st.multiselect(
        "ID columns (employee identifiers)",
        options=list(df.columns),
        default=list(df.columns[:2]) if len(df.columns) >= 2 else list(df.columns[:1])
    )
    
    if id_cols:
        st.markdown("#### Configure Absence Marker")
        absence_marker = st.text_input("Absence marker (usually 'A')", value="A", max_chars=3)
        
        st.markdown("#### Adjust Risk Thresholds (Optional)")
        col1, col2, col3 = st.columns(3)
        
        thresholds = []
        with col1:
            informed_min = st.number_input("Informed Leave: Min days", value=3, min_value=1)
            informed_max = st.number_input("Informed Leave: Max days", value=4, min_value=1)
        with col2:
            check_min = st.number_input("To Be Checked: Min days", value=5, min_value=1)
            check_max = st.number_input("To Be Checked: Max days", value=20, min_value=1)
        with col3:
            exit_min = st.number_input("Exit Case: Min days", value=21, min_value=1)
        
        thresholds = [
            (f"Informed Leave ({informed_min}-{informed_max} days)", informed_min, informed_max),
            (f"To Be Checked ({check_min}-{check_max} days)", check_min, check_max),
            (f"Probable Exit Case (>{exit_min} days)", exit_min, None),
        ]
        
        if st.button("Analyze Attendance"):
            # Count absences
            result = count_absences(df, id_cols, absence_marker)
            result = categorize_absences(result, thresholds)
            
            st.session_state["absence_result"] = result
            st.success("Analysis complete!")
        
        if "absence_result" in st.session_state:
            result = st.session_state["absence_result"]
            
            st.markdown("#### Results")
            st.dataframe(result, use_container_width=True, height=400)
            
            # Summary metrics
            st.markdown("#### Summary")
            mcol1, mcol2, mcol3, mcol4 = st.columns(4)
            mcol1.metric("Total Employees", len(result))
            mcol2.metric("Avg Absent Days", f"{result['Total_Absent_Days'].mean():.1f}")
            mcol3.metric("Max Absent Days", result['Total_Absent_Days'].max())
            mcol4.metric("High Risk Count", len(result[result['Absence_Category'].str.contains("Exit", na=False)]))
            
            # Category breakdown
            st.markdown("#### Category Breakdown")
            st.bar_chart(result['Absence_Category'].value_counts())
            
            # Download button
            download_button_for_df(result, "⬇️ Download Analysis Results", f"absence_analysis_{datetime.now().strftime('%Y%m%d')}.xlsx")
else:
    # Show sample
    if st.button("Load Sample Data"):
        sample_df = sample_attendance_muster()
        st.session_state["attendance_muster"] = sample_df
        st.rerun()
