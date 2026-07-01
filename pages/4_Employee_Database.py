import sys, os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import streamlit as st
import pandas as pd

try:
    from Utils import render_top_nav, read_any_table, download_button_for_df
except ModuleNotFoundError:
    from utils import render_top_nav, read_any_table, download_button_for_df

render_top_nav("Employee Database")

st.title("👥 Employee Database")
st.caption("Manage employee master data, apply filters, and download reports.")

st.markdown("#### Upload Employee Master Sheet")
st.caption("The first row is always treated as the column header.")

uploaded = st.file_uploader("Upload Employee Database (.xlsx or .csv)", type=["xlsx", "xls", "csv"])

if uploaded:
    df = read_any_table(uploaded)
    st.session_state["employee_db"] = df
    st.success(f"Loaded {len(df)} employee records with {len(df.columns)} columns.")

if "employee_db" in st.session_state and not st.session_state["employee_db"].empty:
    df = st.session_state["employee_db"]
    
    st.markdown("#### Overview")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Employees", len(df))
    
    branch_col = next((c for c in df.columns if "branch" in c.lower()), None)
    status_col = next((c for c in df.columns if "status" in c.lower() or "type" in c.lower()), None)
    
    if branch_col:
        col2.metric("Branches", df[branch_col].nunique())
    if status_col:
        col3.metric("Categories (Status/Type)", df[status_col].nunique())
    
    st.markdown("#### Filter & Search")
    fcol1, fcol2 = st.columns([2, 1])
    with fcol1:
        search_term = st.text_input("Search across all columns", "")
    with fcol2:
        filter_col = st.selectbox("Filter by column (optional)", ["(none)"] + list(df.columns))
    
    filtered = df.copy()
    if filter_col != "(none)":
        options = ["(all)"] + sorted(filtered[filter_col].dropna().astype(str).unique().tolist())
        chosen = st.selectbox(f"Value for '{filter_col}'", options)
        if chosen != "(all)":
            filtered = filtered[filtered[filter_col].astype(str) == chosen]
    
    if search_term:
        mask = filtered.apply(
            lambda row: row.astype(str).str.contains(search_term, case=False, na=False).any(),
            axis=1,
        )
        filtered = filtered[mask]
    
    st.markdown("#### Results")
    st.dataframe(filtered, use_container_width=True, height=400)
    download_button_for_df(filtered, "⬇️ Download filtered view", "employee_database_filtered.xlsx")
    
    if branch_col:
        st.markdown("#### Headcount by Branch")
        st.bar_chart(df[branch_col].value_counts())
else:
    st.info("Upload an employee database sheet to get started.")
