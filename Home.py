import os
import importlib.util

import streamlit as st
import pandas as pd


def _load_utils():
    """Loads utils.py by its exact file path (no sys.path / package
    resolution involved), so it works no matter how Streamlit was launched
    or what the current working directory is."""
    this_dir = os.path.dirname(os.path.abspath(__file__))
    for candidate in (
        os.path.join(this_dir, "utils.py"),
        os.path.join(this_dir, "Utils.py"),
    ):
        if os.path.exists(candidate):
            spec = importlib.util.spec_from_file_location("hr_utils", candidate)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise FileNotFoundError(
        "Could not find utils.py. Make sure it sits directly inside the "
        "same folder as Home.py."
    )


_u = _load_utils()
read_any_table = _u.read_any_table
download_button_for_df = _u.download_button_for_df
hide_default_sidebar_nav = _u.hide_default_sidebar_nav

st.set_page_config(
    page_title="HR Assistant",
    page_icon="🗂️",
    layout="wide",
)

hide_default_sidebar_nav()

st.title("🗂️ HR Assistant")
st.caption("Branch HR operations toolkit — employee database, long-absence tracking, "
           "incentive validation, and payroll (PF/ESIC/TDS) calculations.")

st.markdown("#### Navigate")
nav1, nav2, nav3, nav4 = st.columns(4)

with nav1:
    with st.container(border=True):
        st.markdown("##### 📅 Long Absence Tracker")
        st.caption("Upload attendance muster, auto-count `A` days, bucket employees into risk categories.")
        st.page_link("pages/1_Long_Absence_Tracker.py", label="Open →", icon="📅")

with nav2:
    with st.container(border=True):
        st.markdown("##### 💰 Incentive Validator")
        st.caption("Bulk-upload branch incentive data, validate & compute incentives.")
        st.page_link("pages/2_Incentive_Validator.py", label="Open →", icon="💰")

with nav3:
    with st.container(border=True):
        st.markdown("##### 🧾 Payroll Calculator")
        st.caption("Compute monthly gross for full-time staff (PF/ESIC) and gig workers (TDS billing).")
        st.page_link("pages/3_Payroll_Calculator.py", label="Open →", icon="🧾")

with nav4:
    with st.container(border=True):
        st.markdown("##### 🏦 Gratuity & Advance")
        st.caption("Statutory gratuity calculator, plus ERP/OIS-based paysheet lookup for salary advances.")
        st.page_link("pages/4_Gratuity_and_Advance.py", label="Open →", icon="🏦")

st.divider()
st.subheader("📋 Employee Database")
st.caption("Upload your employee master sheet. The **first row is always treated as the column header**.")

uploaded = st.file_uploader("Upload Employee Database (.xlsx or .csv)", type=["xlsx", "xls", "csv"])

if uploaded:
    df = read_any_table(uploaded)
    st.session_state["employee_db"] = df
    st.success(f"Loaded {len(df)} employee records with {len(df.columns)} columns.")

if "employee_db" in st.session_state and not st.session_state["employee_db"].empty:
    df = st.session_state["employee_db"]

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

    st.dataframe(filtered, use_container_width=True, height=420)
    download_button_for_df(filtered, "⬇️ Download filtered view", "employee_database_filtered.xlsx")

    if branch_col:
        st.markdown("#### Headcount by Branch")
        st.bar_chart(df[branch_col].value_counts())
else:
    st.info("Upload an employee database sheet to get started. This data is also used "
            "as a reference on other pages during this session.")
