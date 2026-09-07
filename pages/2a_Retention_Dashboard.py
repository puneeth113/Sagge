import os
import importlib.util

import streamlit as st
import pandas as pd
from datetime import datetime


def _load_utils():
    """Loads utils.py by its exact file path (not via sys.path / package
    resolution), so it works regardless of how Streamlit was launched, the
    current working directory, or filename case (utils.py vs Utils.py)."""
    this_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(this_dir)
    candidates = [
        os.path.join(this_dir, "utils.py"),
        os.path.join(this_dir, "Utils.py"),
        os.path.join(root_dir, "utils.py"),
        os.path.join(root_dir, "Utils.py"),
    ]
    for path in candidates:
        if os.path.exists(path):
            spec = importlib.util.spec_from_file_location("hr_utils", path)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise FileNotFoundError(
        "Could not find utils.py. Make sure it sits directly inside the "
        "app's root folder (one level above 'pages/')."
    )


_u = _load_utils()
render_top_nav = _u.render_top_nav
download_button_for_df = _u.download_button_for_df
to_excel_bytes = _u.to_excel_bytes
render_clear_data_button = _u.render_clear_data_button

st.set_page_config(page_title="Retention Fund Dashboard", page_icon="📊", layout="wide")
render_top_nav("Retention Fund Tracker")

st.title("📊 Retention Fund Dashboard")
st.caption(
    "Statistical analysis and KPI tracking for retention fund deductions. "
    "View comprehensive analytics, deduction trends, and generate detailed reports."
)

with st.expander("🔒 Data handling on this page", expanded=False):
    st.markdown(
        "- Uploaded files are processed **only in memory** for this browser session — nothing is written "
        "to disk, logged, or sent to any external service.\n"
        "- Downloaded reports are automatically sanitized against Excel/CSV formula-injection payloads.\n"
        "- Use the button below to explicitly wipe all cached data from this session once you're done."
    )
    render_clear_data_button()


# ================================================================
# DASHBOARD ANALYTICS
# ================================================================

st.markdown("#### Dashboard Overview")

if "retention_result" in st.session_state and not st.session_state["retention_result"].empty:
    result = st.session_state["retention_result"]
    
    # Key Metrics
    st.markdown("### Key Performance Indicators (KPIs)")
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    total_employees = len(result)
    deduction_employees = (result["Deduction Applicable"] == True).sum()
    no_deduction_employees = (result["Deduction Applicable"] == False).sum()
    total_deduction = result["Deduction Amount"].sum()
    pending_release = (result["Status"] == "Pending Release").sum()
    completed_release = (result["Status"] == "No Deduction").sum()
    
    col1.metric("Total Employees", total_employees)
    col2.metric("With Deduction", deduction_employees)
    col3.metric("Without Deduction", no_deduction_employees)
    col4.metric("Pending Release", pending_release)
    col5.metric("Completed Release", completed_release)
    col6.metric("Total Deduction Amount", f"₹{total_deduction:,.0f}")
    
    # Average Deduction
    if deduction_employees > 0:
        avg_deduction = result[result["Deduction Amount"] > 0]["Deduction Amount"].mean()
        st.metric("Average Deduction per Employee (with deduction)", f"₹{avg_deduction:,.0f}")
    
    # Status Distribution
    st.markdown("### Status Distribution")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### Deduction Status Breakdown")
        status_counts = result["Status"].value_counts()
        st.bar_chart(status_counts)
    
    with col2:
        st.markdown("#### Deduction Applicable")
        applicable_counts = result["Deduction Applicable"].value_counts().map({True: "Applicable", False: "Not Applicable"})
        st.bar_chart(applicable_counts)
    
    # Company Code Analysis
    st.markdown("### Analysis by Company Code")
    cc_col = "Company Code"
    if cc_col in result.columns:
        cc_summary = result.groupby(cc_col).agg({
            "Deduction Amount": ["sum", "mean", "count"],
            "Status": lambda x: (x == "Pending Release").sum(),
        }).round(2)
        cc_summary.columns = ["Total Deduction", "Average Deduction", "Employee Count", "Pending Count"]
        cc_summary = cc_summary.sort_values("Total Deduction", ascending=False)
        
        st.dataframe(cc_summary, use_container_width=True)
        
        # Chart: Total Deduction by Company Code
        st.markdown("#### Total Deduction by Company Code")
        cc_chart_data = result.groupby(cc_col)["Deduction Amount"].sum().sort_values(ascending=False)
        st.bar_chart(cc_chart_data)
        
        # Chart: Pending vs Completed by Company Code
        st.markdown("#### Pending vs Completed by Company Code")
        cc_status_data = result.groupby(cc_col)["Status"].apply(lambda x: (x == "Pending Release").sum()).sort_values(ascending=False)
        st.bar_chart(cc_status_data)
    
    # Branch Analysis
    st.markdown("### Analysis by Branch")
    branch_col = "Branch"
    if branch_col in result.columns:
        branch_summary = result.groupby(branch_col).agg({
            "Deduction Amount": ["sum", "mean", "count"],
            "Status": lambda x: (x == "Pending Release").sum(),
        }).round(2)
        branch_summary.columns = ["Total Deduction", "Average Deduction", "Employee Count", "Pending Count"]
        branch_summary = branch_summary.sort_values("Total Deduction", ascending=False)
        
        st.dataframe(branch_summary, use_container_width=True)
        
        # Chart: Total Deduction by Branch
        st.markdown("#### Total Deduction by Branch")
        branch_chart_data = result.groupby(branch_col)["Deduction Amount"].sum().sort_values(ascending=False)
        st.bar_chart(branch_chart_data)
        
        # Chart: Pending vs Completed by Branch
        st.markdown("#### Pending vs Completed by Branch")
        branch_status_data = result.groupby(branch_col)["Status"].apply(lambda x: (x == "Pending Release").sum()).sort_values(ascending=False)
        st.bar_chart(branch_status_data)
    
    # Salary Range Analysis
    st.markdown("### Deduction by Salary Range")
    if "Earned Salary" in result.columns:
        result_copy = result.copy()
        result_copy["Salary Range"] = pd.cut(
            result_copy["Earned Salary"],
            bins=[0, 20000, 30000, 40000, 50000, float('inf')],
            labels=["<20K", "20K-30K", "30K-40K", "40K-50K", ">50K"]
        )
        
        salary_summary = result_copy.groupby("Salary Range").agg({
            "Deduction Amount": ["sum", "mean", "count"],
            "Status": lambda x: (x == "Pending Release").sum(),
        }).round(2)
        salary_summary.columns = ["Total Deduction", "Average Deduction", "Employee Count", "Pending Count"]
        
        st.dataframe(salary_summary, use_container_width=True)
        
        # Chart
        st.markdown("#### Distribution Across Salary Ranges")
        salary_chart = result_copy.groupby("Salary Range")["Deduction Amount"].sum()
        st.bar_chart(salary_chart)
    
    # Expected Release Date Analysis
    st.markdown("### Expected Release Date Analysis")
    if "Expected Release Date" in result.columns:
        # Extract year-month for grouping
        result_copy = result.copy()
        result_copy["Release Year-Month"] = pd.to_datetime(result_copy["Expected Release Date"]).dt.to_period('M')
        
        release_summary = result_copy[result_copy["Status"] == "Pending Release"].groupby("Release Year-Month").agg({
            "Deduction Amount": ["sum", "count"],
        }).round(2)
        
        if not release_summary.empty:
            release_summary.columns = ["Total Deduction", "Employee Count"]
            st.dataframe(release_summary, use_container_width=True)
            
            # Chart: Expected Releases by Month
            st.markdown("#### Expected Releases by Month")
            release_chart = result_copy[result_copy["Status"] == "Pending Release"].groupby("Release Year-Month")["Deduction Amount"].sum()
            st.bar_chart(release_chart)
    
    # Top Deduction Recipients
    st.markdown("### Top 10 Employees by Deduction Amount")
    erp_col = "ERP" if "ERP" in result.columns else result.columns[0]
    name_col = "Name" if "Name" in result.columns else result.columns[1]
    cc_col = "Company Code" if "Company Code" in result.columns else "Company Code"
    branch_col = "Branch" if "Branch" in result.columns else "Branch"
    salary_col = "Earned Salary" if "Earned Salary" in result.columns else "Salary"
    
    top_deductions = result.nlargest(10, "Deduction Amount")[[
        erp_col, name_col, cc_col, branch_col, salary_col, "Deduction Amount", "Status"
    ]].copy()
    
    if "Expected Release Date Formatted" in result.columns:
        top_deductions["Expected Release Date"] = result.nlargest(10, "Deduction Amount")["Expected Release Date Formatted"].values
    
    st.dataframe(top_deductions, use_container_width=True)
    
    # Deduction Timeline
    st.markdown("### Deduction Completion Timeline")
    if "Expected Release Date" in result.columns:
        pending_by_release = result[result["Status"] == "Pending Release"].copy()
        if not pending_by_release.empty:
            pending_by_release["Release Date"] = pd.to_datetime(pending_by_release["Expected Release Date"]).dt.date
            timeline = pending_by_release.groupby("Release Date").agg({
                "Deduction Amount": ["sum", "count"],
            }).round(2)
            timeline.columns = ["Total Deduction Amount", "Number of Employees"]
            timeline = timeline.sort_index()
            st.dataframe(timeline, use_container_width=True)
    
    # Summary Report
    st.markdown("### Summary Report")
    summary_data = {
        "Metric": [
            "Total Employees",
            "Employees with Active Deduction",
            "Employees without Deduction",
            "Pending Release Count",
            "Completed Release Count",
            "Total Deduction Amount",
            "Average Deduction (Active)",
            "Report Generated On"
        ],
        "Value": [
            str(total_employees),
            str(deduction_employees),
            str(no_deduction_employees),
            str(pending_release),
            str(completed_release),
            f"₹{total_deduction:,.2f}",
            f"₹{result[result['Deduction Amount'] > 0]['Deduction Amount'].mean():,.2f}" if deduction_employees > 0 else "N/A",
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ]
    }
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, use_container_width=True)
    
    # Download Full Dashboard Report
    st.markdown("### Download Reports")
    
    # Full detailed report
    download_button_for_df(
        result,
        "⬇️ Download Full Dashboard Report",
        f"retention_dashboard_full_{datetime.now().strftime('%Y%m%d')}.xlsx",
        key="dl_dashboard_full",
    )
    
    # Pending releases report
    pending = result[result["Status"] == "Pending Release"].copy()
    if not pending.empty:
        download_button_for_df(
            pending,
            "⬇️ Download Pending Releases Report",
            f"retention_pending_releases_{datetime.now().strftime('%Y%m%d')}.xlsx",
            key="dl_pending_releases",
        )
    
    # Summary report
    download_button_for_df(
        summary_df,
        "⬇️ Download Summary Report",
        f"retention_summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
        key="dl_summary",
    )
    
else:
    st.info("📌 No retention fund data available. Please calculate retention fund deductions from the Computation page first.")
