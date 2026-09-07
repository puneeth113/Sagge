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
read_any_table = _u.read_any_table
download_button_for_df = _u.download_button_for_df
to_excel_bytes = _u.to_excel_bytes
safe_error_message = _u.safe_error_message
render_clear_data_button = _u.render_clear_data_button
sample_internal_transfer_template = _u.sample_internal_transfer_template
merge_internal_transfers = _u.merge_internal_transfers
process_internal_transfers = _u.process_internal_transfers

st.set_page_config(page_title="Retention Fund Dashboard", page_icon="📊", layout="wide")
render_top_nav("Retention Fund Tracker")

st.title("📊 Retention Fund Dashboard")
st.caption(
    "Statistical analysis and KPI tracking for retention fund deductions. "
    "Manage internal employee transfers and track deduction updates."
)

with st.expander("🔒 Data handling on this page", expanded=False):
    st.markdown(
        "- Uploaded files are processed **only in memory** for this browser session — nothing is written "
        "to disk, logged, or sent to any external service.\n"
        "- Downloaded reports are automatically sanitized against Excel/CSV formula-injection payloads.\n"
        "- Use the button below to explicitly wipe all cached data from this session once you're done."
    )
    render_clear_data_button()


def _show_error(e: Exception, context: str):
    """ValueError messages here are hand-written to be safe and helpful to
    show directly (e.g. 'row 5 has a non-numeric value') — only truly
    unexpected exceptions get the generic safe_error_message treatment."""
    if isinstance(e, ValueError):
        st.error(str(e))
    else:
        st.error(safe_error_message(e, context=context))


# Create tabs for Dashboard and Internal Transfers
tab1, tab2 = st.tabs(["📈 Dashboard Analytics", "🔄 Internal Transfers"])

# ================================================================
# TAB 1: DASHBOARD ANALYTICS
# ================================================================
with tab1:
    st.markdown("#### Dashboard Overview")
    
    if "retention_result" in st.session_state and not st.session_state["retention_result"].empty:
        result = st.session_state["retention_result"]
        
        # Key Metrics
        st.markdown("### Key Performance Indicators (KPIs)")
        
        col1, col2, col3, col4, col5 = st.columns(5)
        
        total_employees = len(result)
        deduction_employees = (result["Deduction Applicable"] == True).sum()
        no_deduction_employees = (result["Deduction Applicable"] == False).sum()
        total_deduction = result["Deduction Amount"].sum()
        pending_release = (result["Status"] == "Pending Release").sum()
        
        col1.metric("Total Employees", total_employees)
        col2.metric("With Deduction", deduction_employees)
        col3.metric("Without Deduction", no_deduction_employees)
        col4.metric("Pending Release", pending_release)
        col5.metric("Total Deduction Amount", f"₹{total_deduction:,.0f}")
        
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
            }).round(2)
            cc_summary.columns = ["Total Deduction", "Average Deduction", "Employee Count"]
            cc_summary = cc_summary.sort_values("Total Deduction", ascending=False)
            
            st.dataframe(cc_summary, use_container_width=True)
            
            # Chart: Total Deduction by Company Code
            st.markdown("#### Total Deduction by Company Code")
            cc_chart_data = result.groupby(cc_col)["Deduction Amount"].sum().sort_values(ascending=False)
            st.bar_chart(cc_chart_data)
        
        # Branch Analysis
        st.markdown("### Analysis by Branch")
        branch_col = "Branch"
        if branch_col in result.columns:
            branch_summary = result.groupby(branch_col).agg({
                "Deduction Amount": ["sum", "mean", "count"],
            }).round(2)
            branch_summary.columns = ["Total Deduction", "Average Deduction", "Employee Count"]
            branch_summary = branch_summary.sort_values("Total Deduction", ascending=False)
            
            st.dataframe(branch_summary, use_container_width=True)
            
            # Chart: Total Deduction by Branch
            st.markdown("#### Total Deduction by Branch")
            branch_chart_data = result.groupby(branch_col)["Deduction Amount"].sum().sort_values(ascending=False)
            st.bar_chart(branch_chart_data)
        
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
            }).round(2)
            salary_summary.columns = ["Total Deduction", "Average Deduction", "Employee Count"]
            
            st.dataframe(salary_summary, use_container_width=True)
            
            # Chart
            st.markdown("#### Distribution Across Salary Ranges")
            salary_chart = result_copy.groupby("Salary Range")["Deduction Amount"].sum()
            st.bar_chart(salary_chart)
        
        # Top Deduction Recipients
        st.markdown("### Top 10 Employees by Deduction Amount")
        top_deductions = result.nlargest(10, "Deduction Amount")[
            ["Employee ID" if "Employee ID" in result.columns else "ERP", 
             "Name" if "Name" in result.columns else "Employee Name",
             "Company Code" if "Company Code" in result.columns else "CC",
             "Branch",
             "Earned Salary" if "Earned Salary" in result.columns else "Salary",
             "Deduction Amount"]
        ].copy()
        st.dataframe(top_deductions, use_container_width=True)
        
        # Download Dashboard Report
        st.markdown("### Download Dashboard Report")
        download_button_for_df(
            result,
            "⬇️ Download Full Analytics Report",
            f"retention_dashboard_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
            key="dl_dashboard_full",
        )
        
    else:
        st.info("📌 No retention fund data available. Please upload employee master data from the main Retention Fund Tracker page first.")


# ================================================================
# TAB 2: INTERNAL TRANSFERS
# ================================================================
with tab2:
    st.markdown("#### Manage Internal Employee Transfers")
    st.caption(
        "Upload internal transfer records to update retention fund deductions based on employee movements. "
        "Supports branch changes, company code changes, and salary adjustments."
    )
    
    # Sample Template
    st.markdown("### Download Sample Template")
    sample = sample_internal_transfer_template()
    st.dataframe(sample, use_container_width=True)
    st.download_button(
        "⬇️ Download Internal Transfer Template",
        data=to_excel_bytes({"Template": sample}),
        file_name="internal_transfer_template.xlsx",
        key="dl_transfer_template",
    )
    
    # Upload Internal Transfers
    st.markdown("### Upload Internal Transfer Data")
    uploaded_transfers = st.file_uploader(
        "Upload Internal Transfers (.xlsx or .csv)",
        type=["xlsx", "xls", "csv"],
        key="internal_transfers",
    )
    
    if uploaded_transfers:
        transfers_df = read_any_table(uploaded_transfers)
        st.session_state["internal_transfers"] = transfers_df
        st.success(f"Loaded {len(transfers_df)} transfer records.")
    
    if "internal_transfers" in st.session_state and not st.session_state["internal_transfers"].empty:
        transfers_df = st.session_state["internal_transfers"]
        
        st.markdown("### Transfer Records Preview")
        st.dataframe(transfers_df, use_container_width=True, height=300)
        
        # Map Columns
        st.markdown("#### Map Transfer Data Columns")
        col1, col2, col3, col4, col5, col6, col7 = st.columns(7)
        
        with col1:
            erp_col = st.selectbox("ERP ID Column", options=transfers_df.columns, key="t_erp_col")
        with col2:
            name_col = st.selectbox("Name Column", options=transfers_df.columns, key="t_name_col")
        with col3:
            old_branch_col = st.selectbox("Old Branch Column", options=transfers_df.columns, key="t_old_branch")
        with col4:
            old_cc_col = st.selectbox("Old CC Column", options=transfers_df.columns, key="t_old_cc")
        with col5:
            new_branch_col = st.selectbox("New Branch Column", options=transfers_df.columns, key="t_new_branch")
        with col6:
            new_cc_col = st.selectbox("New CC Column", options=transfers_df.columns, key="t_new_cc")
        with col7:
            gross_change_col = st.selectbox("Gross Change Column", options=transfers_df.columns, key="t_gross_change")
        
        # Process and merge transfers
        if st.button("Apply Internal Transfers", key="apply_transfers"):
            try:
                if "retention_result" not in st.session_state or st.session_state["retention_result"].empty:
                    st.error("❌ No retention fund data available. Please calculate retention fund deductions first.")
                else:
                    # Process transfer data
                    processed_transfers = process_internal_transfers(
                        transfers_df,
                        erp_col,
                        gross_change_col,
                    )
                    
                    # Merge with retention data
                    updated_retention = merge_internal_transfers(
                        st.session_state["retention_result"].copy(),
                        processed_transfers,
                        erp_col,
                        old_cc_col,
                        new_cc_col,
                    )
                    
                    st.session_state["retention_result"] = updated_retention
                    st.success("✅ Internal transfers applied successfully!")
                    st.rerun()
                    
            except Exception as e:
                _show_error(e, "applying internal transfers")
        
        # Transfer Summary
        st.markdown("### Transfer Summary")
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Transfers", len(transfers_df))
        
        if "Old Branch" in transfers_df.columns and "New Branch" in transfers_df.columns:
            branches_changed = (transfers_df["Old Branch"] != transfers_df["New Branch"]).sum()
            col2.metric("Branch Changes", branches_changed)
        
        if "Old Branch CC" in transfers_df.columns and "New Branch CC" in transfers_df.columns:
            cc_changed = (transfers_df["Old Branch CC"] != transfers_df["New Branch CC"]).sum()
            col3.metric("Company Code Changes", cc_changed)
        
        # Transfers by Type
        st.markdown("### Transfers by Type")
        
        transfer_types = []
        if "Old Branch" in transfers_df.columns and "New Branch" in transfers_df.columns:
            branch_only = ((transfers_df["Old Branch"] != transfers_df["New Branch"]) & 
                          (transfers_df["Old Branch CC"] == transfers_df["New Branch CC"])).sum()
            transfer_types.append({"Type": "Branch Change Only", "Count": branch_only})
        
        if "Old Branch CC" in transfers_df.columns and "New Branch CC" in transfers_df.columns:
            cc_only = ((transfers_df["Old Branch CC"] != transfers_df["New Branch CC"]) & 
                      (transfers_df["Old Branch"] == transfers_df["New Branch"])).sum()
            transfer_types.append({"Type": "CC Change Only", "Count": cc_only})
            
            both = ((transfers_df["Old Branch CC"] != transfers_df["New Branch CC"]) & 
                   (transfers_df["Old Branch"] != transfers_df["New Branch"])).sum()
            transfer_types.append({"Type": "Branch & CC Change", "Count": both})
        
        if transfer_types:
            transfer_type_df = pd.DataFrame(transfer_types)
            st.dataframe(transfer_type_df, use_container_width=True)
            st.bar_chart(transfer_type_df.set_index("Type"))
        
        # Gross Changes Analysis
        if gross_change_col in transfers_df.columns:
            st.markdown("### Salary Impact Analysis")
            
            col1, col2 = st.columns(2)
            
            with col1:
                total_gross_change = pd.to_numeric(
                    transfers_df[gross_change_col].astype(str).str.replace(r"[₹$,\s]", "", regex=True),
                    errors="coerce"
                ).sum()
                st.metric("Total Gross Change", f"₹{total_gross_change:,.0f}")
            
            with col2:
                salary_increase_count = (pd.to_numeric(
                    transfers_df[gross_change_col].astype(str).str.replace(r"[₹$,\s]", "", regex=True),
                    errors="coerce"
                ) > 0).sum()
                st.metric("Salary Increases", salary_increase_count)
        
        # Download Updated Records
        st.markdown("### Download Transfer Records")
        download_button_for_df(
            transfers_df,
            "⬇️ Download Transfer Records",
            f"internal_transfers_{datetime.now().strftime('%Y%m%d')}.xlsx",
            key="dl_transfers",
        )
