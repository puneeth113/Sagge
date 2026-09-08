import os
import importlib.util

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
# NOTE: this single file replaces the previous
# pages/2_Retention_Fund_Tracker.py and pages/2a_Retention_Dashboard.py.
# Rename this file to pages/2_Retention_Fund.py so page numbering/ordering
# in the sidebar stays consistent, and update Home.py's nav link to match
# (see the Home.py update provided alongside this file).


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
sample_employee_master_for_retention = _u.sample_employee_master_for_retention
compute_retention_fund_deduction = _u.compute_retention_fund_deduction
categorize_retention_status = _u.categorize_retention_status
filter_by_company_code = _u.filter_by_company_code
filter_by_branch = _u.filter_by_branch
generate_retention_report = _u.generate_retention_report
load_ignored_company_codes = _u.load_ignored_company_codes
save_ignored_company_codes = _u.save_ignored_company_codes
sample_erp_transfer_template = _u.sample_erp_transfer_template
apply_internal_transfers = _u.apply_internal_transfers

st.set_page_config(page_title="Retention Fund", page_icon="💰", layout="wide")
render_top_nav("Retention Fund")

st.title("💰 Retention Fund")
st.caption(
    "Track employee retention fund deductions by company code and branch, and view "
    "statistical analysis and KPI tracking — all in one place."
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


def auto_detect_column(df: pd.DataFrame, keywords: list) -> str:
    """Auto-detect column by matching keywords (case-insensitive)."""
    for col in df.columns:
        col_lower = col.lower()
        for keyword in keywords:
            if keyword.lower() in col_lower:
                return col
    return df.columns[0] if len(df.columns) > 0 else None


def calculate_expected_release_date(joining_date, min_tenure_months=12):
    """Calculate expected release date based on joining date and minimum tenure.

    Args:
        joining_date: Employee joining date
        min_tenure_months: Minimum tenure in months before release

    Returns:
        Expected release date
    """
    try:
        if pd.isna(joining_date):
            return None

        # Convert to datetime if string
        if isinstance(joining_date, str):
            joining_dt = pd.to_datetime(joining_date)
        else:
            joining_dt = joining_date

        # Add minimum tenure months
        release_date = joining_dt + pd.DateOffset(months=min_tenure_months)
        return release_date
    except:
        return None


def count_deductions_by_status(df: pd.DataFrame) -> dict:
    """Count completed and pending deductions.

    Args:
        df: Result DataFrame with deduction data

    Returns:
        Dictionary with deduction counts
    """
    if "Status" not in df.columns:
        return {"Completed": 0, "Pending": 0}

    pending = (df["Status"] == "Pending Release").sum()
    completed = (df["Status"] == "No Deduction").sum()

    return {
        "Pending": pending,
        "Completed": completed,
        "Total": len(df)
    }


# ================================================================
# SUB-NAVIGATION: Tracker | Dashboard
# ================================================================
tab_tracker, tab_dashboard = st.tabs(["💰 Retention Fund Tracker", "📊 Retention Fund Dashboard"])


# ================================================================
# TAB 1: RETENTION FUND TRACKER (computation)
# ================================================================
with tab_tracker:
    # ------------------------------------------------------------
    # Ignored Company Codes (CRUD) — excluded from deduction entirely
    # ------------------------------------------------------------
    st.markdown("#### Ignored Company Codes (excluded from deduction)")
    with st.expander("Manage ignored company codes", expanded=False):
        st.caption(
            "Employees whose Company Code matches any entry below will never have a retention "
            "fund deduction applied. 'IGNITE' is the default, but you can add or remove any code."
        )

        ignored_codes = load_ignored_company_codes()

        if ignored_codes:
            for i, code in enumerate(ignored_codes):
                c1, c2 = st.columns([5, 1])
                c1.write(f"• `{code}`")
                if c2.button("Remove", key=f"remove_ignore_cc_{i}"):
                    save_ignored_company_codes([c for c in ignored_codes if c != code])
                    st.rerun()
        else:
            st.info("No company codes are currently ignored — deduction applies to everyone.")

        add_col1, add_col2 = st.columns([4, 1])
        new_ignore_code = add_col1.text_input("Add a company code to ignore", key="new_ignore_cc_input")
        if add_col2.button("Add", key="add_ignore_cc_btn"):
            new_code_clean = new_ignore_code.strip().upper()
            if not new_code_clean:
                st.warning("Enter a company code first.")
            elif new_code_clean in ignored_codes:
                st.warning(f"'{new_code_clean}' is already in the ignore list.")
            else:
                save_ignored_company_codes(ignored_codes + [new_code_clean])
                st.success(f"Added '{new_code_clean}' to the ignore list.")
                st.rerun()

    st.markdown("#### Upload Employee Master Data")
    st.caption(
        "Columns required: **ERP, Name, Joining Date, Company Code, Branch, Earned Salary**. "
        "Employees whose Company Code is in the ignore list above will not have deduction applied. "
        "**Joining Date is mandatory** to calculate deduction eligibility and expected release date."
    )

    sample = sample_employee_master_for_retention()
    st.dataframe(sample, use_container_width=True)
    st.download_button(
        "⬇️ Download sample template",
        data=to_excel_bytes({"Template": sample}),
        file_name="retention_fund_template.xlsx",
        key="dl_retention_template",
    )

    uploaded = st.file_uploader(
        "Upload Employee Master Data (.xlsx or .csv)",
        type=["xlsx", "xls", "csv"],
        key="retention_master",
    )

    if uploaded:
        df = read_any_table(uploaded)
        st.session_state["retention_fund_data"] = df
        st.success(f"Loaded {len(df)} employees.")

    if "retention_fund_data" in st.session_state and not st.session_state["retention_fund_data"].empty:
        df = st.session_state["retention_fund_data"]

        # Auto-detect columns
        erp_col = auto_detect_column(df, ["erp", "emp id", "employee id"])
        name_col = auto_detect_column(df, ["name", "employee name"])
        cc_col = auto_detect_column(df, ["cc", "company code", "company"])
        branch_col = auto_detect_column(df, ["branch"])
        salary_col = auto_detect_column(df, ["salary", "earned", "gross"])
        joining_col = auto_detect_column(df, ["joining", "doj", "date of joining"])

        # Verify all columns were detected
        missing_cols = []
        if not erp_col:
            missing_cols.append("ERP ID")
        if not name_col:
            missing_cols.append("Name")
        if not cc_col:
            missing_cols.append("Company Code")
        if not branch_col:
            missing_cols.append("Branch")
        if not salary_col:
            missing_cols.append("Earned Salary")
        if not joining_col:
            missing_cols.append("Joining Date (Mandatory)")

        if missing_cols:
            st.error(f"❌ Missing required columns: {', '.join(missing_cols)}")
            st.stop()

        # Display auto-detected columns
        st.markdown("#### Auto-Detected Columns")
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        col1.info(f"📌 ERP ID:\n`{erp_col}`")
        col2.info(f"📌 Name:\n`{name_col}`")
        col3.info(f"📌 CC:\n`{cc_col}`")
        col4.info(f"📌 Branch:\n`{branch_col}`")
        col5.info(f"📌 Salary:\n`{salary_col}`")
        col6.warning(f"📌 DOJ:\n`{joining_col}`\n(Mandatory)")

        st.markdown("#### Configure Retention Settings")

        ret_col1, ret_col2 = st.columns(2)
        with ret_col1:
            deduction_pct = st.number_input(
                "Deduction Percentage (%)",
                min_value=0.0,
                max_value=100.0,
                value=10.0,
                key="deduction_pct",
            )
        with ret_col2:
            # Minimum tenure in months to apply deduction
            min_tenure = st.number_input(
                "Minimum Tenure (months) to Apply Deduction",
                min_value=0,
                value=12,
                key="min_tenure",
                help="Employees with tenure less than this will not have deductions applied.",
            )

        current_ignored = load_ignored_company_codes()
        st.caption(
            f"Ignored company codes currently applied: **{', '.join(current_ignored) if current_ignored else 'None'}** "
            "(manage above)."
        )

        st.markdown("#### Filter by Branch (Optional)")

        all_branches = sorted(df[branch_col].astype(str).unique().tolist())
        selected_branches = st.multiselect(
            "Select Branches (leave empty for all)",
            options=all_branches,
            key="filter_branches",
        )

        if st.button("Calculate Retention Fund Deduction", key="calc_retention"):
            try:
                work_df = df.copy()

                # Apply branch filter
                if selected_branches:
                    work_df = filter_by_branch(work_df, branch_col, selected_branches)

                # Compute deductions
                result = compute_retention_fund_deduction(
                    work_df,
                    erp_col=erp_col,
                    name_col=name_col,
                    cc_col=cc_col,
                    branch_col=branch_col,
                    salary_col=salary_col,
                    deduction_pct=deduction_pct,
                    ignored_company_codes=load_ignored_company_codes(),
                )

                # Categorize retention status
                result = categorize_retention_status(result)

                # Calculate expected release date
                result["Expected Release Date"] = result[joining_col].apply(
                    lambda x: calculate_expected_release_date(x, min_tenure)
                )

                # Format expected release date for display
                result["Expected Release Date Formatted"] = result["Expected Release Date"].dt.strftime('%Y-%m-%d')

                # Store deduction counts
                deduction_counts = count_deductions_by_status(result)
                st.session_state["deduction_counts"] = deduction_counts

                st.session_state["retention_result"] = result
                st.success("✅ Retention fund calculation completed!")

            except Exception as e:
                _show_error(e, "calculating retention fund deduction")

        if "retention_result" in st.session_state:
            result = st.session_state["retention_result"]
            deduction_counts = st.session_state.get("deduction_counts", {})

            st.markdown("#### Retention Fund Deduction Details")
            st.dataframe(result, use_container_width=True, height=400)

            # Summary metrics
            st.markdown("#### Summary Statistics")

            total_employees = len(result)
            deduction_employees = (result["Deduction Applicable"] == True).sum()
            no_deduction_employees = (result["Deduction Applicable"] == False).sum()
            total_deduction = result["Deduction Amount"].sum()
            pending_count = deduction_counts.get("Pending", 0)

            sum_col1, sum_col2, sum_col3, sum_col4, sum_col5 = st.columns(5)
            sum_col1.metric("Total Employees", total_employees)
            sum_col2.metric("Employees with Deduction", deduction_employees)
            sum_col3.metric("Employees without Deduction", no_deduction_employees)
            sum_col4.metric("Pending Deductions", pending_count)
            sum_col5.metric("Total Accumulated Deduction", f"₹{total_deduction:,.0f}")

            if deduction_employees > 0:
                avg_deduction = result[result["Deduction Amount"] > 0]["Deduction Amount"].mean()
                st.metric("Average Deduction per Employee", f"₹{avg_deduction:,.0f}")

            # Breakdown by Company Code
            st.markdown("#### Deduction Breakdown by Company Code")
            cc_summary = (
                result.groupby(cc_col)
                .agg({
                    "Deduction Amount": "sum",
                    "Status": lambda x: (x == "Pending Release").sum(),
                    erp_col: "count",
                })
                .rename(columns={"Status": "Pending Count"})
            )
            cc_summary.columns = ["Total Deduction", "Pending Count", "Employee Count"]
            st.dataframe(cc_summary, use_container_width=True)

            # Breakdown by Branch
            st.markdown("#### Deduction Breakdown by Branch")
            branch_summary = (
                result.groupby(branch_col)
                .agg({
                    "Deduction Amount": "sum",
                    "Status": lambda x: (x == "Pending Release").sum(),
                    erp_col: "count",
                })
                .rename(columns={"Status": "Pending Count"})
            )
            branch_summary.columns = ["Total Deduction", "Pending Count", "Employee Count"]
            st.dataframe(branch_summary, use_container_width=True)

            # Deduction Status Summary
            st.markdown("#### Deduction Status Summary")
            status_summary = result["Status"].value_counts()
            st.dataframe(status_summary, use_container_width=True)

            # List of employees with pending releases
            st.markdown("#### Employees with Pending Release")
            pending = result[result["Status"] == "Pending Release"][
                [erp_col, name_col, cc_col, branch_col, salary_col, "Deduction Amount", "Expected Release Date Formatted"]
            ].copy()
            pending.columns = [erp_col, name_col, cc_col, branch_col, salary_col, "Deduction Amount", "Expected Release Date"]

            if not pending.empty:
                st.dataframe(pending, use_container_width=True)
            else:
                st.info("No employees with pending releases.")

            # ------------------------------------------------------------
            # Internal Transfers: employee issued a NEW ERP on transfer
            # ------------------------------------------------------------
            st.markdown("#### Internal Transfers (Branch/Company Change with New ERP)")
            st.caption(
                "If an employee moved branches (and/or company code) and was issued a **new ERP**, "
                "upload the Old ERP → New ERP mapping below. Any deduction still **Pending Release** "
                "under the old ERP is moved to the new ERP, so the eventual release happens against "
                "the employee's current ERP, not the old one."
            )

            transfer_sample = sample_erp_transfer_template()
            st.dataframe(transfer_sample, use_container_width=True)
            st.download_button(
                "⬇️ Download transfer mapping template",
                data=to_excel_bytes({"Template": transfer_sample}),
                file_name="internal_transfer_template.xlsx",
                key="dl_transfer_template",
            )

            transfer_file = st.file_uploader(
                "Upload Internal Transfer Mapping (.xlsx or .csv)",
                type=["xlsx", "xls", "csv"],
                key="internal_transfer_upload",
            )

            if transfer_file:
                st.session_state["internal_transfers"] = read_any_table(transfer_file)
                st.success(f"Loaded {len(st.session_state['internal_transfers'])} transfer record(s).")

            if "internal_transfers" in st.session_state and not st.session_state["internal_transfers"].empty:
                transfers_df = st.session_state["internal_transfers"]

                old_erp_t_col = auto_detect_column(transfers_df, ["old erp"])
                new_erp_t_col = auto_detect_column(transfers_df, ["new erp"])
                new_branch_t_col = auto_detect_column(transfers_df, ["new branch"])
                new_cc_t_col = auto_detect_column(transfers_df, ["new company", "new cc"])

                st.caption(
                    f"Detected columns → Old ERP: `{old_erp_t_col}` · New ERP: `{new_erp_t_col}` · "
                    f"New Branch: `{new_branch_t_col}` · New Company Code: `{new_cc_t_col}`"
                )

                if st.button("Apply Internal Transfers", key="apply_transfers"):
                    try:
                        updated_result = apply_internal_transfers(
                            st.session_state["retention_result"],
                            transfers_df,
                            erp_col=erp_col,
                            old_erp_col=old_erp_t_col,
                            new_erp_col=new_erp_t_col,
                            branch_col=branch_col,
                            new_branch_col=new_branch_t_col,
                            cc_col=cc_col,
                            new_cc_col=new_cc_t_col,
                        )
                        st.session_state["retention_result"] = updated_result
                        st.success("✅ Internal transfers applied. Pending deductions now follow the new ERP.")
                        st.rerun()
                    except Exception as e:
                        _show_error(e, "applying internal transfers")

                if "Transferred" in result.columns and result["Transferred"].any():
                    st.markdown("##### Employees Transferred (New ERP applied)")
                    transferred_view_cols = [
                        c for c in ["Previous ERP", erp_col, name_col, branch_col, cc_col, "Deduction Amount", "Status"]
                        if c in result.columns
                    ]
                    st.dataframe(
                        result[result["Transferred"] == True][transferred_view_cols],
                        use_container_width=True,
                    )

            # Download results
            st.markdown("#### Download Report")
            download_button_for_df(
                result,
                "⬇️ Download Full Report",
                f"retention_fund_report_{datetime.now().strftime('%Y%m%d')}.xlsx",
            )

            # Download pending releases only
            if not pending.empty:
                download_button_for_df(
                    pending,
                    "⬇️ Download Pending Releases",
                    f"retention_fund_pending_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    key="dl_pending",
                )


# ================================================================
# TAB 2: RETENTION FUND DASHBOARD (analytics)
# ================================================================
with tab_dashboard:
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
        st.info("📌 No retention fund data available. Please calculate retention fund deductions from the Tracker tab first.")
