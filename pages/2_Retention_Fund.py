import os
import importlib.util

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
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
get_payroll_cycle = _u.get_payroll_cycle
tag_new_joiners = _u.tag_new_joiners
format_date_ddmmmyyyy = _u.format_date_ddmmmyyyy
pending_deduction_sentence = _u.pending_deduction_sentence

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
    # Horizontal step indicator
    # ------------------------------------------------------------
    step_col1, step_col2, step_col3 = st.columns(3)
    step_col1.markdown("**① Upload Employee Data**")
    step_col2.markdown("**② Internal Transfers**")
    step_col3.markdown("**③ Configure & Calculate**")
    st.divider()

    # ------------------------------------------------------------
    # STEP 1: Upload Employee Master Data
    # ------------------------------------------------------------
    st.markdown("### Step 1: Upload Employee Master Data")
    st.caption(
        "Columns required: **ERP, Name, Joining Date, Company Code, Branch, Gross Salary**. "
        "**Joining Date is mandatory** — it drives both New Joiner tagging and the expected release date."
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
        salary_col = auto_detect_column(df, ["gross", "salary", "earned"])
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
            missing_cols.append("Gross Salary")
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
        col5.info(f"📌 Gross Salary:\n`{salary_col}`")
        col6.warning(f"📌 DOJ:\n`{joining_col}`\n(Mandatory)")

        st.divider()

        # ------------------------------------------------------------
        # STEP 2: Internal Transfers (ERP change on branch/company transfer)
        # ------------------------------------------------------------
        st.markdown("### Step 2: Internal Transfers (Branch/Company Change with New ERP)")
        st.caption(
            "If an employee moved branches (and/or company code) and was issued a **new ERP**, "
            "upload the Old ERP → New ERP mapping here. It's applied automatically to the "
            "Employee Master Data from Step 1 — before deduction is computed — so any deduction "
            "still pending release follows the employee to their new ERP."
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
            "Upload Internal Transfer Mapping (.xlsx or .csv) — optional",
            type=["xlsx", "xls", "csv"],
            key="internal_transfer_upload",
        )

        if transfer_file:
            st.session_state["internal_transfers"] = read_any_table(transfer_file)
            st.success(f"Loaded {len(st.session_state['internal_transfers'])} transfer record(s).")

        if "internal_transfers" in st.session_state and not st.session_state["internal_transfers"].empty:
            transfers_df = st.session_state["internal_transfers"]

            with st.expander(f"Preview: {len(transfers_df)} transfer record(s) loaded", expanded=False):
                st.dataframe(transfers_df, use_container_width=True)
            if st.button("Clear loaded transfer mapping", key="clear_transfers"):
                del st.session_state["internal_transfers"]
                st.rerun()

            old_erp_t_col = auto_detect_column(transfers_df, ["old erp"])
            new_erp_t_col = auto_detect_column(transfers_df, ["new erp"])
            new_branch_t_col = auto_detect_column(transfers_df, ["new branch"])
            new_cc_t_col = auto_detect_column(transfers_df, ["new company", "new cc"])

            try:
                df = apply_internal_transfers(
                    df,
                    transfers_df,
                    erp_col=erp_col,
                    old_erp_col=old_erp_t_col,
                    new_erp_col=new_erp_t_col,
                    branch_col=branch_col,
                    new_branch_col=new_branch_t_col,
                    cc_col=cc_col,
                    new_cc_col=new_cc_t_col,
                )
                st.session_state["retention_fund_data"] = df

                if "Transferred" in df.columns and df["Transferred"].any():
                    st.info(
                        f"🔁 Internal transfer mapping applied — "
                        f"{int(df['Transferred'].sum())} employee(s) moved to their new ERP."
                    )
                    with st.expander("View transferred employees", expanded=False):
                        preview_cols = [c for c in ["Previous ERP", erp_col, name_col, branch_col, cc_col] if c in df.columns]
                        st.dataframe(df[df["Transferred"] == True][preview_cols], use_container_width=True)
            except Exception as e:
                _show_error(e, "applying internal transfers")

        st.divider()

        # ------------------------------------------------------------
        # STEP 3: Configure Retention Settings & Calculate
        # ------------------------------------------------------------
        st.markdown("### Step 3: Configure & Calculate")

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

        st.markdown("#### Retention Settings")

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
                help="Used to calculate each new joiner's expected release date.",
            )

        current_ignored = load_ignored_company_codes()
        st.caption(
            f"Ignored company codes currently applied: **{', '.join(current_ignored) if current_ignored else 'None'}** "
            "(manage above)."
        )

        st.markdown("##### Payroll Cycle (26th → 25th) & New Joiner Tagging")
        st.caption(
            "The retention fund deduction is applied to **New Joiners** — employees whose Joining "
            "Date falls inside the payroll cycle being processed. Existing employees, already "
            "deducted in an earlier cycle, are not re-deducted."
        )
        cycle_col1, cycle_col2 = st.columns(2)
        with cycle_col1:
            cycle_reference_date = st.date_input(
                "Payroll cycle reference date",
                value=date.today(),
                key="cycle_reference_date",
                help="Pick any date inside the payroll cycle you're processing; the 26th–25th window is derived automatically.",
            )
        with cycle_col2:
            new_joiners_only = st.checkbox(
                "Deduct only New Joiners in this cycle",
                value=True,
                key="new_joiners_only",
                help="Uncheck to apply deduction to every eligible employee regardless of Joining Date.",
            )

        cycle_start, cycle_end = get_payroll_cycle(cycle_reference_date)
        st.caption(
            f"Current payroll cycle: **{format_date_ddmmmyyyy(cycle_start)} → {format_date_ddmmmyyyy(cycle_end)}**"
        )

        if st.button("Calculate Retention Fund Deduction", key="calc_retention"):
            try:
                work_df = df.copy()

                # Tag new joiners for this payroll cycle
                work_df = tag_new_joiners(work_df, joining_col, cycle_start, cycle_end)

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
                    new_joiner_only=new_joiners_only,
                )

                # Categorize retention status
                result = categorize_retention_status(result)

                # Calculate expected release date
                result["Expected Release Date"] = result[joining_col].apply(
                    lambda x: calculate_expected_release_date(x, min_tenure)
                )

                # Format expected release date for display (dd-mmm-yyyy)
                result["Expected Release Date Formatted"] = result["Expected Release Date"].apply(format_date_ddmmmyyyy)

                # Store deduction counts
                deduction_counts = count_deductions_by_status(result)
                st.session_state["deduction_counts"] = deduction_counts
                st.session_state["payroll_cycle"] = (cycle_start, cycle_end)

                st.session_state["retention_result"] = result
                st.success(
                    f"✅ Retention fund calculation completed for payroll cycle "
                    f"{format_date_ddmmmyyyy(cycle_start)} → {format_date_ddmmmyyyy(cycle_end)}!"
                )

            except Exception as e:
                _show_error(e, "calculating retention fund deduction")

        if "retention_result" in st.session_state:
            result = st.session_state["retention_result"]
            deduction_counts = st.session_state.get("deduction_counts", {})

            st.divider()
            st.markdown("### Report")

            st.markdown("#### Retention Fund Deduction Details")
            st.dataframe(result, use_container_width=True, height=400)

            # Summary metrics
            st.markdown("#### Summary Statistics")

            total_employees = len(result)
            new_joiners_count = int(result["New Joiner"].sum()) if "New Joiner" in result.columns else 0
            deduction_employees = (result["Deduction Applicable"] == True).sum()
            no_deduction_employees = (result["Deduction Applicable"] == False).sum()
            total_deduction = result["Deduction Amount"].sum()
            pending_count = deduction_counts.get("Pending", 0)

            sum_col1, sum_col2, sum_col3, sum_col4, sum_col5, sum_col6 = st.columns(6)
            sum_col1.metric("Total Employees", total_employees)
            sum_col2.metric("New Joiners This Cycle", new_joiners_count)
            sum_col3.metric("Employees with Deduction", deduction_employees)
            sum_col4.metric("Employees without Deduction", no_deduction_employees)
            sum_col5.metric("Pending Deductions", pending_count)
            sum_col6.metric("Total Accumulated Deduction", f"₹{total_deduction:,.0f}")

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

            # Pending Deductions Summary — plain-language, e.g. "3 deductions
            # pending", "1 deduction pending with employee <name>"
            st.markdown("#### Pending Deductions Summary")
            pending_all = result[result["Status"] == "Pending Release"]
            st.markdown(pending_deduction_sentence("Overall", pending_all, name_col))
            for code, grp in result.groupby(cc_col):
                grp_pending = grp[grp["Status"] == "Pending Release"]
                st.markdown(pending_deduction_sentence(str(code), grp_pending, name_col))

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

        cycle = st.session_state.get("payroll_cycle")
        if cycle:
            st.caption(f"Payroll cycle: **{format_date_ddmmmyyyy(cycle[0])} → {format_date_ddmmmyyyy(cycle[1])}**")

        total_employees = len(result)
        new_joiners_count = int(result["New Joiner"].sum()) if "New Joiner" in result.columns else 0
        total_deduction_happened = result["Deduction Amount"].sum()
        pending_mask = result["Status"] == "Pending Release"
        total_deduction_pending = result.loc[pending_mask, "Deduction Amount"].sum()
        pending_count = int(pending_mask.sum())

        # Key Metrics — only the crucial attributes
        st.markdown("### Key Metrics")
        kcol1, kcol2, kcol3, kcol4, kcol5 = st.columns(5)
        kcol1.metric("Total Employees", total_employees)
        kcol2.metric("New Joiners This Cycle", new_joiners_count)
        kcol3.metric("Total Deduction Happened", f"₹{total_deduction_happened:,.0f}")
        kcol4.metric("Total Deduction Pending", f"₹{total_deduction_pending:,.0f}")
        kcol5.metric("Employees Awaiting Release", pending_count)

        st.markdown(pending_deduction_sentence(
            "Overall", result[pending_mask],
            "Name" if "Name" in result.columns else result.columns[1],
        ))

        # Release Month — when pending amounts are due for release
        st.markdown("### Release Month")
        if "Expected Release Date" in result.columns:
            pending_only = result[pending_mask].copy()
            if not pending_only.empty:
                pending_only["Release Month"] = pd.to_datetime(
                    pending_only["Expected Release Date"], errors="coerce"
                ).dt.strftime("%b-%Y")
                release_summary = (
                    pending_only.groupby("Release Month")["Deduction Amount"]
                    .agg(["sum", "count"])
                    .round(2)
                )
                release_summary.columns = ["Total Pending Amount", "Employees"]
                try:
                    release_summary = release_summary.reindex(
                        sorted(release_summary.index, key=lambda m: pd.to_datetime(m, format="%b-%Y"))
                    )
                except Exception:
                    pass
                st.dataframe(release_summary, use_container_width=True)
            else:
                st.info("No pending releases to show.")

        # Downloads
        st.markdown("### Download Reports")
        download_button_for_df(
            result,
            "⬇️ Download Full Dashboard Report",
            f"retention_dashboard_full_{datetime.now().strftime('%Y%m%d')}.xlsx",
            key="dl_dashboard_full",
        )

        pending_df = result[pending_mask].copy()
        if not pending_df.empty:
            download_button_for_df(
                pending_df,
                "⬇️ Download Pending Releases Report",
                f"retention_pending_releases_{datetime.now().strftime('%Y%m%d')}.xlsx",
                key="dl_pending_releases",
            )

    else:
        st.info("📌 No retention fund data available. Please calculate retention fund deductions from the Tracker tab first.")
