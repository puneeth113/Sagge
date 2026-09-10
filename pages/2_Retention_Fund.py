import os
import importlib.util

import streamlit as st
import pandas as pd
from datetime import datetime, date

# NOTE: this single file replaces the previous
# pages/2_Retention_Fund_Tracker.py and pages/2a_Retention_Dashboard.py.
# Internal Transfer handling has been removed entirely — the Customize
# Dashboard tab is now the single place Branch / Company Code / Designation
# / Employment Type / Retention Applicable live for each ERP, and retention
# fund deduction is computed straight off each employee's First Hire Date
# rather than a Joining-Date-in-cycle check.


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
format_date_ddmmmyyyy = _u.format_date_ddmmmyyyy
pending_deduction_sentence = _u.pending_deduction_sentence

load_ignored_company_codes = _u.load_ignored_company_codes
save_ignored_company_codes = _u.save_ignored_company_codes
load_exceptional_erps = _u.load_exceptional_erps
save_exceptional_erps = _u.save_exceptional_erps

EMPLOYMENT_TYPE_OPTIONS = _u.EMPLOYMENT_TYPE_OPTIONS
RETENTION_APPLICABLE_OPTIONS = _u.RETENTION_APPLICABLE_OPTIONS
CUSTOMIZE_DASHBOARD_COLUMNS = _u.CUSTOMIZE_DASHBOARD_COLUMNS
load_customize_dashboard = _u.load_customize_dashboard
save_customize_dashboard = _u.save_customize_dashboard
sample_customize_dashboard_bulk_template = _u.sample_customize_dashboard_bulk_template
merge_customize_dashboard_bulk_upload = _u.merge_customize_dashboard_bulk_upload

CONSOLIDATED_PAYSHEET_COLUMNS = _u.CONSOLIDATED_PAYSHEET_COLUMNS
sample_consolidated_paysheet_template = _u.sample_consolidated_paysheet_template
compute_paysheet_deduction = _u.compute_paysheet_deduction
consolidate_paysheet_months = _u.consolidate_paysheet_months

st.set_page_config(page_title="Retention Fund", page_icon="💰", layout="wide")
render_top_nav("Retention Fund")

st.title("💰 Retention Fund")
st.caption(
    "Track employee retention fund deductions from consolidated monthly paysheets, based on each "
    "employee's First Hire Date — with per-employee Branch / Designation / Employment Type / "
    "Retention Applicable settings managed centrally in the Customize Dashboard."
)

with st.expander("🔒 Data handling on this page", expanded=False):
    st.markdown(
        "- Uploaded files are processed **only in memory** for this browser session — nothing is written "
        "to disk, logged, or sent to any external service (except your saved Customize Dashboard / "
        "Exceptional ERP / Ignored Company Code settings, which are small admin config files).\n"
        "- Downloaded reports are automatically sanitized against Excel/CSV formula-injection payloads.\n"
        "- Use the button below to explicitly wipe all cached paysheet data from this session once you're done."
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


# ================================================================
# SUB-NAVIGATION: Compute | Customize Dashboard | Analytics Dashboard
# ================================================================
tab_compute, tab_customize, tab_dashboard = st.tabs(
    ["💰 Retention Fund Compute", "🛠️ Customize Dashboard", "📊 Analytics Dashboard"]
)


# ================================================================
# TAB 1: COMPUTE (Exceptional ERPs + Consolidated Paysheet upload + calc)
# ================================================================
with tab_compute:
    step_col1, step_col2, step_col3 = st.columns(3)
    step_col1.markdown("**① Exceptional ERP Cases**")
    step_col2.markdown("**② Consolidated Paysheet Upload**")
    step_col3.markdown("**③ Calculate & Report**")
    st.divider()

    # ------------------------------------------------------------
    # STEP 1: Exceptional ERP Cases
    # ------------------------------------------------------------
    st.markdown("### Step 1: Exceptional ERP Cases")
    st.caption(
        "ERPs listed here will **never** have a retention fund deduction applied, no matter what "
        "Company Code or Retention Applicable is set to in the Customize Dashboard. Use this for "
        "one-off exceptions (e.g. a settlement, a special contract)."
    )

    with st.expander("Manage exceptional ERPs", expanded=False):
        exceptional_erps = load_exceptional_erps()

        if exceptional_erps:
            for i, erp in enumerate(exceptional_erps):
                c1, c2 = st.columns([5, 1])
                c1.write(f"• `{erp}`")
                if c2.button("Remove", key=f"remove_exceptional_erp_{i}"):
                    save_exceptional_erps([e for e in exceptional_erps if e != erp])
                    st.rerun()
        else:
            st.info("No exceptional ERPs currently configured.")

        add_col1, add_col2 = st.columns([4, 1])
        new_erp = add_col1.text_input("Add an ERP to exclude", key="new_exceptional_erp_input")
        if add_col2.button("Add", key="add_exceptional_erp_btn"):
            new_erp_clean = new_erp.strip().upper()
            if not new_erp_clean:
                st.warning("Enter an ERP first.")
            elif new_erp_clean in exceptional_erps:
                st.warning(f"'{new_erp_clean}' is already excluded.")
            else:
                save_exceptional_erps(exceptional_erps + [new_erp_clean])
                st.success(f"Added '{new_erp_clean}' to the exceptional ERP list.")
                st.rerun()

    current_exceptional = load_exceptional_erps()
    st.caption(
        f"Exceptional ERPs currently excluded: **{', '.join(current_exceptional) if current_exceptional else 'None'}**"
    )

    st.divider()

    # ------------------------------------------------------------
    # STEP 2: Consolidated Paysheet Upload (one file per month)
    # ------------------------------------------------------------
    st.markdown("### Step 2: Consolidated Paysheet Upload")
    st.caption(
        "Upload one paysheet file **per month**. Required columns: **ERP, First Hire Date, Net Pay, "
        "Monthly Gross**. Retention fund deduction for that month = **min(10% of Monthly Gross, Net Pay)**, "
        "and is only computed from the month of the employee's First Hire Date onward."
    )

    paysheet_sample = sample_consolidated_paysheet_template()
    st.dataframe(paysheet_sample, use_container_width=True)
    st.download_button(
        "⬇️ Download sample paysheet template",
        data=to_excel_bytes({"Template": paysheet_sample}),
        file_name="consolidated_paysheet_template.xlsx",
        key="dl_paysheet_template",
    )

    if "consolidated_paysheets" not in st.session_state:
        st.session_state["consolidated_paysheets"] = []  # list of {"month": Timestamp, "df": DataFrame}

    up_col1, up_col2 = st.columns([2, 3])
    with up_col1:
        paysheet_month_input = st.date_input(
            "Payroll month for this file",
            value=date.today().replace(day=1),
            key="paysheet_month_input",
            help="Pick any date inside the month this paysheet covers — only the month/year is used.",
        )
    with up_col2:
        paysheet_file = st.file_uploader(
            "Upload this month's paysheet (.xlsx or .csv)",
            type=["xlsx", "xls", "csv"],
            key="paysheet_file_uploader",
        )

    if st.button("➕ Add this month to the batch", key="add_paysheet_month_btn"):
        if paysheet_file is None:
            st.warning("Choose a file to upload first.")
        else:
            raw_df = read_any_table(paysheet_file)
            month_ts = pd.Timestamp(paysheet_month_input).to_period("M").to_timestamp()
            already = [
                m for m in st.session_state["consolidated_paysheets"]
                if m["month"] == month_ts
            ]
            if already:
                st.warning(
                    f"A paysheet for {month_ts.strftime('%b-%Y')} is already in the batch. "
                    "Remove it below first if you want to replace it."
                )
            else:
                st.session_state["consolidated_paysheets"].append({"month": month_ts, "df": raw_df})
                st.success(f"Added {month_ts.strftime('%b-%Y')} paysheet ({len(raw_df)} rows) to the batch.")
                st.rerun()

    if st.session_state["consolidated_paysheets"]:
        st.markdown("#### Months in this batch")
        for i, entry in enumerate(sorted(st.session_state["consolidated_paysheets"], key=lambda e: e["month"])):
            c1, c2, c3 = st.columns([2, 2, 1])
            c1.write(f"**{entry['month'].strftime('%b-%Y')}**")
            c2.write(f"{len(entry['df'])} rows")
            if c3.button("Remove", key=f"remove_paysheet_month_{i}"):
                st.session_state["consolidated_paysheets"] = [
                    e for e in st.session_state["consolidated_paysheets"] if e["month"] != entry["month"]
                ]
                st.rerun()
    else:
        st.info("No paysheet months added yet.")

    st.divider()

    # ------------------------------------------------------------
    # STEP 3: Calculate
    # ------------------------------------------------------------
    st.markdown("### Step 3: Calculate & Report")

    deduction_pct = st.number_input(
        "Deduction Percentage Cap (%)",
        min_value=0.0,
        max_value=100.0,
        value=10.0,
        key="deduction_pct",
        help="Deduction each month = min(this % of Monthly Gross, that month's Net Pay).",
    )

    if st.button("Calculate Consolidated Retention Fund", key="calc_retention", type="primary"):
        if not st.session_state["consolidated_paysheets"]:
            st.warning("Add at least one month's paysheet before calculating.")
        else:
            try:
                customize_df = load_customize_dashboard()
                exceptional_erps = load_exceptional_erps()

                monthly_results = []
                for entry in st.session_state["consolidated_paysheets"]:
                    raw_df = entry["df"]
                    month_ts = entry["month"]

                    erp_col = auto_detect_column(raw_df, ["erp", "emp id", "employee id"])
                    first_hire_col = auto_detect_column(raw_df, ["first hire", "hire date", "doj", "joining"])
                    net_pay_col = auto_detect_column(raw_df, ["net pay", "net salary", "in-hand", "inhand"])
                    gross_col = auto_detect_column(raw_df, ["monthly gross", "gross"])

                    missing = []
                    if not erp_col:
                        missing.append("ERP")
                    if not first_hire_col:
                        missing.append("First Hire Date")
                    if not net_pay_col:
                        missing.append("Net Pay")
                    if not gross_col:
                        missing.append("Monthly Gross")
                    if missing:
                        st.error(
                            f"❌ {month_ts.strftime('%b-%Y')} paysheet is missing required column(s): "
                            f"{', '.join(missing)}"
                        )
                        st.stop()

                    month_result = compute_paysheet_deduction(
                        raw_df,
                        erp_col=erp_col,
                        first_hire_col=first_hire_col,
                        net_pay_col=net_pay_col,
                        gross_col=gross_col,
                        paysheet_month=month_ts,
                        exceptional_erps=exceptional_erps,
                        customize_dashboard=customize_df,
                        deduction_pct=deduction_pct,
                    )
                    month_result = month_result.rename(columns={erp_col: "ERP"})
                    monthly_results.append(month_result)

                summary, combined = consolidate_paysheet_months(monthly_results, erp_col="ERP")

                st.session_state["consolidated_paysheet_summary"] = summary
                st.session_state["consolidated_paysheet_result"] = combined
                st.success(
                    f"✅ Consolidated retention fund calculated across "
                    f"{len(st.session_state['consolidated_paysheets'])} month(s) and {summary['ERP'].nunique()} employee(s)."
                )
            except Exception as e:
                _show_error(e, "calculating the consolidated retention fund")

    if "consolidated_paysheet_summary" in st.session_state and not st.session_state["consolidated_paysheet_summary"].empty:
        summary = st.session_state["consolidated_paysheet_summary"]
        combined = st.session_state["consolidated_paysheet_result"]

        st.divider()
        st.markdown("### Report")

        st.markdown("#### Consolidated Summary (per employee, across all uploaded months)")
        st.dataframe(summary, use_container_width=True, height=400)

        total_employees = summary["ERP"].nunique()
        pending_mask = summary["Latest Status"] == "Pending Release"
        pending_count = int(pending_mask.sum())
        total_deduction = summary["Total Deduction Accumulated"].sum()

        sum_col1, sum_col2, sum_col3 = st.columns(3)
        sum_col1.metric("Employees in Report", total_employees)
        sum_col2.metric("Pending Deductions", pending_count)
        sum_col3.metric("Total Accumulated Deduction", f"₹{total_deduction:,.0f}")

        st.markdown("#### Breakdown by Company Code")
        if "Company Code" in summary.columns:
            cc_summary = summary.groupby("Company Code").agg(
                **{
                    "Total Deduction": ("Total Deduction Accumulated", "sum"),
                    "Pending Count": ("Latest Status", lambda x: (x == "Pending Release").sum()),
                    "Employee Count": ("ERP", "count"),
                }
            )
            st.dataframe(cc_summary, use_container_width=True)

        st.markdown("#### Breakdown by Branch")
        if "Branch" in summary.columns:
            branch_summary = summary.groupby("Branch").agg(
                **{
                    "Total Deduction": ("Total Deduction Accumulated", "sum"),
                    "Pending Count": ("Latest Status", lambda x: (x == "Pending Release").sum()),
                    "Employee Count": ("ERP", "count"),
                }
            )
            st.dataframe(branch_summary, use_container_width=True)

        st.markdown("#### Pending Deductions Summary")
        pending_all = summary[pending_mask]
        st.markdown(pending_deduction_sentence("Overall", pending_all, "ERP"))

        st.markdown("#### Employees with Pending Release")
        pending_cols = [c for c in ["ERP", "Branch", "Designation", "Company Code", "Total Deduction Accumulated", "Months Processed"] if c in summary.columns]
        pending_view = summary[pending_mask][pending_cols]
        if not pending_view.empty:
            st.dataframe(pending_view, use_container_width=True)
        else:
            st.info("No employees with pending releases.")

        st.markdown("#### Month-wise Detail")
        with st.expander("View month-by-month deduction detail", expanded=False):
            st.dataframe(combined, use_container_width=True, height=400)

        st.markdown("#### Download Report")
        download_button_for_df(
            summary,
            "⬇️ Download Consolidated Summary",
            f"retention_fund_consolidated_summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
        )
        download_button_for_df(
            combined,
            "⬇️ Download Month-wise Detail",
            f"retention_fund_monthwise_detail_{datetime.now().strftime('%Y%m%d')}.xlsx",
            key="dl_monthwise",
        )
        if not pending_view.empty:
            download_button_for_df(
                pending_view,
                "⬇️ Download Pending Releases",
                f"retention_fund_pending_{datetime.now().strftime('%Y%m%d')}.xlsx",
                key="dl_pending",
            )


# ================================================================
# TAB 2: CUSTOMIZE DASHBOARD
# ================================================================
with tab_customize:
    st.markdown("### Customize Dashboard")
    st.caption(
        "The per-employee profile that drives reporting and retention eligibility: Branch, Designation, "
        "Employment Type, Company Code, and whether Retention Fund is Applicable (Yes/No). "
        "An ERP with no profile here, or **Retention Applicable = No**, will never get a deduction."
    )

    st.markdown("#### Bulk Upload (Branch, Company Code, Designation)")
    st.caption(
        "Upload a file with **ERP, Branch, Company Code, Designation** only — Employment Type and "
        "Retention Applicable are managed as dropdowns in the table below and are **not** overwritten "
        "by a bulk upload for ERPs that already have a value set; brand-new ERPs default to "
        "'Full Time' / 'Yes' until you change them."
    )

    bulk_sample = sample_customize_dashboard_bulk_template()
    st.dataframe(bulk_sample, use_container_width=True)
    st.download_button(
        "⬇️ Download bulk upload template",
        data=to_excel_bytes({"Template": bulk_sample}),
        file_name="customize_dashboard_bulk_template.xlsx",
        key="dl_customize_bulk_template",
    )

    bulk_file = st.file_uploader(
        "Upload Branch / Company Code / Designation (.xlsx or .csv)",
        type=["xlsx", "xls", "csv"],
        key="customize_bulk_upload",
    )

    if bulk_file:
        try:
            bulk_df = read_any_table(bulk_file)
            erp_col = auto_detect_column(bulk_df, ["erp", "emp id", "employee id"])
            branch_col = auto_detect_column(bulk_df, ["branch"])
            cc_col = auto_detect_column(bulk_df, ["cc", "company code", "company"])
            designation_col = auto_detect_column(bulk_df, ["designation", "role", "title"])

            missing = []
            if not erp_col:
                missing.append("ERP")
            if not branch_col:
                missing.append("Branch")
            if not cc_col:
                missing.append("Company Code")
            if not designation_col:
                missing.append("Designation")

            if missing:
                st.error(f"❌ Missing required columns: {', '.join(missing)}")
            else:
                existing = load_customize_dashboard()
                merged = merge_customize_dashboard_bulk_upload(
                    existing, bulk_df,
                    erp_col=erp_col, branch_col=branch_col, cc_col=cc_col, designation_col=designation_col,
                )
                save_customize_dashboard(merged)
                st.success(f"✅ Merged {len(bulk_df)} row(s) into the Customize Dashboard ({len(merged)} total employees).")
                st.rerun()
        except Exception as e:
            _show_error(e, "processing the bulk upload")

    st.divider()

    st.markdown("#### Employee Profiles")
    st.caption(
        "Edit Employment Type and Retention Applicable directly below (dropdowns). Add a brand-new "
        "ERP by typing into the blank row at the bottom. Click **Save Changes** when done."
    )

    dashboard_df = load_customize_dashboard()
    if dashboard_df.empty:
        dashboard_df = pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)

    edited_df = st.data_editor(
        dashboard_df,
        use_container_width=True,
        num_rows="dynamic",
        key="customize_dashboard_editor",
        column_config={
            "Employment Type": st.column_config.SelectboxColumn(
                "Employment Type", options=EMPLOYMENT_TYPE_OPTIONS, required=False,
            ),
            "Retention Applicable": st.column_config.SelectboxColumn(
                "Retention Applicable", options=RETENTION_APPLICABLE_OPTIONS, required=False,
            ),
        },
    )

    if st.button("💾 Save Changes", key="save_customize_dashboard_btn"):
        save_customize_dashboard(edited_df)
        st.success("Customize Dashboard saved.")
        st.rerun()


# ================================================================
# TAB 3: ANALYTICS DASHBOARD
# ================================================================
with tab_dashboard:
    st.markdown("#### Dashboard Overview")

    if "consolidated_paysheet_summary" in st.session_state and not st.session_state["consolidated_paysheet_summary"].empty:
        summary = st.session_state["consolidated_paysheet_summary"]
        combined = st.session_state["consolidated_paysheet_result"]

        total_employees = summary["ERP"].nunique()
        pending_mask = summary["Latest Status"] == "Pending Release"
        pending_count = int(pending_mask.sum())
        total_deduction = summary["Total Deduction Accumulated"].sum()
        months_covered = combined["Paysheet Month"].nunique() if "Paysheet Month" in combined.columns else 0

        st.markdown("### Key Metrics")
        kcol1, kcol2, kcol3, kcol4 = st.columns(4)
        kcol1.metric("Employees in Report", total_employees)
        kcol2.metric("Months Covered", months_covered)
        kcol3.metric("Total Deduction Accumulated", f"₹{total_deduction:,.0f}")
        kcol4.metric("Employees Awaiting Release", pending_count)

        st.markdown(pending_deduction_sentence("Overall", summary[pending_mask], "ERP"))

        st.markdown("### Employment Type Breakdown")
        if "Employment Type" in summary.columns:
            et_summary = summary.groupby("Employment Type").agg(
                **{
                    "Total Deduction": ("Total Deduction Accumulated", "sum"),
                    "Employee Count": ("ERP", "count"),
                }
            )
            st.dataframe(et_summary, use_container_width=True)

        st.markdown("### Monthly Trend")
        if "Paysheet Month" in combined.columns:
            monthly_trend = (
                combined.groupby(combined["Paysheet Month"].dt.strftime("%b-%Y"))["Deduction Amount"]
                .sum()
                .rename("Total Deduction")
            )
            st.bar_chart(monthly_trend)

        st.markdown("### Download Reports")
        download_button_for_df(
            summary,
            "⬇️ Download Full Dashboard Summary",
            f"retention_dashboard_summary_{datetime.now().strftime('%Y%m%d')}.xlsx",
            key="dl_dashboard_full",
        )
        pending_df = summary[pending_mask]
        if not pending_df.empty:
            download_button_for_df(
                pending_df,
                "⬇️ Download Pending Releases Report",
                f"retention_pending_releases_{datetime.now().strftime('%Y%m%d')}.xlsx",
                key="dl_pending_releases",
            )
    else:
        st.info(
            "📌 No retention fund data available. Upload paysheets and calculate from the "
            "**Retention Fund Compute** tab first."
        )
