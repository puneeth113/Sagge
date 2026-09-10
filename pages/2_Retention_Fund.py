import os
import importlib.util

import streamlit as st
import pandas as pd
from datetime import datetime, date


# Retention Fund page.
# Internal Transfer logic is intentionally absent.
# Compute flow uses one full-book paysheet with Wage Month on each row.


def _load_utils():
    """Load util.py/Utils.py from the app root without changing sys.path."""
    this_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(this_dir)
    candidates = [
        os.path.join(this_dir, "util.py"),
        os.path.join(this_dir, "utils.py"),
        os.path.join(this_dir, "Utils.py"),
        os.path.join(root_dir, "util.py"),
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
        "Could not find util.py/utils.py. Keep it in the app root, one level above pages/."
    )


_u = _load_utils()
render_top_nav = _u.render_top_nav
read_any_table = _u.read_any_table
to_excel_bytes = _u.to_excel_bytes
safe_error_message = _u.safe_error_message
render_clear_data_button = _u.render_clear_data_button

load_ignored_company_codes = _u.load_ignored_company_codes
save_ignored_company_codes = _u.save_ignored_company_codes
sample_ignored_company_code_bulk_template = _u.sample_ignored_company_code_bulk_template
merge_ignored_company_code_bulk_upload = _u.merge_ignored_company_code_bulk_upload

load_exceptional_erps = _u.load_exceptional_erps
save_exceptional_erps = _u.save_exceptional_erps
sample_exceptional_erp_bulk_template = _u.sample_exceptional_erp_bulk_template
merge_exceptional_erp_bulk_upload = _u.merge_exceptional_erp_bulk_upload

EMPLOYMENT_TYPE_OPTIONS = _u.EMPLOYMENT_TYPE_OPTIONS
RETENTION_APPLICABLE_OPTIONS = _u.RETENTION_APPLICABLE_OPTIONS
CUSTOMIZE_DASHBOARD_COLUMNS = _u.CUSTOMIZE_DASHBOARD_COLUMNS
CONSOLIDATED_PAYSHEET_COLUMNS = _u.CONSOLIDATED_PAYSHEET_COLUMNS

load_customize_dashboard = _u.load_customize_dashboard
save_customize_dashboard = _u.save_customize_dashboard
sample_customize_dashboard_bulk_template = _u.sample_customize_dashboard_bulk_template
merge_customize_dashboard_bulk_upload = _u.merge_customize_dashboard_bulk_upload
upsert_customize_dashboard_rule = _u.upsert_customize_dashboard_rule
sample_consolidated_paysheet_template = _u.sample_consolidated_paysheet_template
normalise_employee_master = _u.normalise_employee_master
parse_wage_month_series = _u.parse_wage_month_series
process_retention_paysheet = _u.process_retention_paysheet
build_retention_reports = _u.build_retention_reports


st.set_page_config(page_title="Retention Fund", page_icon="💰", layout="wide")
render_top_nav("Retention Fund")

st.title("💰 Retention Fund")
st.caption(
    "Upload one Retention Full Book with Wage Month on every row. The system reconstructs the first 3 "
    "retention deductions from First Hire Date and Wage Month using min(10% of Monthly Gross, Net Pay), "
    "then separates Release, Hold and Review cases using First Hire Date + the configured release days."
)

with st.expander("🔒 Data handling on this page", expanded=False):
    st.markdown(
        "- The uploaded Retention Full Book is processed in the current Streamlit session.\n"
        "- Customize rules, excluded Company Codes and Exceptional ERPs are saved as admin configuration.\n"
        "- Downloaded reports use the shared Excel export sanitizer from util.py.\n"
        "- Use the button below to clear cached payroll/report data from this session."
    )
    render_clear_data_button()


def _show_error(exc: Exception, context: str):
    if isinstance(exc, ValueError):
        st.error(str(exc))
    else:
        st.error(safe_error_message(exc, context=context))


def auto_detect_column(df: pd.DataFrame, keywords: list):
    if df is None or len(df.columns) == 0:
        return None
    columns = list(df.columns)
    normalized = {str(c).strip().lower(): c for c in columns}
    for keyword in keywords:
        key = str(keyword).strip().lower()
        if key in normalized:
            return normalized[key]
    for col in columns:
        col_key = str(col).strip().lower()
        for keyword in keywords:
            key = str(keyword).strip().lower()
            if key and key in col_key:
                return col
    return None


def standardize_retention_full_book(raw_df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """Map common payroll headings to the five required retention fields."""
    mapping = {
        "ERP": auto_detect_column(raw_df, ["erp", "employee id", "employee code", "emp code", "erp code"]),
        "First Hire Date": auto_detect_column(
            raw_df,
            ["first hire date", "first hired date", "first hire", "first hired", "fhd"],
        ),
        "Wage Month": auto_detect_column(
            raw_df,
            ["wage month", "wage_month", "salary month", "pay month", "payroll month", "month"],
        ),
        "Net Pay": auto_detect_column(raw_df, ["net pay", "net salary", "in-hand", "inhand"]),
        "Monthly Gross": auto_detect_column(raw_df, ["monthly gross", "gross pay", "gross salary", "gross"]),
    }
    missing = [target for target, source in mapping.items() if source is None]
    if missing:
        raise ValueError(
            f"{source_name} is missing required column(s): {', '.join(missing)}. "
            "Required format: ERP, First Hire Date, Wage Month, Net Pay, Monthly Gross."
        )
    # Prevent one source column from accidentally satisfying two target fields.
    used = [mapping[c] for c in CONSOLIDATED_PAYSHEET_COLUMNS]
    if len(set(used)) != len(used):
        raise ValueError(
            "Some required fields were mapped to the same source column. Please use clear headings: "
            "ERP, First Hire Date, Wage Month, Net Pay and Monthly Gross."
        )
    out = raw_df[used].copy()
    out.columns = CONSOLIDATED_PAYSHEET_COLUMNS
    return out


def _employee_master_from_session():
    employee_db = st.session_state.get("employee_db")
    return employee_db if isinstance(employee_db, pd.DataFrame) else pd.DataFrame()


def _download_report_pack(reports: dict, key: str, label: str = "⬇️ Download Complete Retention Report"):
    nonempty = {
        name: df for name, df in reports.items()
        if isinstance(df, pd.DataFrame) and not df.empty
    }
    if not nonempty:
        return
    st.download_button(
        label,
        data=to_excel_bytes(nonempty),
        file_name=f"retention_fund_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
    )


def _clear_retention_results():
    for result_key in [
        "consolidated_paysheet_summary",
        "consolidated_paysheet_result",
        "retention_reports",
        "retention_release_days",
        "retention_release_review_date",
    ]:
        st.session_state.pop(result_key, None)


# ======================================================================
# Tabs
# ======================================================================
tab_compute, tab_customize, tab_dashboard = st.tabs(
    [
        "💰 Retention Compute",
        "🛠️ Customize Dashboard",
        "📊 Reports & Release",
    ]
)


# ======================================================================
# TAB 1 — RETENTION COMPUTE
# ======================================================================
with tab_compute:
    st.markdown("### Retention Full Book Compute")
    st.caption(
        "One file can contain many employees and many Wage Months. The system checks each ERP against "
        "First Hire Date, Customize eligibility and hard exclusions, then reconstructs the first, second "
        "and final retention deductions."
    )

    employee_master_raw = _employee_master_from_session()
    employee_master = normalise_employee_master(employee_master_raw)
    customize_rules = load_customize_dashboard()
    ignored_company_codes = load_ignored_company_codes()
    exceptional_erps = load_exceptional_erps()

    ready1, ready2, ready3, ready4 = st.columns(4)
    ready1.metric("Employee Mappings", len(employee_master))
    ready2.metric("Customize Rules", len(customize_rules))
    ready3.metric("Excluded Company Codes", len(ignored_company_codes))
    ready4.metric("Exceptional ERPs", len(exceptional_erps))

    if employee_master.empty:
        st.warning(
            "Employee Database mapping is unavailable or incomplete. Rows without ERP → Branch / Designation / "
            "Company Code mapping will be marked as no deduction for safety."
        )
    if customize_rules.empty:
        st.warning(
            "No Customize Dashboard rules are configured. Configure eligibility before computing the Retention Full Book."
        )

    st.markdown("#### 1. Retention & Release Settings")
    rel1, rel2, rel3 = st.columns([2, 2, 3])
    release_days = rel1.number_input(
        "Release after First Hire Date + days",
        min_value=0,
        max_value=1000,
        value=340,
        step=1,
        key="retention_release_days_input",
        help="Editable policy value. Example: use 330 or 340 days depending on the release cycle.",
    )
    release_review_date = rel2.date_input(
        "Evaluate release cases as of",
        value=date.today(),
        key="retention_release_review_date_input",
        help="Normally keep today's date. Change it only when you want to review release eligibility for another date.",
    )
    rel3.info(
        "**Deduction:** only employment months 1, 2 and 3 are considered. "
        "**Release Due:** 3/3 deductions completed and First Hire Date + release days has been reached. "
        "**Hold:** 1/3, 2/3 or 3/3 deductions while the release date is still pending."
    )

    st.markdown("#### 2. Upload Single Retention Full Book")
    st.caption(
        "Required columns: **ERP | First Hire Date | Wage Month | Net Pay | Monthly Gross**. "
        "Wage Month may vary row by row. Upload one consolidated full book instead of separate monthly files."
    )
    st.info(
        "For old joiners, keep the historical Wage Month rows for their first 3 employment months in this file. "
        "Because the paysheet does not contain a separate historical Retention Deduction column, the system "
        "reconstructs the deduction count from those rows using the retention formula."
    )
    sample = sample_consolidated_paysheet_template()
    with st.expander("View sample format", expanded=False):
        st.dataframe(sample, use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Download Full Book Sample",
        data=to_excel_bytes({"Template": sample}),
        file_name="retention_full_book_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_retention_full_book_template",
    )

    full_book_file = st.file_uploader(
        "Upload Retention Full Book (.xlsx / .csv)",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=False,
        key="retention_full_book_upload",
    )

    standardized_book = None
    if full_book_file is not None:
        try:
            full_book_file.seek(0)
            raw_book = read_any_table(full_book_file)
            standardized_book = standardize_retention_full_book(raw_book, full_book_file.name)
            parsed_months = parse_wage_month_series(standardized_book["Wage Month"])
            valid_months = parsed_months.dropna()

            p1, p2, p3, p4 = st.columns(4)
            p1.metric("Rows", len(standardized_book))
            p2.metric("ERPs", standardized_book["ERP"].astype(str).str.strip().nunique())
            p3.metric("Wage Months", valid_months.nunique())
            p4.metric(
                "Current Payroll Month",
                valid_months.max().strftime("%b-%Y") if not valid_months.empty else "Invalid / Missing",
            )

            with st.expander("Preview standardized full book", expanded=False):
                st.dataframe(standardized_book.head(100), use_container_width=True, hide_index=True)
        except Exception as exc:
            standardized_book = None
            _show_error(exc, "reading the Retention Full Book")

    st.markdown("#### 3. Analyze & Compute")
    if st.button("Calculate Retention Fund", type="primary", key="calculate_retention_full_book"):
        if standardized_book is None or standardized_book.empty:
            st.warning("Upload a valid Retention Full Book first.")
        else:
            try:
                summary, detail = process_retention_paysheet(
                    standardized_book,
                    exceptional_erps=exceptional_erps,
                    customize_dashboard=customize_rules,
                    employee_master=employee_master_raw,
                    ignored_company_codes=ignored_company_codes,
                    release_days=int(release_days),
                    release_review_date=release_review_date,
                    deduction_pct=10.0,
                )
                reports = build_retention_reports(summary, detail)
                st.session_state["consolidated_paysheet_summary"] = summary
                st.session_state["consolidated_paysheet_result"] = detail
                st.session_state["retention_reports"] = reports
                st.session_state["retention_release_days"] = int(release_days)
                st.session_state["retention_release_review_date"] = release_review_date
                st.success(
                    f"Retention analysis completed for {summary['ERP'].nunique() if not summary.empty else 0} employee(s)."
                )
            except Exception as exc:
                _show_error(exc, "calculating Retention Fund")

    reports = st.session_state.get("retention_reports", {})
    summary = st.session_state.get("consolidated_paysheet_summary", pd.DataFrame())
    detail = st.session_state.get("consolidated_paysheet_result", pd.DataFrame())

    if isinstance(summary, pd.DataFrame) and not summary.empty:
        st.divider()
        st.markdown("### Compute Result")

        current_month = detail["Current Payroll Month"].dropna().max() if "Current Payroll Month" in detail.columns else pd.NaT
        total_held = float(summary["Total Retention Held"].sum())
        release_cases = reports.get("Release Cases", pd.DataFrame())
        hold_cases = reports.get("Hold Cases", pd.DataFrame())
        review_cases = reports.get("Review Cases", pd.DataFrame())
        new_joiners = reports.get("New Joiners - Current Payroll", pd.DataFrame())
        current_payroll_df = reports.get("Current Payroll", pd.DataFrame())

        first_now = int(current_payroll_df["Status"].eq("First Month Deduction").sum()) if not current_payroll_df.empty else 0
        second_now = int(current_payroll_df["Status"].eq("Second Month Deduction").sum()) if not current_payroll_df.empty else 0
        final_now = int(current_payroll_df["Status"].eq("Final Month Deduction").sum()) if not current_payroll_df.empty else 0

        eligible_new_joiners = reports.get("Eligible New Joiners - Current Payroll", pd.DataFrame())
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Employees", summary["ERP"].nunique())
        m2.metric("Current Payroll", current_month.strftime("%b-%Y") if pd.notna(current_month) else "-")
        m3.metric("New Joiners", len(new_joiners))
        m4.metric("Eligible New Joiners", len(eligible_new_joiners))
        m5.metric("Total Retention Held", f"₹{total_held:,.2f}")

        d1m, d2m, d3m, d4m, d5m, d6m = st.columns(6)
        d1m.metric("1st Deduction", first_now)
        d2m.metric("2nd Deduction", second_now)
        d3m.metric("Final Deduction", final_now)
        d4m.metric("Release Due", len(release_cases))
        d5m.metric("Hold", len(hold_cases))
        d6m.metric("Review", len(review_cases))

        st.markdown("#### Employee Retention Summary")
        summary_cols = [
            "ERP", "First Hire Date", "Branch", "Designation", "Company Code",
            "Employment Type", "Retention Eligibility", "Current Payroll Action",
            "Current Payroll Deduction Amount", "Deduction Count", "Deduction Progress",
            "First Month Deduction Amount", "Second Month Deduction Amount", "Final Month Deduction Amount",
            "Total Retention Held", "Expected Release Date", "Days Until Release",
            "Release Status", "Release / Hold Remark",
        ]
        st.dataframe(summary[[c for c in summary_cols if c in summary.columns]], use_container_width=True, height=430)

        result_tabs = st.tabs([
            "Current Payroll", "New Joiners", "Eligible New Joiners", "Deduction History",
            "Release Cases", "Hold Cases", "Review Cases", "Full Book Detail"
        ])
        result_report_names = [
            "Current Payroll", "New Joiners - Current Payroll", "Eligible New Joiners - Current Payroll",
            "Deduction History", "Release Cases", "Hold Cases", "Review Cases", "Full Book Detail"
        ]
        for result_tab, report_name in zip(result_tabs, result_report_names):
            with result_tab:
                report_df = reports.get(report_name, pd.DataFrame())
                if isinstance(report_df, pd.DataFrame) and not report_df.empty:
                    st.dataframe(report_df, use_container_width=True, height=420)
                else:
                    st.info(f"No {report_name} records.")

        st.markdown("#### Downloads")
        d1, d2, d3, d4 = st.columns(4)
        with d1:
            _download_report_pack(
                reports,
                key="download_complete_retention_report_compute",
                label="⬇️ Complete Report Pack",
            )
        with d2:
            if isinstance(release_cases, pd.DataFrame) and not release_cases.empty:
                st.download_button(
                    "⬇️ Release Cases",
                    data=to_excel_bytes({"Release Cases": release_cases}),
                    file_name="retention_release_cases.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_release_cases_compute",
                )
        with d3:
            if isinstance(hold_cases, pd.DataFrame) and not hold_cases.empty:
                st.download_button(
                    "⬇️ Hold Cases",
                    data=to_excel_bytes({"Hold Cases": hold_cases}),
                    file_name="retention_hold_cases.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_hold_cases_compute",
                )
        with d4:
            if isinstance(review_cases, pd.DataFrame) and not review_cases.empty:
                st.download_button(
                    "⬇️ Review Cases",
                    data=to_excel_bytes({"Review Cases": review_cases}),
                    file_name="retention_review_cases.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_review_cases_compute",
                )


# ======================================================================
# TAB 2 — CUSTOMIZE DASHBOARD
# ======================================================================
with tab_customize:
    st.markdown("### Customize Dashboard")
    st.caption(
        "Configure who is eligible before running Retention Fund. A deduction is allowed only when the "
        "employee is Full Time, Retention Applicable is Yes, the Company Code is not excluded, and the ERP "
        "is not in the Exceptional ERP list."
    )

    st.info(
        "**No deduction** when any one of these applies: Excluded Company Code • Exceptional ERP • "
        "Non-Full Time • Retention Applicable = No."
    )

    # ------------------------------------------------------------------
    # 1. Rule Master — individual add + bulk add + editable table
    # ------------------------------------------------------------------
    st.markdown("#### 1. Branch / Designation Rule Master")
    st.caption(
        "The rule key is **Branch + Company Code + Designation**. Bulk upload needs only these three fields. "
        "Employment Type and Retention Applicable are automatically added and can be changed from dropdowns."
    )

    add_col, bulk_col = st.columns(2)

    with add_col:
        st.markdown("##### Add One Rule")
        single_branch = st.text_input(
            "Branch",
            key="customize_single_branch",
            placeholder="Example: Koramangala",
        )
        single_designation = st.text_input(
            "Designation",
            key="customize_single_designation",
            placeholder="Example: Teacher",
        )
        single_company = st.text_input(
            "Company Code",
            key="customize_single_company_code",
            placeholder="Example: MAIN",
        )
        single_employment = st.selectbox(
            "Employment Type",
            options=EMPLOYMENT_TYPE_OPTIONS,
            index=0,
            key="customize_single_employment_type",
        )
        single_retention = st.selectbox(
            "Eligible for Retention",
            options=RETENTION_APPLICABLE_OPTIONS,
            index=0,
            key="customize_single_retention_applicable",
        )

        if st.button("➕ Add / Update Rule", key="customize_add_single_rule", type="primary"):
            try:
                updated = upsert_customize_dashboard_rule(
                    existing_df=load_customize_dashboard(),
                    branch=single_branch,
                    designation=single_designation,
                    company_code=single_company,
                    employment_type=single_employment,
                    retention_applicable=single_retention,
                )
                save_customize_dashboard(updated)
                _clear_retention_results()
                st.success("Rule added/updated successfully.")
                st.rerun()
            except Exception as exc:
                _show_error(exc, "adding the Customize Dashboard rule")

    with bulk_col:
        st.markdown("##### Bulk Add Rules")
        st.caption(
            "Upload only **Branch, Company Code, Designation**. New rules automatically appear as "
            "**Full Time / Yes** until you edit the dropdowns below. Existing rule choices are preserved."
        )
        customize_sample = sample_customize_dashboard_bulk_template()
        st.download_button(
            "⬇️ Download Rule Template",
            data=to_excel_bytes({"Template": customize_sample}),
            file_name="retention_customize_rules_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_customize_rule_template",
        )
        customize_file = st.file_uploader(
            "Upload Branch / Company Code / Designation",
            type=["xlsx", "xls", "csv"],
            key="customize_bulk_rule_upload",
        )

        if customize_file is not None and st.button(
            "Upload / Merge Rules",
            key="customize_merge_bulk_rules",
            type="primary",
        ):
            try:
                customize_file.seek(0)
                uploaded_rules = read_any_table(customize_file)
                branch_col = auto_detect_column(uploaded_rules, ["branch"])
                company_col = auto_detect_column(uploaded_rules, ["company code", "company", "cc"])
                designation_col = auto_detect_column(uploaded_rules, ["designation", "role", "title"])

                missing = []
                if branch_col is None:
                    missing.append("Branch")
                if company_col is None:
                    missing.append("Company Code")
                if designation_col is None:
                    missing.append("Designation")
                if missing:
                    raise ValueError(
                        f"Customize file is missing required column(s): {', '.join(missing)}"
                    )

                merged = merge_customize_dashboard_bulk_upload(
                    existing_df=load_customize_dashboard(),
                    upload_df=uploaded_rules,
                    branch_col=branch_col,
                    cc_col=company_col,
                    designation_col=designation_col,
                )
                save_customize_dashboard(merged)
                _clear_retention_results()
                st.success(f"Bulk upload completed. Total rules: {len(merged)}.")
                st.rerun()
            except Exception as exc:
                _show_error(exc, "processing the Customize Dashboard bulk upload")

    st.markdown("##### Editable Rule Table")
    rules_df = load_customize_dashboard()
    if rules_df.empty:
        rules_df = pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)

    edited_rules = st.data_editor(
        rules_df,
        use_container_width=True,
        num_rows="dynamic",
        key="customize_rule_table_editor",
        column_order=[
            "Branch",
            "Designation",
            "Company Code",
            "Employment Type",
            "Retention Applicable",
        ],
        column_config={
            "Branch": st.column_config.TextColumn("Branch", required=True),
            "Designation": st.column_config.TextColumn("Designation", required=True),
            "Company Code": st.column_config.TextColumn("Company Code", required=True),
            "Employment Type": st.column_config.SelectboxColumn(
                "Employment Type",
                options=EMPLOYMENT_TYPE_OPTIONS,
                required=True,
                help="Non-Full Time is always excluded from Retention Fund deduction.",
            ),
            "Retention Applicable": st.column_config.SelectboxColumn(
                "Eligible for Retention",
                options=RETENTION_APPLICABLE_OPTIONS,
                required=True,
                help="Yes/No eligibility control. Non-Full Time remains excluded even if Yes is selected.",
            ),
        },
    )

    save_c1, save_c2, save_c3, save_c4 = st.columns([2, 1, 1, 1])
    if save_c1.button("💾 Save Rule Table", key="customize_save_rule_table", type="primary"):
        try:
            save_customize_dashboard(edited_rules)
            _clear_retention_results()
            st.success("Customize Dashboard saved.")
            st.rerun()
        except Exception as exc:
            _show_error(exc, "saving the Customize Dashboard")

    current_rules = load_customize_dashboard()
    if not current_rules.empty:
        save_c2.metric("Rules", len(current_rules))
        save_c3.metric(
            "Full Time",
            int(current_rules["Employment Type"].eq("Full Time").sum()),
        )
        save_c4.metric(
            "Non-Full Time",
            int(current_rules["Employment Type"].eq("Non-Full Time").sum()),
        )

    st.divider()

    # ------------------------------------------------------------------
    # 2. Hard exclusions — Company Code + ERP
    # ------------------------------------------------------------------
    st.markdown("#### 2. Deduction Exclusions")
    st.caption(
        "These are hard exclusions. They override the rule table: an employee gets ₹0 deduction when the "
        "Company Code is excluded or the ERP is listed as Exceptional."
    )

    company_col_ui, erp_col_ui = st.columns(2)

    # -------------------- Excluded Company Codes ----------------------
    with company_col_ui:
        st.markdown("##### Excluded Company Codes")
        ignored_codes = load_ignored_company_codes()

        if ignored_codes:
            st.dataframe(
                pd.DataFrame({"Company Code": ignored_codes}),
                use_container_width=True,
                hide_index=True,
                height=170,
            )
        else:
            st.info("No Company Codes are excluded.")

        company_code_input = st.text_input(
            "Add one Company Code",
            key="excluded_company_code_manual_input",
            placeholder="Example: IGNITE",
        )
        if st.button("Add Company Code", key="excluded_company_code_manual_add"):
            code = company_code_input.strip().upper()
            if not code:
                st.warning("Enter a Company Code first.")
            elif code in ignored_codes:
                st.warning(f"{code} is already excluded.")
            else:
                save_ignored_company_codes(ignored_codes + [code])
                _clear_retention_results()
                st.success(f"{code} added to the no-deduction Company Code list.")
                st.rerun()

        company_sample = sample_ignored_company_code_bulk_template()
        st.download_button(
            "⬇️ Company Code Bulk Template",
            data=to_excel_bytes({"Template": company_sample}),
            file_name="retention_excluded_company_codes_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excluded_company_code_template",
        )
        company_file = st.file_uploader(
            "Bulk upload Company Codes",
            type=["xlsx", "xls", "csv"],
            key="excluded_company_code_bulk_upload",
        )
        if company_file is not None and st.button(
            "Merge Company Codes",
            key="merge_excluded_company_codes",
        ):
            try:
                company_file.seek(0)
                company_df = read_any_table(company_file)
                cc_col = auto_detect_column(company_df, ["company code", "company", "cc"])
                if cc_col is None:
                    raise ValueError("Company Code file must contain a Company Code column.")
                merged_codes = merge_ignored_company_code_bulk_upload(
                    ignored_codes,
                    company_df,
                    company_code_col=cc_col,
                )
                save_ignored_company_codes(merged_codes)
                _clear_retention_results()
                st.success(f"Excluded Company Code list updated. Total: {len(merged_codes)}.")
                st.rerun()
            except Exception as exc:
                _show_error(exc, "processing the excluded Company Code upload")

        if ignored_codes:
            remove_company = st.selectbox(
                "Remove Company Code",
                options=[""] + ignored_codes,
                key="excluded_company_code_remove_select",
            )
            if st.button("Remove Selected Company Code", key="excluded_company_code_remove_button"):
                if not remove_company:
                    st.warning("Select a Company Code first.")
                else:
                    save_ignored_company_codes(
                        [code for code in ignored_codes if code != remove_company]
                    )
                    _clear_retention_results()
                    st.success(f"{remove_company} removed from the exclusion list.")
                    st.rerun()

    # ------------------------ Exceptional ERPs ------------------------
    with erp_col_ui:
        st.markdown("##### Exceptional ERPs")
        exceptional = load_exceptional_erps()

        if exceptional:
            st.dataframe(
                pd.DataFrame({"ERP": exceptional}),
                use_container_width=True,
                hide_index=True,
                height=170,
            )
        else:
            st.info("No Exceptional ERPs are configured.")

        exceptional_input = st.text_input(
            "Add one Exceptional ERP",
            key="exceptional_erp_manual_input",
            placeholder="Example: ERP00123",
        )
        if st.button("Add Exceptional ERP", key="exceptional_erp_manual_add"):
            erp = exceptional_input.strip().upper()
            if not erp:
                st.warning("Enter an ERP first.")
            elif erp in exceptional:
                st.warning(f"{erp} is already exceptional.")
            else:
                save_exceptional_erps(exceptional + [erp])
                _clear_retention_results()
                st.success(f"{erp} added to the no-deduction ERP list.")
                st.rerun()

        exceptional_sample = sample_exceptional_erp_bulk_template()
        st.download_button(
            "⬇️ Exceptional ERP Bulk Template",
            data=to_excel_bytes({"Template": exceptional_sample}),
            file_name="retention_exceptional_erp_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_exceptional_erp_template",
        )
        exceptional_file = st.file_uploader(
            "Bulk upload Exceptional ERPs",
            type=["xlsx", "xls", "csv"],
            key="exceptional_erp_bulk_upload",
        )
        if exceptional_file is not None and st.button(
            "Merge Exceptional ERPs",
            key="merge_exceptional_erps",
        ):
            try:
                exceptional_file.seek(0)
                exceptional_df = read_any_table(exceptional_file)
                erp_col = auto_detect_column(
                    exceptional_df,
                    ["erp", "employee id", "employee code", "emp code"],
                )
                if erp_col is None:
                    raise ValueError("Exceptional ERP file must contain an ERP column.")
                merged_erps = merge_exceptional_erp_bulk_upload(
                    exceptional,
                    exceptional_df,
                    erp_col=erp_col,
                )
                save_exceptional_erps(merged_erps)
                _clear_retention_results()
                st.success(f"Exceptional ERP list updated. Total: {len(merged_erps)}.")
                st.rerun()
            except Exception as exc:
                _show_error(exc, "processing the Exceptional ERP upload")

        if exceptional:
            remove_erp = st.selectbox(
                "Remove Exceptional ERP",
                options=[""] + exceptional,
                key="exceptional_erp_remove_select",
            )
            if st.button("Remove Selected ERP", key="exceptional_erp_remove_button"):
                if not remove_erp:
                    st.warning("Select an ERP first.")
                else:
                    save_exceptional_erps(
                        [erp for erp in exceptional if erp != remove_erp]
                    )
                    _clear_retention_results()
                    st.success(f"{remove_erp} removed from the Exceptional ERP list.")
                    st.rerun()

# ======================================================================
# TAB 3 — REPORTS & RELEASE
# ======================================================================
with tab_dashboard:
    st.markdown("### Reports & Release Dashboard")
    reports = st.session_state.get("retention_reports", {})
    summary = st.session_state.get("consolidated_paysheet_summary", pd.DataFrame())
    detail = st.session_state.get("consolidated_paysheet_result", pd.DataFrame())

    if not isinstance(summary, pd.DataFrame) or summary.empty:
        st.info("No Retention Fund result is available. Upload the Retention Full Book and calculate it first.")
    else:
        release_cases = reports.get("Release Cases", pd.DataFrame())
        hold_cases = reports.get("Hold Cases", pd.DataFrame())
        deduction_history = reports.get("Deduction History", pd.DataFrame())
        review_cases = reports.get("Review Cases", pd.DataFrame())
        not_applicable = reports.get("Not Applicable", pd.DataFrame())
        current_month = detail["Current Payroll Month"].dropna().max() if "Current Payroll Month" in detail.columns else pd.NaT

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("Employees", summary["ERP"].nunique())
        k2.metric("Current Payroll", current_month.strftime("%b-%Y") if pd.notna(current_month) else "-")
        k3.metric("3/3 Deductions", int(summary["Deduction Count"].eq(3).sum()))
        k4.metric("Release Due", len(release_cases))
        k5.metric("On Hold", len(hold_cases))
        k6.metric("Review Required", len(review_cases))

        a1, a2, a3, a4 = st.columns(4)
        a1.metric("Total Retention Held", f"₹{float(summary['Total Retention Held'].sum()):,.2f}")
        a2.metric("Release Amount", f"₹{float(summary['Release Amount'].sum()):,.2f}")
        a3.metric("Hold Amount", f"₹{float(summary['Hold Amount'].sum()):,.2f}")
        a4.metric("Not Applicable", len(not_applicable))

        st.markdown("#### Release Cases")
        st.caption("3/3 deductions completed and First Hire Date + configured release days has been reached.")
        if isinstance(release_cases, pd.DataFrame) and not release_cases.empty:
            st.dataframe(release_cases, use_container_width=True, height=360)
        else:
            st.info("No release-due cases as of the selected review date.")

        st.markdown("#### Hold Cases")
        st.caption(
            "Contains 1/3, 2/3 and 3/3 deduction cases where First Hire Date + configured release days has NOT yet been reached."
        )
        if isinstance(hold_cases, pd.DataFrame) and not hold_cases.empty:
            st.dataframe(hold_cases, use_container_width=True, height=360)
        else:
            st.info("No hold cases.")

        st.markdown("#### Review Cases")
        st.caption(
            "Release date has already been reached, but fewer than 3 deductions were reconstructed from the uploaded full book. "
            "These cases are separated from normal Hold cases for payroll review."
        )
        if isinstance(review_cases, pd.DataFrame) and not review_cases.empty:
            st.dataframe(review_cases, use_container_width=True, height=360)
        else:
            st.info("No review-required cases.")

        monthly = reports.get("Monthly Summary", pd.DataFrame())
        if isinstance(monthly, pd.DataFrame) and not monthly.empty:
            st.markdown("#### Wage Month Trend")
            chart_data = monthly.dropna(subset=["Wage Month"]).set_index("Wage Month")["Total_Retention_Fund"]
            if not chart_data.empty:
                st.bar_chart(chart_data)
            st.dataframe(monthly, use_container_width=True)

        report_tabs = st.tabs([
            "ERP Summary", "Current Payroll", "Deduction History", "Review Cases", "Not Applicable",
            "Branch", "Designation", "Company", "Employment Type", "Status", "Audit"
        ])
        report_names = [
            "ERP Summary", "Current Payroll", "Deduction History", "Review Cases", "Not Applicable",
            "Branch Summary", "Designation Summary", "Company Summary", "Employment Type Summary",
            "Status Summary", "First Hire Date Audit"
        ]
        for report_tab, report_name in zip(report_tabs, report_names):
            with report_tab:
                report_df = reports.get(report_name, pd.DataFrame())
                if isinstance(report_df, pd.DataFrame) and not report_df.empty:
                    st.dataframe(report_df, use_container_width=True, height=420)
                else:
                    st.info(f"No data available for {report_name}.")

        st.markdown("#### Download Reports")
        dl1, dl2, dl3, dl4 = st.columns(4)
        with dl1:
            _download_report_pack(
                reports,
                key="download_complete_retention_report_analytics",
                label="⬇️ Complete Report Pack",
            )
        with dl2:
            if isinstance(release_cases, pd.DataFrame) and not release_cases.empty:
                st.download_button(
                    "⬇️ Release Cases",
                    data=to_excel_bytes({"Release Cases": release_cases}),
                    file_name="retention_release_cases.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_release_cases_analytics",
                )
        with dl3:
            if isinstance(hold_cases, pd.DataFrame) and not hold_cases.empty:
                st.download_button(
                    "⬇️ Hold Cases",
                    data=to_excel_bytes({"Hold Cases": hold_cases}),
                    file_name="retention_hold_cases.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_hold_cases_analytics",
                )
        with dl4:
            if isinstance(review_cases, pd.DataFrame) and not review_cases.empty:
                st.download_button(
                    "⬇️ Review Cases",
                    data=to_excel_bytes({"Review Cases": review_cases}),
                    file_name="retention_review_cases.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_review_cases_analytics",
                )
