import os
import re
import importlib.util

import streamlit as st
import pandas as pd
from datetime import datetime, date


# This page is dedicated to Retention Fund only.
# Internal Transfer logic is intentionally absent.
# Eligibility is based on First Hire Date month + Customize Dashboard rules.


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
download_button_for_df = _u.download_button_for_df
to_excel_bytes = _u.to_excel_bytes
safe_error_message = _u.safe_error_message
render_clear_data_button = _u.render_clear_data_button
pending_deduction_sentence = _u.pending_deduction_sentence

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
process_consolidated_paysheets = _u.process_consolidated_paysheets
build_retention_reports = _u.build_retention_reports


st.set_page_config(page_title="Retention Fund", page_icon="💰", layout="wide")
render_top_nav("Retention Fund")

st.title("💰 Retention Fund")
st.caption(
    "Retention is calculated from the employee's First Hire Date month. "
    "The Customize Dashboard is the first eligibility checkpoint. Full Time + Retention Applicable = Yes "
    "is eligible only when the Company Code and ERP are not excluded. Formula: min(10% of Monthly Gross, Net Pay)."
)

with st.expander("🔒 Data handling on this page", expanded=False):
    st.markdown(
        "- Paysheets are processed in the current Streamlit session.\n"
        "- Customize Dashboard, excluded Company Codes and Exceptional ERP settings are saved as admin configuration.\n"
        "- Downloaded Excel reports use the shared export sanitizer from util.py.\n"
        "- Use the button below when you want to clear cached payroll data from this session."
    )
    render_clear_data_button()


def _show_error(exc: Exception, context: str):
    if isinstance(exc, ValueError):
        st.error(str(exc))
    else:
        st.error(safe_error_message(exc, context=context))


def auto_detect_column(df: pd.DataFrame, keywords: list):
    """Return a matching column or None. Never silently fall back to column 1."""
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


def standardize_paysheet(raw_df: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """Convert common paysheet headings into the required four standard names."""
    mapping = {
        "ERP": auto_detect_column(raw_df, ["erp", "employee id", "employee code", "emp code"]),
        "First Hire Date": auto_detect_column(
            raw_df,
            ["first hire date", "first hired date", "first hire", "first hired"],
        ),
        "Net Pay": auto_detect_column(raw_df, ["net pay", "net salary", "in-hand", "inhand"]),
        "Monthly Gross": auto_detect_column(raw_df, ["monthly gross", "gross pay", "gross salary", "gross"]),
    }
    missing = [target for target, source in mapping.items() if source is None]
    if missing:
        raise ValueError(
            f"{source_name} is missing required column(s): {', '.join(missing)}. "
            "Required format: ERP, First Hire Date, Net Pay, Monthly Gross."
        )

    out = raw_df[[mapping[c] for c in CONSOLIDATED_PAYSHEET_COLUMNS]].copy()
    out.columns = CONSOLIDATED_PAYSHEET_COLUMNS
    return out


def infer_month_from_filename(filename: str):
    """Best-effort month detection for names like Paysheet_2026-09.xlsx or Sep-2026.xlsx."""
    name = os.path.splitext(os.path.basename(filename or ""))[0]
    match = re.search(r"(?<!\d)(20\d{2})[-_ ]?(0?[1-9]|1[0-2])(?!\d)", name)
    if match:
        return date(int(match.group(1)), int(match.group(2)), 1)

    month_map = {
        "jan": 1, "january": 1,
        "feb": 2, "february": 2,
        "mar": 3, "march": 3,
        "apr": 4, "april": 4,
        "may": 5,
        "jun": 6, "june": 6,
        "jul": 7, "july": 7,
        "aug": 8, "august": 8,
        "sep": 9, "sept": 9, "september": 9,
        "oct": 10, "october": 10,
        "nov": 11, "november": 11,
        "dec": 12, "december": 12,
    }
    lower = name.lower()
    year_match = re.search(r"(?<!\d)(20\d{2})(?!\d)", lower)
    if year_match:
        for token, month_number in month_map.items():
            if re.search(rf"\b{re.escape(token)}\b", lower):
                return date(int(year_match.group(1)), month_number, 1)
    return date.today().replace(day=1)


def _employee_master_from_session():
    """Use the Employee Database module's existing session data without modifying it."""
    employee_db = st.session_state.get("employee_db")
    if isinstance(employee_db, pd.DataFrame):
        return employee_db
    return pd.DataFrame()


def _download_report_pack(reports: dict, key: str, label: str = "⬇️ Download Complete Retention Report"):
    """Render a complete report download button with a caller-supplied unique Streamlit key.

    The same report pack is shown in more than one tab. Streamlit executes all tab
    blocks on each rerun, so each rendered widget must have its own key even when
    the buttons are on different tabs.
    """
    nonempty = {name: df for name, df in reports.items() if isinstance(df, pd.DataFrame) and not df.empty}
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
    """Clear only Retention Fund calculation outputs after a rule/exclusion change."""
    for result_key in [
        "consolidated_paysheet_summary",
        "consolidated_paysheet_result",
        "retention_reports",
    ]:
        st.session_state.pop(result_key, None)


# ======================================================================
# Tabs
# ======================================================================
tab_compute, tab_customize, tab_dashboard = st.tabs(
    [
        "💰 Retention Fund Compute",
        "🛠️ Customize Dashboard",
        "📊 Analytics Dashboard",
    ]
)


# ======================================================================
# TAB 1 — RETENTION FUND COMPUTE
# ======================================================================
with tab_compute:
    st.markdown("### Consolidated Paysheet Compute")
    st.caption(
        "Upload paysheets for different months. Each monthly file must contain only the required "
        "financial fields: ERP, First Hire Date, Net Pay and Monthly Gross. Branch, Designation and "
        "Company Code are automatically mapped from the Employee Database module."
    )

    employee_master_raw = _employee_master_from_session()
    employee_master = normalise_employee_master(employee_master_raw)

    status_c1, status_c2, status_c3 = st.columns(3)
    status_c1.metric("Employee Master Rows", len(employee_master_raw) if isinstance(employee_master_raw, pd.DataFrame) else 0)
    status_c2.metric("Usable ERP Mappings", len(employee_master))
    status_c3.metric("Deduction Formula", "Min(10% Gross, Net Pay)")

    if employee_master.empty:
        st.warning(
            "Employee Database mapping is not available or does not contain ERP, Branch, Designation and Company Code. "
            "Paysheets can still be loaded, but affected ERPs will receive ₹0 deduction with status "
            "'Employee Master Mapping Missing'."
        )

    st.markdown("#### Required Paysheet Format")
    sample = sample_consolidated_paysheet_template()
    st.dataframe(sample, use_container_width=True)
    st.download_button(
        "⬇️ Download Paysheet Sample",
        data=to_excel_bytes({"Template": sample}),
        file_name="retention_consolidated_paysheet_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_retention_paysheet_template",
    )

    if "consolidated_paysheets" not in st.session_state:
        st.session_state["consolidated_paysheets"] = []

    st.markdown("#### Bulk Upload Monthly Paysheets")
    uploaded_files = st.file_uploader(
        "Select one or more monthly paysheets",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        key="retention_multi_paysheet_upload",
        help="Tip: filenames such as Paysheet_2026-09.xlsx or Sep-2026.xlsx let the month pre-fill automatically.",
    )

    file_months = []
    if uploaded_files:
        st.caption("Confirm the payroll month for each selected file before adding it to the batch.")
        for index, uploaded in enumerate(uploaded_files):
            default_month = infer_month_from_filename(uploaded.name)
            safe_key = re.sub(r"[^A-Za-z0-9_-]", "_", uploaded.name)[:70]
            c1, c2 = st.columns([3, 2])
            c1.write(f"**{uploaded.name}**")
            selected_month = c2.date_input(
                "Payroll month",
                value=default_month,
                key=f"retention_file_month_{index}_{safe_key}",
                help="Only the month and year are used.",
            )
            file_months.append((uploaded, pd.Timestamp(selected_month).to_period("M").to_timestamp()))

        if st.button("➕ Add Selected Paysheets to Batch", type="primary", key="add_selected_paysheets"):
            try:
                existing_months = {entry["month"] for entry in st.session_state["consolidated_paysheets"]}
                selected_month_values = [month for _, month in file_months]
                if len(selected_month_values) != len(set(selected_month_values)):
                    raise ValueError("Two selected files have the same payroll month. Keep only one file per month.")

                additions = []
                for uploaded, month_ts in file_months:
                    if month_ts in existing_months:
                        raise ValueError(
                            f"A paysheet for {month_ts.strftime('%b-%Y')} is already in the batch. "
                            "Remove the existing month before replacing it."
                        )
                    uploaded.seek(0)
                    raw = read_any_table(uploaded)
                    standardized = standardize_paysheet(raw, uploaded.name)
                    additions.append(
                        {
                            "month": month_ts,
                            "data": standardized,
                            "source": uploaded.name,
                        }
                    )

                st.session_state["consolidated_paysheets"].extend(additions)
                st.success(f"Added {len(additions)} monthly paysheet(s) to the batch.")
                st.rerun()
            except Exception as exc:
                _show_error(exc, "adding paysheets to the consolidated batch")

    st.markdown("#### Months Ready for Compute")
    batch = st.session_state["consolidated_paysheets"]
    if batch:
        for index, entry in enumerate(sorted(batch, key=lambda item: item["month"])):
            c1, c2, c3, c4 = st.columns([2, 3, 2, 1])
            c1.write(f"**{entry['month'].strftime('%b-%Y')}**")
            c2.write(entry.get("source", "Uploaded file"))
            c3.write(f"{len(entry['data'])} rows")
            if c4.button("Remove", key=f"remove_retention_batch_{entry['month'].strftime('%Y%m')}_{index}"):
                st.session_state["consolidated_paysheets"] = [
                    item for item in batch if item is not entry
                ]
                for result_key in [
                    "consolidated_paysheet_summary",
                    "consolidated_paysheet_result",
                    "retention_reports",
                ]:
                    st.session_state.pop(result_key, None)
                st.rerun()

        if st.button("🗑️ Clear Paysheet Batch", key="clear_retention_batch"):
            st.session_state["consolidated_paysheets"] = []
            for result_key in [
                "consolidated_paysheet_summary",
                "consolidated_paysheet_result",
                "retention_reports",
            ]:
                st.session_state.pop(result_key, None)
            st.rerun()
    else:
        st.info("No monthly paysheets have been added yet.")

    st.divider()
    st.markdown("### Calculate Retention Fund")
    customize_rules = load_customize_dashboard()
    ignored_company_codes = load_ignored_company_codes()
    exceptional_erps = load_exceptional_erps()

    ready_c1, ready_c2, ready_c3, ready_c4 = st.columns(4)
    ready_c1.metric("Customize Rules", len(customize_rules))
    ready_c2.metric("Excluded Company Codes", len(ignored_company_codes))
    ready_c3.metric("Exceptional ERPs", len(exceptional_erps))
    ready_c4.metric("Paysheet Months", len(batch))

    if customize_rules.empty:
        st.warning(
            "No Customize Dashboard rules are configured. Configure Branch + Company Code + Designation rules "
            "before computing deductions; otherwise all mapped employees will be marked 'Customize Rule Missing'."
        )

    if st.button("Calculate Consolidated Retention Fund", type="primary", key="calculate_retention_fund"):
        if not batch:
            st.warning("Add at least one monthly paysheet before calculating.")
        else:
            try:
                summary, combined = process_consolidated_paysheets(
                    monthly_inputs=batch,
                    exceptional_erps=exceptional_erps,
                    customize_dashboard=customize_rules,
                    employee_master=employee_master_raw,
                    deduction_pct=10.0,
                    ignored_company_codes=ignored_company_codes,
                )
                reports = build_retention_reports(summary, combined)

                st.session_state["consolidated_paysheet_summary"] = summary
                st.session_state["consolidated_paysheet_result"] = combined
                st.session_state["retention_reports"] = reports

                st.success(
                    f"Retention Fund computed for {summary['ERP'].nunique() if not summary.empty else 0} employee(s) "
                    f"across {combined['Paysheet Month'].nunique() if not combined.empty else 0} month(s)."
                )
                st.rerun()
            except Exception as exc:
                _show_error(exc, "calculating the consolidated Retention Fund")

    reports = st.session_state.get("retention_reports", {})
    summary = st.session_state.get("consolidated_paysheet_summary", pd.DataFrame())
    combined = st.session_state.get("consolidated_paysheet_result", pd.DataFrame())

    if isinstance(summary, pd.DataFrame) and not summary.empty:
        st.divider()
        st.markdown("### Compute Result")

        total_retention = float(summary["Total Deduction Accumulated"].sum())
        pending_employees = int((summary["Total Deduction Accumulated"] > 0).sum())
        audit_employees = int(
            summary["First Hire Date Variance"].fillna(False).sum()
            + summary["First Hire Date Recovered"].fillna(False).sum()
        )

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Employees", summary["ERP"].nunique())
        m2.metric("Months Covered", combined["Paysheet Month"].nunique())
        m3.metric("Total Retention", f"₹{total_retention:,.2f}")
        m4.metric("Hire-Date Audit Flags", audit_employees)

        st.dataframe(summary, use_container_width=True, height=420)
        _download_report_pack(reports, key="download_complete_retention_report_compute")

        pending_df = reports.get("Pending Releases", pd.DataFrame())
        if isinstance(pending_df, pd.DataFrame) and not pending_df.empty:
            st.markdown("#### Pending Releases")
            pending_view_columns = [
                c for c in [
                    "ERP", "Paysheet Month", "First Hire Date", "Branch", "Designation",
                    "Company Code", "Employment Type", "Retention Applicable",
                    "Company Code Excluded", "Exceptional ERP",
                    "Monthly Gross", "Net Pay", "Deduction Amount", "Status",
                    "Final Retention Remark",
                ] if c in pending_df.columns
            ]
            st.dataframe(pending_df[pending_view_columns], use_container_width=True)

        exception_df = reports.get("Exceptions - No Deduction", pd.DataFrame())
        if isinstance(exception_df, pd.DataFrame) and not exception_df.empty:
            with st.expander("View No-Deduction / Exception Rows", expanded=False):
                st.dataframe(exception_df, use_container_width=True)

        audit_df = reports.get("First Hire Date Audit", pd.DataFrame())
        if isinstance(audit_df, pd.DataFrame) and not audit_df.empty:
            st.warning(
                "Some ERPs have inconsistent or recovered First Hire Dates. The system used the earliest valid "
                "First Hire Date across uploaded months as the canonical date. Review the audit below."
            )
            with st.expander("First Hire Date Audit", expanded=True):
                audit_cols = [
                    c for c in [
                        "ERP", "Source File", "Paysheet Month", "Uploaded First Hire Date",
                        "First Hire Date", "First Hire Date Variance", "First Hire Date Recovered", "Status",
                    ] if c in audit_df.columns
                ]
                st.dataframe(audit_df[audit_cols], use_container_width=True)


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
# TAB 3 — ANALYTICS DASHBOARD
# ======================================================================
with tab_dashboard:
    st.markdown("### Analytics Dashboard")
    reports = st.session_state.get("retention_reports", {})
    summary = st.session_state.get("consolidated_paysheet_summary", pd.DataFrame())
    combined = st.session_state.get("consolidated_paysheet_result", pd.DataFrame())

    if not isinstance(summary, pd.DataFrame) or summary.empty:
        st.info(
            "No calculated Retention Fund data is available. Add monthly paysheets and run the calculation "
            "from the Retention Fund Compute tab."
        )
    else:
        total_retention = float(summary["Total Deduction Accumulated"].sum())
        employees_with_deduction = int((summary["Total Deduction Accumulated"] > 0).sum())
        exceptional_rows = int(combined["Exceptional ERP"].fillna(False).sum())
        excluded_company_rows = int(combined["Company Code Excluded"].fillna(False).sum()) if "Company Code Excluded" in combined.columns else 0
        non_full_time_rows = int(combined["Employment Type"].fillna("").astype(str).eq("Non-Full Time").sum()) if "Employment Type" in combined.columns else 0
        missing_mapping_rows = int(combined["Employee Master Mapping Missing"].fillna(False).sum())

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Employees", summary["ERP"].nunique())
        k2.metric("Employees With Deduction", employees_with_deduction)
        k3.metric("Total Retention Fund", f"₹{total_retention:,.2f}")
        k4.metric("Months Covered", combined["Paysheet Month"].nunique())

        q1, q2, q3, q4 = st.columns(4)
        q1.metric("Excluded Company Code Rows", excluded_company_rows)
        q2.metric("Exceptional ERP Rows", exceptional_rows)
        q3.metric("Non-Full Time Rows", non_full_time_rows)
        q4.metric("Missing Mapping Rows", missing_mapping_rows)

        monthly_summary = reports.get("Monthly Summary", pd.DataFrame())
        if isinstance(monthly_summary, pd.DataFrame) and not monthly_summary.empty:
            st.markdown("#### Monthly Trend")
            trend = monthly_summary.set_index("Paysheet Month")["Total_Retention_Fund"]
            st.bar_chart(trend)
            st.dataframe(monthly_summary, use_container_width=True)

        report_tabs = st.tabs(
            [
                "Branch",
                "Designation",
                "Company",
                "Employment Type",
                "Retention Remarks",
                "Status",
                "Exclusions",
                "Missing Mapping",
                "First Hire Audit",
            ]
        )
        report_names = [
            "Branch Summary",
            "Designation Summary",
            "Company Summary",
            "Employment Type Summary",
            "Retention Remarks Summary",
            "Status Summary",
            "Exclusion Summary",
            "Missing Mapping",
            "First Hire Date Audit",
        ]
        for report_tab, report_name in zip(report_tabs, report_names):
            with report_tab:
                report_df = reports.get(report_name, pd.DataFrame())
                if isinstance(report_df, pd.DataFrame) and not report_df.empty:
                    st.dataframe(report_df, use_container_width=True, height=420)
                else:
                    st.info(f"No data available for {report_name}.")

        st.markdown("#### Download Reports")
        _download_report_pack(reports, key="download_complete_retention_report_analytics")
