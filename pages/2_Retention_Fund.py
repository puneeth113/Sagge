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
parse_wage_month_series = _u.parse_wage_month_series
process_retention_paysheet = _u.process_retention_paysheet
build_retention_reports = _u.build_retention_reports


st.set_page_config(page_title="Retention Fund", page_icon="💰", layout="wide")
render_top_nav("Retention Fund")

st.title("💰 Retention Fund")
st.caption(
    "Upload the paysheet and click Compute Retention. The system calculates the applicable retention using "
    "the configured rules and the paysheet Gross Salary / Net Pay values."
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
    """Map common payroll headings to the required Retention Full Book fields.

    Branch + Company Code + Designation are mandatory because these three
    values are the key used to match the Customize Dashboard eligibility rule.
    """
    mapping = {
        "ERP": auto_detect_column(raw_df, ["erp", "employee id", "employee code", "emp code", "erp code"]),
        "First Hire Date": auto_detect_column(
            raw_df,
            ["first hire date", "first hired date", "first hire", "first hired", "fhd"],
        ),
        "Wage Month": auto_detect_column(
            raw_df,
            ["wage month", "wage_month", "salary month", "pay month", "payroll month"],
        ),
        "Branch": auto_detect_column(
            raw_df,
            ["branch", "current branch", "curr.branch", "curr branch", "branch name", "location"],
        ),
        "Designation": auto_detect_column(
            raw_df,
            ["designation", "current designation", "curr.designation", "curr designation", "role", "job title", "position"],
        ),
        "Company Code": auto_detect_column(
            raw_df,
            ["company code", "companycode", "comp code", "company_code", "entity code", "company"],
        ),
        "Net Pay": auto_detect_column(raw_df, ["net pay", "net salary", "in-hand", "inhand"]),
        "Monthly Gross": auto_detect_column(raw_df, ["monthly gross", "gross pay", "gross salary", "gross"]),
    }
    missing = [target for target, source in mapping.items() if source is None]
    if missing:
        raise ValueError(
            f"{source_name} is missing required column(s): {', '.join(missing)}. "
            "Required format: ERP, First Hire Date, Wage Month, Branch, Designation, "
            "Company Code, Net Pay, Monthly Gross."
        )
    used = [mapping[c] for c in CONSOLIDATED_PAYSHEET_COLUMNS]
    if len(set(used)) != len(used):
        raise ValueError(
            "Some required fields were mapped to the same source column. Please use clear headings: "
            "ERP, First Hire Date, Wage Month, Branch, Designation, Company Code, Net Pay and Monthly Gross."
        )
    out = raw_df[used].copy()
    out.columns = CONSOLIDATED_PAYSHEET_COLUMNS
    return out


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
tab_compute, tab_customize = st.tabs(
    [
        "💰 Retention Compute",
        "🛠️ Customize Dashboard",
    ]
)


# ======================================================================
# TAB 1 — RETENTION COMPUTE
# ======================================================================
with tab_compute:
    st.markdown("### Retention Compute")
    st.caption(
        "Upload the paysheet, then click **Compute Retention**. Gross Salary and Net Pay are used for the "
        "retention calculation, while the other required fields apply the configured eligibility rules."
    )

    st.markdown("### Upload Paysheet")
    st.caption(
        "Upload one consolidated paysheet. Required fields are **ERP, First Hire Date, Wage Month, Branch, "
        "Designation, Company Code, Net Pay and Monthly Gross**. Gross Salary and Net Pay are used to calculate "
        "the retention amount; the other fields are used only to apply the configured retention rules."
    )

    sample = sample_consolidated_paysheet_template()
    with st.expander("View sample format", expanded=False):
        st.dataframe(sample, use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Download Paysheet Sample",
        data=to_excel_bytes({"Template": sample}),
        file_name="retention_paysheet_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_retention_full_book_template",
    )

    full_book_file = st.file_uploader(
        "Upload Paysheet (.xlsx / .csv)",
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
            with st.expander("Preview uploaded paysheet", expanded=False):
                st.dataframe(standardized_book.head(100), use_container_width=True, hide_index=True)
        except Exception as exc:
            standardized_book = None
            _show_error(exc, "reading the Retention paysheet")

    if st.button("Compute Retention", type="primary", key="calculate_retention_full_book"):
        if standardized_book is None or standardized_book.empty:
            st.warning("Upload a valid paysheet first.")
        else:
            try:
                customize_rules = load_customize_dashboard()
                ignored_company_codes = load_ignored_company_codes()
                exceptional_erps = load_exceptional_erps()
                summary, detail = process_retention_paysheet(
                    standardized_book,
                    exceptional_erps=exceptional_erps,
                    customize_dashboard=customize_rules,
                    employee_master=None,
                    ignored_company_codes=ignored_company_codes,
                    release_days=340,
                    release_review_date=date.today(),
                    deduction_pct=10.0,
                )
                reports = build_retention_reports(summary, detail)
                st.session_state["consolidated_paysheet_summary"] = summary
                st.session_state["consolidated_paysheet_result"] = detail
                st.session_state["retention_reports"] = reports
                st.success("Retention computed successfully.")
            except Exception as exc:
                _show_error(exc, "calculating Retention Fund")

    reports = st.session_state.get("retention_reports", {})
    summary = st.session_state.get("consolidated_paysheet_summary", pd.DataFrame())

    if isinstance(summary, pd.DataFrame) and not summary.empty:
        st.divider()
        st.markdown("### Download Reports")
        validation = reports.get("Deduction Validation", pd.DataFrame())
        ledger = reports.get("Retention Ledger", pd.DataFrame())

        if isinstance(validation, pd.DataFrame) and not validation.empty:
            st.download_button(
                "⬇️ Download Deduction Validation Report",
                data=to_excel_bytes({"Deduction Validation": validation}),
                file_name="retention_deduction_validation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_deduction_validation_report_compute",
            )
        if isinstance(ledger, pd.DataFrame) and not ledger.empty:
            st.download_button(
                "⬇️ Download Retention Ledger",
                data=to_excel_bytes({"Retention Ledger": ledger}),
                file_name="retention_ledger.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_retention_ledger_compute",
            )


# ======================================================================
# TAB 2 — CUSTOMIZE DASHBOARD
# ======================================================================
with tab_customize:
    st.markdown("### Customize Dashboard")
    st.caption(
        "Upload the retention eligibility master using Company Code + Designation + Eligible for Retention (Yes/No). "
        "Retention is calculated only when the employee matches an eligible rule."
    )

    st.info(
        "**Retention is calculated only when Company Code + Designation matches a rule with "
        "Eligible for Retention = Yes.** Excluded Company Codes, Exceptional ERPs and Non-Full Time "
        "employees remain excluded."
    )

    # ------------------------------------------------------------------
    # 1. Rule Master — individual add + bulk add + editable table
    # ------------------------------------------------------------------
    st.markdown("#### 1. Retention Eligibility Rule Master")
    st.caption(
        "The rule key is **Company Code + Designation**. Bulk upload requires only **Designation, Company Code "
        "and Eligible for Retention (Yes/No)**. The uploaded eligibility directly controls whether Retention Fund is calculated."
    )

    add_col, bulk_col = st.columns(2)

    with add_col:
        st.markdown("##### Add One Rule")
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
                    branch="",
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
            "Upload **Designation, Company Code and Eligible for Retention (Yes/No)**. "
            "This upload directly creates/updates the retention eligibility rules."
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
            "Upload Designation / Company Code / Eligible for Retention (Yes/No)",
            type=["xlsx", "xls", "csv"],
            key="customize_bulk_rule_upload",
        )

        if customize_file is not None:
            try:
                customize_file.seek(0)
                preview_rules = read_any_table(customize_file)
                st.dataframe(preview_rules.head(100), use_container_width=True, hide_index=True)
            except Exception as exc:
                _show_error(exc, "previewing the Customize Dashboard upload")

        if customize_file is not None and st.button(
            "Upload & Save Rules",
            key="customize_merge_bulk_rules",
            type="primary",
        ):
            try:
                customize_file.seek(0)
                uploaded_rules = read_any_table(customize_file)
                company_col = auto_detect_column(uploaded_rules, ["company code", "company", "cc"])
                designation_col = auto_detect_column(uploaded_rules, ["designation", "role", "title"])
                eligible_col = auto_detect_column(
                    uploaded_rules,
                    ["eligible for retention", "eligible for deduction", "retention applicable", "eligible"],
                )

                missing = []
                if company_col is None:
                    missing.append("Company Code")
                if designation_col is None:
                    missing.append("Designation")
                if eligible_col is None:
                    missing.append("Eligible for Retention")
                if missing:
                    raise ValueError(
                        f"Customize file is missing required column(s): {', '.join(missing)}"
                    )

                merged = merge_customize_dashboard_bulk_upload(
                    existing_df=load_customize_dashboard(),
                    upload_df=uploaded_rules,
                    branch_col=None,
                    cc_col=company_col,
                    designation_col=designation_col,
                    eligible_col=eligible_col,
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
            "Designation",
            "Company Code",
            "Employment Type",
            "Retention Applicable",
        ],
        column_config={
            "Designation": st.column_config.TextColumn("Designation", required=True),
            "Company Code": st.column_config.TextColumn("Company Code", required=True),
            "Employment Type": st.column_config.SelectboxColumn(
                "Employment Type",
                options=EMPLOYMENT_TYPE_OPTIONS,
                required=True,
                help="Non-Full Time is always excluded from Retention Fund deduction.",
            ),
            "Retention Applicable": st.column_config.SelectboxColumn(
                "Eligible for Deduction",
                options=RETENTION_APPLICABLE_OPTIONS,
                required=True,
                help="Yes/No deduction eligibility control. Non-Full Time remains excluded even if Yes is selected.",
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

