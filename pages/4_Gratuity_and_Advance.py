import os
import importlib.util
from datetime import date, datetime

import streamlit as st
import pandas as pd


def _load_utils():
    this_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(this_dir)
    for candidate in (
        os.path.join(this_dir, "utils.py"),
        os.path.join(this_dir, "Utils.py"),
        os.path.join(root_dir, "utils.py"),
        os.path.join(root_dir, "Utils.py"),
    ):
        if os.path.exists(candidate):
            spec = importlib.util.spec_from_file_location("hr_utils", candidate)
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
compute_years_of_service = _u.compute_years_of_service
compute_gratuity = _u.compute_gratuity
validate_erp_id = _u.validate_erp_id
lookup_employee_by_erp = _u.lookup_employee_by_erp
safe_error_message = _u.safe_error_message

st.set_page_config(page_title="Gratuity & Advance", page_icon="🏦", layout="wide")
render_top_nav("Gratuity & Advance")

st.title("🏦 Gratuity Calculator & Salary Advance Lookup")

tab_gratuity, tab_advance = st.tabs(["🎖️ Gratuity Calculator", "💵 Salary Advance — Paysheet Lookup"])

# ==================================================================== #
# TAB 1: Gratuity Calculator
# ==================================================================== #
with tab_gratuity:
    st.caption(
        "Statutory gratuity under the Payment of Gratuity Act, 1972: "
        "`Gratuity = (Last Drawn Basic+DA × 15 × Years of Service) / 26` for covered establishments. "
        "⚠️ Verify the current statutory exemption limit before relying on this for compliance."
    )

    mode_g = st.radio("Input method", ["Single employee", "Bulk upload"], horizontal=True, key="grat_mode")

    with st.expander("Settings", expanded=False):
        s1, s2, s3 = st.columns(3)
        with s1:
            covered_under_act = st.checkbox("Establishment covered under the Act", value=True, key="grat_covered")
            st.caption("Covered → divisor 26. Not covered → common practice uses divisor 30.")
        with s2:
            min_years = st.number_input("Minimum years for eligibility", min_value=0.0, value=5.0, key="grat_min_years")
        with s3:
            max_limit = st.number_input("Statutory max exemption limit (₹)", min_value=0.0, value=2000000.0, key="grat_max_limit")

    if mode_g == "Single employee":
        c1, c2 = st.columns(2)
        with c1:
            last_drawn = st.number_input("Last Drawn Basic + DA (₹/month)", min_value=0.0, value=25000.0, key="grat_basic")
            doj = st.date_input("Date of Joining", value=date(2017, 1, 1), key="grat_doj")
        with c2:
            dol = st.date_input("Date of Leaving (or 'as of' date)", value=date.today(), key="grat_dol")
            exempt = st.checkbox("Exempt from 5-year rule (death/disablement)", value=False, key="grat_exempt")

        if st.button("Calculate Gratuity", key="grat_calc_single"):
            if dol <= doj:
                st.error("Date of Leaving must be after Date of Joining.")
            else:
                yos = compute_years_of_service(doj, dol)
                result = compute_gratuity(
                    last_drawn, yos,
                    covered_under_act=covered_under_act,
                    min_years_for_eligibility=min_years,
                    exempt_from_min_years=exempt,
                    max_limit=max_limit,
                )
                if result["Eligible"]:
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("Years of Service", f"{result['Years of Service']}")
                    m2.metric("Gratuity (Uncapped)", f"₹{result['Gratuity (Uncapped)']:,.0f}")
                    m3.metric("Gratuity Payable", f"₹{result['Gratuity Payable']:,.0f}")
                    m4.metric("Capped at Statutory Limit?", "Yes" if result["Capped"] else "No")
                else:
                    st.warning(result["Note"])
                    st.metric("Years of Service", f"{result['Years of Service']}")

    else:  # Bulk upload
        st.caption(
            "Upload a sheet with at least: Last Drawn Basic+DA, and either "
            "**Years of Service** directly, or **Date of Joining** + **Date of Leaving** to compute it."
        )
        uploaded = st.file_uploader("Upload employee gratuity data (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="grat_bulk")

        if uploaded:
            df = read_any_table(uploaded)
            st.dataframe(df.head(5), use_container_width=True)

            cols = list(df.columns)
            basic_col = st.selectbox("Column for Last Drawn Basic + DA", cols, key="grat_basic_col")

            yos_source = st.radio(
                "Years of Service source",
                ["I have a Years of Service column", "Calculate from Date of Joining / Date of Leaving"],
                key="grat_yos_source",
            )

            if yos_source == "I have a Years of Service column":
                yos_col = st.selectbox("Column for Years of Service", cols, key="grat_yos_col")
                doj_col = dol_col = None
            else:
                yc1, yc2 = st.columns(2)
                with yc1:
                    doj_col = st.selectbox("Column for Date of Joining", cols, key="grat_doj_col")
                with yc2:
                    dol_col = st.selectbox("Column for Date of Leaving", cols, key="grat_dol_col")
                yos_col = None

            if st.button("Calculate Gratuity for all rows", key="grat_calc_bulk"):
                try:
                    work = df.copy()
                    if yos_col:
                        work["_years_of_service"] = pd.to_numeric(work[yos_col], errors="coerce")
                    else:
                        work[doj_col] = pd.to_datetime(work[doj_col], errors="coerce")
                        work[dol_col] = pd.to_datetime(work[dol_col], errors="coerce")
                        work["_years_of_service"] = work.apply(
                            lambda r: compute_years_of_service(r[doj_col], r[dol_col])
                            if pd.notna(r[doj_col]) and pd.notna(r[dol_col]) else None,
                            axis=1,
                        )

                    results_list = []
                    for _, row in work.iterrows():
                        res = compute_gratuity(
                            row[basic_col], row["_years_of_service"],
                            covered_under_act=covered_under_act,
                            min_years_for_eligibility=min_years,
                            max_limit=max_limit,
                        ) if pd.notna(row["_years_of_service"]) else {
                            "Eligible": False, "Years of Service": None,
                            "Gratuity (Uncapped)": None, "Gratuity Payable": None,
                            "Capped": None, "Note": "Missing/invalid service data",
                        }
                        results_list.append(res)

                    results_df = pd.concat([df.reset_index(drop=True), pd.DataFrame(results_list)], axis=1)
                    st.session_state["gratuity_results"] = results_df
                    st.success(f"Calculated gratuity for {len(results_df)} employees.")
                except Exception as e:
                    st.error(safe_error_message(e, context="calculating gratuity"))

            if "gratuity_results" in st.session_state:
                results_df = st.session_state["gratuity_results"]
                eligible_count = int(results_df["Eligible"].sum()) if "Eligible" in results_df else 0
                total_payable = results_df["Gratuity Payable"].fillna(0).sum() if "Gratuity Payable" in results_df else 0

                m1, m2 = st.columns(2)
                m1.metric("Eligible Employees", eligible_count)
                m2.metric("Total Gratuity Payable", f"₹{total_payable:,.0f}")

                st.dataframe(results_df, use_container_width=True, height=420)
                download_button_for_df(results_df, "⬇️ Download gratuity report", "gratuity_report.xlsx")

# ==================================================================== #
# TAB 2: Salary Advance — ERP/OIS Paysheet Lookup
# ==================================================================== #
with tab_advance:
    st.caption(
        "Upload the paysheet (first row = column headers), enter the employee's 11-digit ERP/OIS number, "
        "and choose which columns to pull — useful for checking current salary details before processing "
        "a salary advance request."
    )

    uploaded_pay = st.file_uploader("Upload Paysheet (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="adv_paysheet")

    if uploaded_pay:
        pay_df = read_any_table(uploaded_pay)
        st.session_state["advance_paysheet"] = pay_df
        st.success(f"Loaded paysheet with {len(pay_df)} rows and {len(pay_df.columns)} columns.")

    if "advance_paysheet" in st.session_state and not st.session_state["advance_paysheet"].empty:
        pay_df = st.session_state["advance_paysheet"]

        st.markdown("#### Step 1 — Which column holds the ERP/OIS number?")
        erp_col = st.selectbox("ERP/OIS Number column", list(pay_df.columns), key="adv_erp_col")

        st.markdown("#### Step 2 — Which columns should be fetched?")
        fetch_cols = st.multiselect(
            "Columns to display when an employee is found",
            options=list(pay_df.columns),
            default=list(pay_df.columns)[:5],
            key="adv_fetch_cols",
        )

        st.markdown("#### Step 3 — Enter the ERP/OIS Number")
        erp_input = st.text_input("11-digit ERP/OIS Number", max_chars=11, key="adv_erp_input")

        if erp_input:
            if not validate_erp_id(erp_input):
                st.error("ERP/OIS Number must be exactly 11 digits.")
            else:
                matches = lookup_employee_by_erp(pay_df, erp_col, erp_input)
                if matches.empty:
                    st.warning(f"No employee found with ERP/OIS Number '{erp_input}' in this paysheet.")
                else:
                    st.success(f"Found {len(matches)} matching record(s).")
                    display_cols = fetch_cols if fetch_cols else list(matches.columns)
                    st.dataframe(matches[display_cols], use_container_width=True)

                    st.divider()
                    st.markdown("#### Log a Salary Advance Request")
                    st.caption("Optional: record this request in a running ledger for this session.")

                    with st.form("advance_request_form", clear_on_submit=True):
                        ac1, ac2, ac3 = st.columns(3)
                        with ac1:
                            adv_amount = st.number_input("Advance Amount (₹)", min_value=0.0, value=0.0)
                        with ac2:
                            adv_recovery_months = st.number_input("Recovery over (months)", min_value=1, value=1)
                        with ac3:
                            adv_reason = st.text_input("Reason", value="")
                        submitted = st.form_submit_button("Add to Ledger")

                        if submitted:
                            if "advance_ledger" not in st.session_state:
                                st.session_state["advance_ledger"] = []
                            name_guess = ""
                            for c in matches.columns:
                                if "name" in c.lower():
                                    name_guess = matches.iloc[0][c]
                                    break
                            st.session_state["advance_ledger"].append({
                                "ERP/OIS Number": erp_input,
                                "Employee Name": name_guess,
                                "Advance Amount": adv_amount,
                                "Recovery (months)": adv_recovery_months,
                                "Monthly Recovery": round(adv_amount / adv_recovery_months, 2) if adv_recovery_months else adv_amount,
                                "Reason": adv_reason,
                                "Date Logged": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            })
                            st.success("Added to this session's advance ledger.")

    if "advance_ledger" in st.session_state and st.session_state["advance_ledger"]:
        st.divider()
        st.markdown("#### 📒 Salary Advance Ledger (this session)")
        ledger_df = pd.DataFrame(st.session_state["advance_ledger"])
        st.dataframe(ledger_df, use_container_width=True, height=300)
        lc1, lc2 = st.columns(2)
        lc1.metric("Total Advances Logged", f"₹{ledger_df['Advance Amount'].sum():,.0f}")
        download_button_for_df(ledger_df, "⬇️ Download advance ledger", "salary_advance_ledger.xlsx")
        if lc2.button("🗑️ Clear ledger"):
            st.session_state["advance_ledger"] = []
            st.rerun()
