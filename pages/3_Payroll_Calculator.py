import os
import sys
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
compute_fulltime_payroll = _u.compute_fulltime_payroll
compute_gig_worker_billing = _u.compute_gig_worker_billing
solve_gross_for_net_fulltime = _u.solve_gross_for_net_fulltime
gig_inhand_to_billing = _u.gig_inhand_to_billing
gig_billing_to_inhand = _u.gig_billing_to_inhand
safe_error_message = _u.safe_error_message
render_clear_data_button = _u.render_clear_data_button

st.set_page_config(page_title="Payroll Calculator", page_icon="🧾", layout="wide")
render_top_nav("Payroll Calculator")

st.title("🧾 Payroll Calculator")
st.caption("Compute monthly gross, deductions (PF/ESIC), and net pay for full-time staff and gig workers.")

with st.expander("🔒 Data handling on this page", expanded=False):
    st.markdown(
        "- Uploaded files are processed **only in memory** for this browser session — nothing is written "
        "to disk, logged, or sent to any external service.\n"
        "- Downloaded reports are automatically sanitized against Excel/CSV formula-injection payloads.\n"
        "- Use the button below to explicitly wipe all cached salary data from this session once you're done."
    )
    render_clear_data_button()

st.markdown("#### Select Employee Type")
employee_type = st.radio("Employee Category", options=["Full-Time", "Gig/Contract"], horizontal=True)

if employee_type == "Full-Time":

    # ------------------------------------------------------------------ #
    # Quick converter: Gross <-> In-Hand (single employee, no upload needed)
    # ------------------------------------------------------------------ #
    with st.expander("🔄 Quick Convert: Gross ⇄ In-Hand", expanded=True):
        st.caption(
            "Convert a single employee's Gross to In-Hand (Net) or work backwards from a "
            "desired In-Hand amount to the Gross needed to pay it. Uses the PF/ESIC rates "
            "set below (defaults are standard rates — adjust if yours differ)."
        )

        with st.expander("Advanced: customize PF/ESIC rates for this converter", expanded=False):
            qc1, qc2, qc3 = st.columns(3)
            with qc1:
                qc_apply_cap = st.checkbox("Cap PF wage at ₹15,000", value=True, key="qc_apply_cap")
                qc_pf_cap = st.number_input("PF wage cap (₹)", min_value=0.0, value=15000.0, key="qc_pf_cap", disabled=not qc_apply_cap)
            with qc2:
                qc_pf_emp = st.number_input("PF % — Employee", min_value=0.0, value=12.0, key="qc_pf_emp")
                qc_pf_employer = st.number_input("PF % — Employer", min_value=0.0, value=12.0, key="qc_pf_employer")
            with qc3:
                qc_esic_thresh = st.number_input("ESIC applicable if Gross ≤ (₹)", min_value=0.0, value=21000.0, key="qc_esic_thresh")
                qc_esic_emp = st.number_input("ESIC % — Employee", min_value=0.0, value=0.75, key="qc_esic_emp")
                qc_esic_employer = st.number_input("ESIC % — Employer", min_value=0.0, value=3.25, key="qc_esic_employer")

        conv_mode = st.radio(
            "Conversion direction",
            ["Gross → In-Hand", "In-Hand → Gross"],
            horizontal=True,
            key="ft_conv_mode",
        )

        cc1, cc2 = st.columns(2)
        with cc1:
            fixed_allow = st.number_input(
                "Fixed Allowances — HRA + Other (₹, not subject to PF)",
                min_value=0.0, value=0.0, key="ft_conv_allow",
                help="Added on top of Basic for Gross, but excluded from PF (only Basic is PF wage). Leave at 0 to treat the whole amount as Basic.",
            )

        if conv_mode == "Gross → In-Hand":
            with cc2:
                gross_input = st.number_input("Gross Salary (₹)", min_value=0.0, value=25000.0, key="ft_conv_gross")

            if st.button("Convert", key="ft_conv_btn_g2n"):
                basic = max(gross_input - fixed_allow, 0)
                df_single = pd.DataFrame([{"Basic": basic, "Other": fixed_allow}])
                row = compute_fulltime_payroll(
                    df_single, "Basic", None, "Other",
                    pf_wage_cap=qc_pf_cap, apply_pf_cap=qc_apply_cap,
                    pf_employee_pct=qc_pf_emp, pf_employer_pct=qc_pf_employer,
                    esic_threshold=qc_esic_thresh,
                    esic_employee_pct=qc_esic_emp, esic_employer_pct=qc_esic_employer,
                ).iloc[0]

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Gross Salary", f"₹{row['Gross Salary']:,.0f}")
                m2.metric("PF (Employee)", f"₹{row['PF (Employee)']:,.0f}")
                m3.metric("ESIC (Employee)", f"₹{row['ESIC (Employee)']:,.0f}")
                m4.metric("In-Hand (Net)", f"₹{row['Net Pay (Employee Take-home)']:,.0f}")

        else:  # In-Hand -> Gross
            with cc2:
                net_input = st.number_input("Desired In-Hand / Net Pay (₹)", min_value=0.0, value=20000.0, key="ft_conv_net")

            if st.button("Convert", key="ft_conv_btn_n2g"):
                solved = solve_gross_for_net_fulltime(
                    target_net=net_input, fixed_allowances=fixed_allow,
                    pf_wage_cap=qc_pf_cap, apply_pf_cap=qc_apply_cap,
                    pf_employee_pct=qc_pf_emp, pf_employer_pct=qc_pf_employer,
                    esic_threshold=qc_esic_thresh,
                    esic_employee_pct=qc_esic_emp, esic_employer_pct=qc_esic_employer,
                )
                row = solved["row"]

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Required Gross", f"₹{row['Gross Salary']:,.0f}")
                m2.metric("PF (Employee)", f"₹{row['PF (Employee)']:,.0f}")
                m3.metric("ESIC (Employee)", f"₹{row['ESIC (Employee)']:,.0f}")
                m4.metric("Resulting In-Hand", f"₹{row['Net Pay (Employee Take-home)']:,.0f}")
                st.caption(f"↳ Basic Salary component: ₹{solved['Basic']:,.0f}  |  Fixed Allowances: ₹{fixed_allow:,.0f}")

    st.divider()

    # ------------------------------------------------------------------ #
    # Bulk upload payroll calculation (unchanged from your version)
    # ------------------------------------------------------------------ #
    st.markdown("#### Upload Full-Time Employee Data")
    st.caption("Columns: Employee ID, Name, Basic, HRA (optional), Other Allowances (optional)")

    uploaded = st.file_uploader("Upload Full-Time Employee Data (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="fulltime")

    if uploaded:
        df = read_any_table(uploaded)
        st.session_state["fulltime_data"] = df
        st.success(f"Loaded {len(df)} employees.")

    if "fulltime_data" in st.session_state and not st.session_state["fulltime_data"].empty:
        df = st.session_state["fulltime_data"]

        st.markdown("#### Map Columns")
        col1, col2, col3 = st.columns(3)
        with col1:
            basic_col = st.selectbox("Basic Column", options=df.columns, key="basic_col")
        with col2:
            hra_col = st.selectbox("HRA Column (optional)", options=["-"] + list(df.columns), key="hra_col")
        with col3:
            other_col = st.selectbox("Other Allowances (optional)", options=["-"] + list(df.columns), key="other_col")

        st.markdown("#### Configure Statutory Rates & Thresholds")

        st.markdown("**PF Settings**")
        pf_col1, pf_col2, pf_col3 = st.columns(3)
        with pf_col1:
            pf_emp_pct = st.number_input("PF Employee %", value=12.0, min_value=0.0, max_value=50.0, key="pf_emp")
        with pf_col2:
            pf_employer_pct = st.number_input("PF Employer %", value=12.0, min_value=0.0, max_value=50.0, key="pf_emp_rate")
        with pf_col3:
            pf_wage_cap = st.number_input("PF Wage Cap (₹)", value=15000.0, min_value=0.0, key="pf_cap")

        st.markdown("**ESIC Settings**")
        esic_col1, esic_col2, esic_col3 = st.columns(3)
        with esic_col1:
            esic_threshold = st.number_input("ESIC Threshold (₹)", value=21000.0, min_value=0.0, key="esic_thresh")
        with esic_col2:
            esic_emp_pct = st.number_input("ESIC Employee %", value=0.75, min_value=0.0, max_value=5.0, key="esic_emp")
        with esic_col3:
            esic_employer_pct = st.number_input("ESIC Employer %", value=3.25, min_value=0.0, max_value=10.0, key="esic_emp_rate")

        st.markdown("**Apply PF Wage Cap**")
        apply_pf_cap = st.checkbox("Apply ₹15,000 wage cap for PF", value=True)

        if st.button("Calculate Payroll", key="calc_fulltime"):
            try:
                result = compute_fulltime_payroll(
                    df,
                    basic_col=basic_col,
                    hra_col=None if hra_col == "-" else hra_col,
                    other_allow_col=None if other_col == "-" else other_col,
                    pf_wage_cap=pf_wage_cap,
                    apply_pf_cap=apply_pf_cap,
                    pf_employee_pct=pf_emp_pct,
                    pf_employer_pct=pf_employer_pct,
                    esic_threshold=esic_threshold,
                    esic_employee_pct=esic_emp_pct,
                    esic_employer_pct=esic_employer_pct,
                )
                st.session_state["payroll_result"] = result
                st.success("✅ Payroll calculated successfully!")
            except Exception as e:
                st.error(safe_error_message(e, context="calculating payroll"))

        if "payroll_result" in st.session_state:
            result = st.session_state["payroll_result"]

            st.markdown("#### Payroll Results")
            st.dataframe(result, use_container_width=True, height=400)

            st.markdown("#### Employee Deductions Summary")
            sum_col1, sum_col2, sum_col3, sum_col4 = st.columns(4)
            sum_col1.metric("Total Gross", f"₹{result['Gross Salary'].sum():,.0f}")
            sum_col2.metric("Total Employee PF", f"₹{result['PF (Employee)'].sum():,.0f}")
            sum_col3.metric("Total Employee ESIC", f"₹{result['ESIC (Employee)'].sum():,.0f}")
            sum_col4.metric("Total Net Pay", f"₹{result['Net Pay (Employee Take-home)'].sum():,.0f}")

            st.markdown("#### Employer Cost Summary")
            emp_col1, emp_col2, emp_col3 = st.columns(3)
            emp_col1.metric("Total Employer PF", f"₹{result['PF (Employer)'].sum():,.0f}")
            emp_col2.metric("Total Employer ESIC", f"₹{result['ESIC (Employer)'].sum():,.0f}")
            emp_col3.metric("Total Employer Cost", f"₹{result['Total Employer Cost (CTC add-on)'].sum():,.0f}")

            st.markdown("#### Cost Breakdown")
            cost_chart_data = pd.DataFrame({
                "Employee": ["Gross Salary", "PF (Employee)", "ESIC (Employee)", "Net Pay"],
                "Amount": [
                    result['Gross Salary'].sum(),
                    result['PF (Employee)'].sum(),
                    result['ESIC (Employee)'].sum(),
                    result['Net Pay (Employee Take-home)'].sum(),
                ]
            })
            st.bar_chart(cost_chart_data.set_index('Employee'))

            download_button_for_df(result, "⬇️ Download Payroll", f"fulltime_payroll_{datetime.now().strftime('%Y%m%d')}.xlsx")

else:  # Gig/Contract

    # ------------------------------------------------------------------ #
    # Quick converter: In-Hand <-> Billing (single worker, no upload needed)
    # ------------------------------------------------------------------ #
    with st.expander("🔄 Quick Convert: In-Hand ⇄ Billing", expanded=True):
        st.caption(
            "Convert a single gig worker's In-Hand payment to the Billing (invoice) amount "
            "that includes TDS, or work backwards from a known Billing amount to what the "
            "worker actually receives in-hand."
        )

        gc1, gc2 = st.columns(2)
        with gc1:
            gig_conv_mode = st.radio(
                "Conversion direction",
                ["In-Hand → Billing", "Billing → In-Hand"],
                horizontal=True,
                key="gig_conv_mode",
            )
        with gc2:
            gig_conv_tds = st.number_input("TDS %", min_value=0.0, value=1.0, step=0.1, key="gig_conv_tds")

        if gig_conv_mode == "In-Hand → Billing":
            inhand_val = st.number_input("In-Hand Amount (₹)", min_value=0.0, value=20000.0, key="gig_conv_inhand")
            if st.button("Convert", key="gig_conv_btn_i2b"):
                res = gig_inhand_to_billing(inhand_val, gig_conv_tds)
                g1, g2, g3 = st.columns(3)
                g1.metric("In-Hand Amount", f"₹{res['In-Hand Amount']:,.0f}")
                g2.metric("TDS Amount", f"₹{res['TDS Amount']:,.0f}")
                g3.metric("Billing Amount", f"₹{res['Billing Amount']:,.0f}")
        else:
            billing_val = st.number_input("Billing Amount (₹)", min_value=0.0, value=20200.0, key="gig_conv_billing")
            if st.button("Convert", key="gig_conv_btn_b2i"):
                res = gig_billing_to_inhand(billing_val, gig_conv_tds)
                g1, g2, g3 = st.columns(3)
                g1.metric("Billing Amount", f"₹{res['Billing Amount']:,.0f}")
                g2.metric("TDS Amount", f"₹{res['TDS Amount']:,.0f}")
                g3.metric("In-Hand Amount", f"₹{res['In-Hand Amount']:,.0f}")

    st.divider()

    # ------------------------------------------------------------------ #
    # Bulk upload billing calculation (unchanged from your version)
    # ------------------------------------------------------------------ #
    st.markdown("#### Upload Gig Worker Data")
    st.caption("Columns: Worker ID, Name, Payment Amount")

    uploaded = st.file_uploader("Upload Gig Worker Data (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="gig")

    if uploaded:
        df = read_any_table(uploaded)
        st.session_state["gig_data"] = df
        st.success(f"Loaded {len(df)} gig workers.")

    if "gig_data" in st.session_state and not st.session_state["gig_data"].empty:
        df = st.session_state["gig_data"]

        st.markdown("#### Map Columns")
        amount_col = st.selectbox("Payment Amount Column", options=df.columns, key="amount_col")

        st.markdown("#### TDS Settings")
        col1, col2 = st.columns(2)
        with col1:
            tds_pct = st.number_input("TDS % (to be added to billing)", value=1.0, min_value=0.0, max_value=50.0, key="tds_pct")
        with col2:
            st.info(f"💡 1% TDS adds: ₹{1000 * 0.01:.0f} per ₹1,000 payment")

        if st.button("Calculate Billing", key="calc_gig"):
            try:
                result = compute_gig_worker_billing(df, amount_col, tds_pct)
                st.session_state["gig_result"] = result
                st.success("✅ Billing calculated successfully!")
            except Exception as e:
                st.error(safe_error_message(e, context="calculating billing"))

        if "gig_result" in st.session_state:
            result = st.session_state["gig_result"]

            st.markdown("#### Billing Results")
            st.dataframe(result, use_container_width=True, height=400)

            st.markdown("#### Billing Summary")
            sum_col1, sum_col2, sum_col3 = st.columns(3)
            sum_col1.metric("Total Payment", f"₹{result[amount_col].sum():,.0f}")
            sum_col2.metric("Total TDS", f"₹{result['TDS Amount'].sum():,.0f}")
            sum_col3.metric("Total Billing", f"₹{result['Monthly Billing Amount'].sum():,.0f}")

            st.markdown("#### Billing Breakdown")
            billing_chart = pd.DataFrame({
                "Category": ["Payment", "TDS"],
                "Amount": [
                    result[amount_col].sum(),
                    result['TDS Amount'].sum()
                ]
            })
            st.bar_chart(billing_chart.set_index('Category'))

            download_button_for_df(result, "⬇️ Download Billing", f"gig_billing_{datetime.now().strftime('%Y%m%d')}.xlsx")
