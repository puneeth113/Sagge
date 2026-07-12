import os
import importlib.util
from datetime import date, timedelta

import streamlit as st
import pandas as pd


def _load_module(module_filename, module_name):
    this_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(this_dir)
    for candidate in (os.path.join(this_dir, module_filename), os.path.join(root_dir, module_filename)):
        if os.path.exists(candidate):
            spec = importlib.util.spec_from_file_location(module_name, candidate)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise FileNotFoundError(f"Could not find {module_filename}.")


def _load_utils():
    this_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(this_dir)
    for candidate in (
        os.path.join(this_dir, "utils.py"), os.path.join(this_dir, "Utils.py"),
        os.path.join(root_dir, "utils.py"), os.path.join(root_dir, "Utils.py"),
    ):
        if os.path.exists(candidate):
            spec = importlib.util.spec_from_file_location("hr_utils", candidate)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise FileNotFoundError("Could not find utils.py.")


_u = _load_utils()
render_top_nav = _u.render_top_nav
read_any_table = _u.read_any_table
download_button_for_df = _u.download_button_for_df
to_excel_bytes = _u.to_excel_bytes
safe_error_message = _u.safe_error_message

mat = _load_module("maternity_utils.py", "hr_maternity")
compute_maternity_payment = mat.compute_maternity_payment
parse_excluded_dates = mat.parse_excluded_dates

st.set_page_config(page_title="Maternity Payment", page_icon="🤱", layout="wide")
render_top_nav("Maternity Payment")

st.title("🤱 Maternity Payment Calculator")
st.caption(
    "Enter the leave's Start Date and End Date. The payable span is capped at the entitlement "
    "(182 days by default) even if the requested end date runs longer. Any already-paid days that "
    "fall inside that span are subtracted before the flat per-day rate (monthly salary ÷ divisor) is applied."
)


def _show_error(e: Exception, context: str):
    if isinstance(e, ValueError):
        st.error(str(e))
    else:
        st.error(safe_error_message(e, context=context))


tab_single, tab_bulk = st.tabs(["✍️ Single Employee", "📤 Bulk Upload"])

with tab_single:
    c1, c2, c3 = st.columns(3)
    with c1:
        monthly_salary = st.number_input("Monthly Salary (₹)", min_value=0.0, value=25000.0, step=500.0)
        leave_start = st.date_input("Leave Start Date", value=date.today())
    with c2:
        leave_end = st.date_input("Leave End Date", value=date.today() + timedelta(days=181))
        max_entitled = st.number_input(
            "Max Entitled Days (cap)", min_value=1, value=182, step=1,
            help="182 = standard 26-week entitlement. Use 84 for the 12-week third-child case, or any figure that applies. The payable span is capped here even if the leave dates span longer.",
        )
    with c3:
        daily_divisor = st.number_input(
            "Daily Rate Divisor", min_value=1, value=30, step=1,
            help="Daily rate = Monthly Salary ÷ this number. 30 is the common 'monthly basis' convention; some orgs use 26.",
        )

    st.markdown("#### Already-paid days to exclude (optional)")
    st.caption(
        "If any days inside the leave window were already paid separately (e.g. overlapping paid leave), "
        "list them here — one per line, either a single date (`2026-02-05`) or a range (`2026-02-01:2026-02-10`). "
        "Only dates that fall inside the (possibly capped) payable window are subtracted. Leave blank if none."
    )
    excluded_text = st.text_area("Excluded dates", value="", height=100, key="single_excluded")

    if st.button("Calculate Maternity Payment", key="mat_calc_single"):
        try:
            excluded_dates = parse_excluded_dates(excluded_text)
            result = compute_maternity_payment(
                monthly_salary, leave_start, leave_end,
                max_entitled_days=int(max_entitled),
                already_paid_dates=excluded_dates,
                daily_rate_divisor=int(daily_divisor),
            )
        except Exception as e:
            _show_error(e, "calculating the maternity payment")
        else:
            if result["was_capped"]:
                st.warning(
                    f"⚠️ Requested span was {result['actual_days_requested']} days, which exceeds the "
                    f"{result['max_entitled_days']}-day cap. Payable window capped to "
                    f"{result['leave_start_date']} → {result['effective_end_date']}."
                )

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Requested Span", f"{result['actual_days_requested']} days")
            m2.metric("Payable Days", result["payable_days"])
            m3.metric("Daily Rate", f"₹{result['daily_rate']:,.2f}")
            m4.metric("Total Payable Amount", f"₹{result['total_amount']:,.2f}")

            st.caption(
                f"Leave window: {result['leave_start_date']} → {result['leave_end_date']}"
                + (f" (effective payable end: {result['effective_end_date']})" if result["was_capped"] else "")
                + f" | Decremented (already-paid) days: {result['decremented_days']}"
            )

with tab_bulk:
    st.caption(
        "Columns required: **Employee Name, Monthly Salary, Leave Start Date, Leave End Date**. "
        "Optional columns: **Max Entitled Days** (defaults to 182), **Daily Rate Divisor** (defaults to 30), "
        "**Excluded Dates** (multiple entries separated by `;`, e.g. `2026-02-05;2026-03-01:2026-03-05`)."
    )
    sample = pd.DataFrame([
        {"Employee Name": "Anjali Rao", "Monthly Salary": 25000, "Leave Start Date": "2026-01-01",
         "Leave End Date": "2026-07-01", "Max Entitled Days": 182, "Daily Rate Divisor": 30, "Excluded Dates": ""},
        {"Employee Name": "Fatima Sheikh", "Monthly Salary": 30000, "Leave Start Date": "2026-02-01",
         "Leave End Date": "2026-08-01", "Max Entitled Days": 182, "Daily Rate Divisor": 30,
         "Excluded Dates": "2026-03-01:2026-03-05"},
    ])
    st.dataframe(sample, use_container_width=True)
    st.download_button(
        "⬇️ Download sample template",
        data=to_excel_bytes({"Template": sample}),
        file_name="maternity_bulk_template.xlsx",
    )

    uploaded = st.file_uploader("Upload bulk maternity sheet (.xlsx or .csv)", type=["xlsx", "xls", "csv"])

    if uploaded:
        df = read_any_table(uploaded)
        st.dataframe(df, use_container_width=True)

        required = {"Employee Name", "Monthly Salary", "Leave Start Date", "Leave End Date"}
        missing = required - set(df.columns)
        if missing:
            st.error(f"Missing required column(s): {', '.join(missing)}")
        elif st.button("Calculate for all rows"):
            results_rows, errors = [], []
            for idx, row in df.iterrows():
                try:
                    excl_raw = str(row.get("Excluded Dates", "") or "").replace(";", "\n")
                    excl_dates = parse_excluded_dates(excl_raw)
                    leave_start_row = pd.to_datetime(row["Leave Start Date"]).date()
                    leave_end_row = pd.to_datetime(row["Leave End Date"]).date()
                    max_days_row = int(row["Max Entitled Days"]) if "Max Entitled Days" in df.columns and pd.notna(row.get("Max Entitled Days")) else 182
                    divisor_row = int(row["Daily Rate Divisor"]) if "Daily Rate Divisor" in df.columns and pd.notna(row.get("Daily Rate Divisor")) else 30

                    res = compute_maternity_payment(
                        float(row["Monthly Salary"]), leave_start_row, leave_end_row,
                        max_entitled_days=max_days_row,
                        already_paid_dates=excl_dates,
                        daily_rate_divisor=divisor_row,
                    )
                    results_rows.append({
                        "Employee Name": row["Employee Name"],
                        "Leave Start": res["leave_start_date"],
                        "Leave End (Requested)": res["leave_end_date"],
                        "Requested Span (days)": res["actual_days_requested"],
                        "Capped?": res["was_capped"],
                        "Effective Payable End": res["effective_end_date"],
                        "Decremented Days": res["decremented_days"],
                        "Payable Days": res["payable_days"],
                        "Daily Rate (₹)": res["daily_rate"],
                        "Total Amount (₹)": res["total_amount"],
                    })
                except Exception as e:
                    label = str(e) if isinstance(e, ValueError) else safe_error_message(e)
                    errors.append(f"Row {idx + 2}: {label}")

            if results_rows:
                results_df = pd.DataFrame(results_rows)
                st.success(f"Calculated {len(results_df)} of {len(df)} row(s).")
                st.dataframe(results_df, use_container_width=True, height=400)
                download_button_for_df(results_df, "⬇️ Download results", "maternity_bulk_results.xlsx")
            if errors:
                st.warning("Some rows had issues:")
                for err in errors:
                    st.caption(f"- {err}")
