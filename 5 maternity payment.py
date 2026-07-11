import os
import importlib.util
from datetime import date

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
safe_error_message = _u.safe_error_message

mat = _load_module("maternity_utils.py", "hr_maternity")
compute_maternity_payment = mat.compute_maternity_payment
parse_excluded_dates = mat.parse_excluded_dates

st.set_page_config(page_title="Maternity Payment", page_icon="🤱", layout="wide")
render_top_nav("Maternity Payment")

st.title("🤱 Maternity Payment Calculator")
st.caption(
    "Calculates maternity benefit on a per-day basis using each calendar month's own day count "
    "(so a February day and a January day are valued slightly differently for the same monthly salary). "
    "Any dates already paid through another route within the leave window are excluded automatically."
)

tab_single, tab_bulk = st.tabs(["✍️ Single Employee", "📤 Bulk Upload"])

with tab_single:
    c1, c2 = st.columns(2)
    with c1:
        monthly_salary = st.number_input("Monthly Salary (₹)", min_value=0.0, value=25000.0, step=500.0)
        leave_start = st.date_input("Leave Start Date", value=date.today())
    with c2:
        entitled_days = st.number_input(
            "Total Entitled Days", min_value=1, value=182, step=1,
            help="182 days = standard 26-week entitlement. Use 84 for the 12-week third-child case, or any other figure that applies.",
        )

    st.markdown("#### Already-paid days to exclude (optional)")
    st.caption(
        "If any days inside the leave window were already paid separately (e.g. overlapping paid leave), "
        "list them here — one per line, either a single date (`2026-02-05`) or a range (`2026-02-01:2026-02-10`). "
        "Leave blank if none."
    )
    excluded_text = st.text_area("Excluded dates", value="", height=100, key="single_excluded")

    if st.button("Calculate Maternity Payment", key="mat_calc_single"):
        try:
            excluded_dates = parse_excluded_dates(excluded_text)
        except ValueError as e:
            st.error(str(e))
        else:
            try:
                result = compute_maternity_payment(monthly_salary, leave_start, int(entitled_days), excluded_dates)
            except Exception as e:
                st.error(safe_error_message(e, context="calculating the maternity payment"))
            else:
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Leave Window", f"{result['leave_start_date']} → {result['leave_end_date']}")
                m2.metric("Entitled Days", result["total_entitled_days"])
                m3.metric("Decremented (already paid)", result["decremented_days"])
                m4.metric("Total Payable Amount", f"₹{result['total_amount']:,.2f}")

                st.markdown("#### Month-by-month breakdown")
                breakdown_df = pd.DataFrame(result["monthly_breakdown"])
                st.dataframe(breakdown_df, use_container_width=True)
                download_button_for_df(breakdown_df, "⬇️ Download breakdown", "maternity_payment_breakdown.xlsx")

with tab_bulk:
    st.caption(
        "Columns required: **Employee Name, Monthly Salary, Leave Start Date, Entitled Days**. "
        "Optional column: **Excluded Dates** — same format as above (multiple entries separated by `;`)."
    )
    sample = pd.DataFrame([
        {"Employee Name": "Anjali Rao", "Monthly Salary": 25000, "Leave Start Date": "2026-01-01",
         "Entitled Days": 182, "Excluded Dates": ""},
        {"Employee Name": "Fatima Sheikh", "Monthly Salary": 30000, "Leave Start Date": "2026-02-01",
         "Entitled Days": 182, "Excluded Dates": "2026-03-01:2026-03-05"},
    ])
    st.download_button(
        "⬇️ Download sample template",
        data=_u.to_excel_bytes({"Template": sample}),
        file_name="maternity_bulk_template.xlsx",
    )

    uploaded = st.file_uploader("Upload bulk maternity sheet (.xlsx or .csv)", type=["xlsx", "xls", "csv"])

    if uploaded:
        df = read_any_table(uploaded)
        st.dataframe(df, use_container_width=True)

        required = {"Employee Name", "Monthly Salary", "Leave Start Date", "Entitled Days"}
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
                    res = compute_maternity_payment(
                        float(row["Monthly Salary"]), leave_start_row,
                        int(row["Entitled Days"]), excl_dates,
                    )
                    results_rows.append({
                        "Employee Name": row["Employee Name"],
                        "Leave Start": res["leave_start_date"],
                        "Leave End": res["leave_end_date"],
                        "Entitled Days": res["total_entitled_days"],
                        "Decremented Days": res["decremented_days"],
                        "Payable Days": res["payable_days"],
                        "Total Amount (₹)": res["total_amount"],
                    })
                except Exception as e:
                    errors.append(f"Row {idx + 2}: {safe_error_message(e)}")

            if results_rows:
                results_df = pd.DataFrame(results_rows)
                st.success(f"Calculated {len(results_df)} of {len(df)} row(s).")
                st.dataframe(results_df, use_container_width=True, height=400)
                download_button_for_df(results_df, "⬇️ Download results", "maternity_bulk_results.xlsx")
            if errors:
                st.warning("Some rows had issues:")
                for err in errors:
                    st.caption(f"- {err}")
