import os
import importlib.util

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
read_any_table = _u.read_any_table
to_excel_bytes = _u.to_excel_bytes
download_button_for_df = _u.download_button_for_df
sample_incentive_template = _u.sample_incentive_template
compute_incentives = _u.compute_incentives
INCENTIVE_TEMPLATE_COLUMNS = _u.INCENTIVE_TEMPLATE_COLUMNS
render_top_nav = _u.render_top_nav

st.set_page_config(page_title="Incentive Validator", page_icon="💰", layout="wide")
render_top_nav("Incentive Validator")
st.title("💰 Incentive Validator")
st.caption("Bulk-upload branch-wise maid counts and incentive caps to validate and compute incentives.")

tab_bulk, tab_manual = st.tabs(["📤 Bulk Upload", "✍️ Single Branch (manual)"])

with tab_bulk:
    st.markdown("#### Step 1 — Download the sample sheet")
    st.caption(
        "Required columns: **Branch Name, Bus Maid Count, Washroom Maid Count, "
        "Max Washroom Incentive, Max Bus Incentive**."
    )
    sample_df = sample_incentive_template()
    st.dataframe(sample_df, use_container_width=True)
    st.download_button(
        "⬇️ Download sample template (.xlsx)",
        data=to_excel_bytes({"Incentive Upload Template": sample_df}),
        file_name="incentive_bulk_upload_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.markdown("#### Step 2 — Set incentive rate per maid")
    st.info(
        "⚠️ **Assumption to confirm:** since the exact incentive formula wasn't specified, "
        "this tool calculates `Incentive = Maid Count × Rate per Maid`, capped at the branch's "
        "Max Incentive value. Adjust the rates below to match your actual policy, or edit "
        "`compute_incentives()` in `utils.py` if the logic should be slab-based instead of linear."
    )
    rc1, rc2 = st.columns(2)
    with rc1:
        rate_washroom = st.number_input("Rate per Washroom Maid (₹)", min_value=0.0, value=500.0, step=50.0)
    with rc2:
        rate_bus = st.number_input("Rate per Bus Maid (₹)", min_value=0.0, value=500.0, step=50.0)

    st.markdown("#### Step 3 — Upload your branch data")
    uploaded = st.file_uploader("Upload filled incentive sheet (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="incentive_upload")

    if uploaded:
        df = read_any_table(uploaded)

        missing = [c for c in INCENTIVE_TEMPLATE_COLUMNS if c not in df.columns]
        if missing:
            st.error(f"Uploaded sheet is missing required column(s): {', '.join(missing)}. "
                      "Please match the sample template's headers exactly.")
        else:
            numeric_cols = ["Bus Maid Count", "Washroom Maid Count", "Max Washroom Incentive", "Max Bus Incentive"]
            bad_rows = df[df[numeric_cols].apply(pd.to_numeric, errors="coerce").isna().any(axis=1)]
            if not bad_rows.empty:
                st.warning(f"{len(bad_rows)} row(s) have non-numeric values in count/incentive columns and were excluded.")
                df = df.drop(bad_rows.index)

            for c in numeric_cols:
                df[c] = pd.to_numeric(df[c], errors="coerce")

            results = compute_incentives(df, rate_washroom, rate_bus)

            st.success(f"Computed incentives for {len(results)} branch(es).")

            capped_count = int((results["Capped? (Washroom)"] | results["Capped? (Bus)"]).sum())
            m1, m2, m3 = st.columns(3)
            m1.metric("Branches Processed", len(results))
            m2.metric("Total Incentive Payout (₹)", f"{results['Total Incentive'].sum():,.0f}")
            m3.metric("Branches Hitting Cap", capped_count)

            st.dataframe(results, use_container_width=True, height=420)
            download_button_for_df(results, "⬇️ Download computed incentives", "incentive_results.xlsx")

with tab_manual:
    st.markdown("Quickly validate incentive for a single branch without uploading a file.")
    c1, c2 = st.columns(2)
    with c1:
        branch = st.text_input("Branch Name", "")
        bus_count = st.number_input("Bus Maid Count", min_value=0, value=0)
        washroom_count = st.number_input("Washroom Maid Count", min_value=0, value=0)
    with c2:
        max_washroom = st.number_input("Max Washroom Incentive (₹)", min_value=0.0, value=3000.0)
        max_bus = st.number_input("Max Bus Incentive (₹)", min_value=0.0, value=2000.0)
        rate_washroom_m = st.number_input("Rate per Washroom Maid (₹)", min_value=0.0, value=500.0, key="rwm")
        rate_bus_m = st.number_input("Rate per Bus Maid (₹)", min_value=0.0, value=500.0, key="rbm")

    if st.button("Calculate"):
        single = pd.DataFrame([{
            "Branch Name": branch or "Branch",
            "Bus Maid Count": bus_count,
            "Washroom Maid Count": washroom_count,
            "Max Washroom Incentive": max_washroom,
            "Max Bus Incentive": max_bus,
        }])
        res = compute_incentives(single, rate_washroom_m, rate_bus_m)
        st.dataframe(res, use_container_width=True)
        st.metric("Total Incentive", f"₹{res['Total Incentive'].iloc[0]:,.0f}")
