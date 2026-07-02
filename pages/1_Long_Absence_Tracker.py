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
count_absences = _u.count_absences
categorize_absences = _u.categorize_absences
DEFAULT_ABSENCE_THRESHOLDS = _u.DEFAULT_ABSENCE_THRESHOLDS
to_excel_bytes = _u.to_excel_bytes
render_top_nav = _u.render_top_nav
sample_attendance_muster = _u.sample_attendance_muster

st.set_page_config(page_title="Long Absence Tracker", page_icon="📅", layout="wide")
render_top_nav("Long Absence Tracker")
st.title("📅 Long Absence Tracker")
st.caption(
    "Upload your attendance muster (employees as rows, dates as columns, "
    "attendance marks like P/A/L/WO in each cell). This tool counts every "
    "'A' per employee and buckets them into risk categories."
)

with st.expander("📥 Download a sample muster template", expanded=False):
    st.caption(
        "Not sure how to structure your file? Here's the expected shape — "
        "identifier columns first (Employee ID, Employee Name, Branch), "
        "then one column per day with P/A/L/WO marks."
    )
    sample_df = sample_attendance_muster()
    st.dataframe(sample_df, use_container_width=True)
    st.download_button(
        "⬇️ Download sample template (.xlsx)",
        data=to_excel_bytes({"Sample Muster": sample_df}),
        file_name="attendance_muster_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


uploaded = st.file_uploader("Upload Attendance Muster (.xlsx or .csv)", type=["xlsx", "xls", "csv"])

if not uploaded:
    st.info("Upload a muster file to begin. Expected shape: one row per employee, "
             "one column per day, with an ID/Name column (or columns) at the start.")
    st.stop()

raw_df = read_any_table(uploaded)
st.markdown("#### Preview of uploaded file")
st.dataframe(raw_df.head(10), use_container_width=True)

st.markdown("#### Step 1 — Identify employee columns")
st.caption("Select the column(s) that identify an employee (e.g. Employee ID, Employee Name, Branch). "
           "Every other column will be treated as a day-wise attendance column and scanned for 'A'.")
id_cols = st.multiselect(
    "Employee identifier column(s)",
    options=list(raw_df.columns),
    default=list(raw_df.columns[:2]) if len(raw_df.columns) >= 2 else list(raw_df.columns[:1]),
)

absence_marker = st.text_input("Marker used for 'Absent' in your muster", value="A")

if not id_cols:
    st.warning("Select at least one identifier column to continue.")
    st.stop()

st.markdown("#### Step 2 — Set category thresholds")
st.caption("Default buckets follow: 3-4 days = Informed Leave, 5-20 days = To Be Checked, "
           ">20 days = Probable Exit Case. Adjust below if needed.")

use_custom = st.checkbox("Customize thresholds", value=False)
if use_custom:
    with st.expander("Adjust thresholds", expanded=True):
        t1 = st.slider("Informed Leave range (days)", 1, 30, (3, 4))
        t2 = st.slider("To Be Checked range (days)", 1, 60, (5, 20))
        t3_min = st.number_input("Probable Exit Case — minimum days (and above)", min_value=1, value=21)

        thresholds = [
            (f"Informed Leave ({t1[0]}-{t1[1]} days)", t1[0], t1[1]),
            (f"To Be Checked ({t2[0]}-{t2[1]} days)", t2[0], t2[1]),
            (f"Probable Exit Case (>{t3_min - 1} days)", t3_min, None),
        ]
else:
    thresholds = DEFAULT_ABSENCE_THRESHOLDS


result = count_absences(raw_df, id_cols, absence_marker=absence_marker)
result = categorize_absences(result, thresholds)

st.divider()
st.markdown("### Results")

summary = result["Absence_Category"].value_counts().reindex(
    [t[0] for t in thresholds] + ["Not Flagged"], fill_value=0
)
cols = st.columns(len(summary))
for c, (label, count) in zip(cols, summary.items()):
    c.metric(label, int(count))

st.markdown("#### Full breakdown")
category_filter = st.selectbox("Filter by category", ["(all)"] + list(summary.index))
view = result if category_filter == "(all)" else result[result["Absence_Category"] == category_filter]
view_sorted = view.sort_values("Total_Absent_Days", ascending=False)
st.dataframe(view_sorted, use_container_width=True, height=420)

st.divider()
st.markdown("#### Download report")
sheets = {"All Employees": result}
for label, _, _ in thresholds:
    sheets[label[:31]] = result[result["Absence_Category"] == label]

st.download_button(
    "⬇️ Download full report (multi-sheet Excel)",
    data=to_excel_bytes(sheets),
    file_name="long_absence_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
