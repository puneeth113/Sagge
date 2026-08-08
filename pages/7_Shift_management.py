import os
import importlib.util
from datetime import datetime, date, timedelta, time

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
safe_error_message = _u.safe_error_message


st.set_page_config(page_title="Shift Management", page_icon="🕒", layout="wide")
render_top_nav("Shift Management")

st.title("🕒 Shift Management")
st.caption("Category master and bulk shift operations. Use the masters to configure categories and shifts, then run bulk processing to validate and update assignments.")


# Default working hours per category (HH:MM)
DEFAULT_WORKING_HOURS = {
    "PP": "6:15",  # Preprimary
    "APP": "6:45",  # Above Primary
    "L": "8:15",  # Leader
}

# fixed buffer rule (15 minutes) — not exposed in the UI
BUFFER_MINUTES = 15


def parse_duration_to_timedelta(s: str) -> timedelta:
    if pd.isna(s):
        return None
    if isinstance(s, (int, float)):
        return timedelta(minutes=int(s))
    text = str(s).strip()
    if not text:
        return None
    try:
        parts = text.split(":")
        if len(parts) == 1:
            hours = int(parts[0])
            minutes = 0
        else:
            hours = int(parts[0])
            minutes = int(parts[1])
        return timedelta(hours=hours, minutes=minutes)
    except Exception:
        try:
            td = pd.to_timedelta(text)
            return td
        except Exception:
            return None


def format_timedelta(td: timedelta) -> str:
    if td is None:
        return ""
    total_seconds = int(td.total_seconds())
    if total_seconds < 0:
        sign = "-"
        total_seconds = abs(total_seconds)
    else:
        sign = ""
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{sign}{hours}:{minutes:02d}"


def _parse_time_to_time(v) -> time:
    if pd.isna(v):
        return None
    if isinstance(v, time):
        return v
    if isinstance(v, datetime):
        return v.time()
    try:
        import pandas as _pd
        if isinstance(v, _pd.Timestamp):
            return v.time()
    except Exception:
        pass
    text = str(v).strip()
    if not text:
        return None
    # try common formats
    for fmt in ("%H:%M", "%I:%M %p", "%H:%M:%S", "%I %p", "%H%M"):
        try:
            return datetime.strptime(text, fmt).time()
        except Exception:
            continue
    # try pandas parsing
    try:
        ts = pd.to_datetime(text, errors="coerce")
        if pd.notna(ts):
            return ts.time()
    except Exception:
        pass
    # excel fractional day
    try:
        f = float(text)
        if 0.0 <= f < 1.0:
            total_seconds = int(f * 24 * 3600)
            h = total_seconds // 3600
            m = (total_seconds % 3600) // 60
            return time(h, m)
    except Exception:
        pass
    return None


def split_assigned_shift(s: str):
    """Split 'HH:MM-HH:MM' into (start_time, end_time) as datetime.time objects.
    Returns (None, None) if parsing fails.
    """
    if pd.isna(s):
        return None, None
    text = str(s).strip()
    if not text:
        return None, None
    if "-" in text:
        a, b = text.split("-", 1)
    elif " to " in text.lower():
        a, b = text.lower().split(" to ", 1)
    else:
        return None, None
    ta = _parse_time_to_time(a.strip())
    tb = _parse_time_to_time(b.strip())
    return ta, tb


# --- Helpers for storing masters in session state ---------------------------------


def _ensure_masters():
    if "shift_categories" not in st.session_state:
        # dict: code -> HH:MM string
        st.session_state["shift_categories"] = DEFAULT_WORKING_HOURS.copy()
    if "named_shifts" not in st.session_state:
        # dict: shift_name -> 'HH:MM-HH:MM'
        st.session_state["named_shifts"] = {
            "Standard A": "08:00-16:15",
            "Standard B": "07:45-16:00",
        }


_ensure_masters()


# Top-level sub-navigation for the two required pages
page = st.radio("Shift Management — Section", options=["Category & Shift Master", "Bulk Shift Operations"], horizontal=True)

if page == "Category & Shift Master":
    st.markdown("### Category & Shift Master")
    st.caption("Create / edit categories and named shifts used by bulk operations.")

    col1, col2 = st.columns([2, 1])
    with col1:
        st.subheader("Categories (code -> default working hours HH:MM)")
        cat_df = pd.DataFrame([{"Category": k, "Default Hours": v} for k, v in st.session_state["shift_categories"].items()])
        # use data_editor when available for inline edits, otherwise show dataframe
        if hasattr(st, "data_editor"):
            _ = st.data_editor(cat_df, use_container_width=True)
        else:
            st.dataframe(cat_df)

        st.markdown("**Add new category**")
        new_code = st.text_input("Category code (e.g. G1)", value="", key="new_cat_code")
        new_hours = st.text_input("Default hours (HH:MM)", value="6:00", key="new_cat_hours")
        if st.button("Add category"):
            if not new_code.strip():
                st.error("Provide a category code.")
            else:
                st.session_state["shift_categories"][new_code.strip()] = new_hours.strip()
                st.success(f"Added category {new_code.strip()} -> {new_hours.strip()}")
                st.experimental_rerun()

        st.markdown("**Edit existing category**")
        sel_cat = st.selectbox("Select category to edit", options=list(st.session_state["shift_categories"].keys()))
        if sel_cat:
            cur = st.session_state["shift_categories"][sel_cat]
            new_val = st.text_input("Default hours (HH:MM)", value=cur, key=f"edit_cat_{sel_cat}")
            if st.button("Save category", key=f"save_cat_{sel_cat}"):
                st.session_state["shift_categories"][sel_cat] = new_val.strip()
                st.success(f"Saved {sel_cat} -> {new_val.strip()}")
                st.experimental_rerun()

        if st.button("Delete selected category"):
            if sel_cat in st.session_state["shift_categories"]:
                del st.session_state["shift_categories"][sel_cat]
                st.success(f"Deleted category {sel_cat}")
                st.experimental_rerun()

    with col2:
        st.subheader("Named Shifts (label -> HH:MM-HH:MM)")
        nsh = st.session_state["named_shifts"]
        nsh_df = pd.DataFrame([{"Shift": k, "Assigned": v} for k, v in nsh.items()])
        st.dataframe(nsh_df, use_container_width=True)

        st.markdown("**Add new named shift**")
        new_shift_name = st.text_input("Shift name", value="", key="new_shift_name")
        new_shift_span = st.text_input("Shift span (HH:MM-HH:MM)", value="08:00-16:00", key="new_shift_span")
        if st.button("Add shift"):
            if not new_shift_name.strip():
                st.error("Provide a shift name.")
            else:
                st.session_state["named_shifts"][new_shift_name.strip()] = new_shift_span.strip()
                st.success(f"Added shift {new_shift_name.strip()} -> {new_shift_span.strip()}")
                st.experimental_rerun()

        st.markdown("**Edit existing shift**")
        sel_shift = st.selectbox("Select shift to edit", options=list(nsh.keys()))
        if sel_shift:
            cur = st.session_state["named_shifts"][sel_shift]
            new_span = st.text_input("Shift span (HH:MM-HH:MM)", value=cur, key=f"edit_shift_{sel_shift}")
            if st.button("Save shift", key=f"save_shift_{sel_shift}"):
                st.session_state["named_shifts"][sel_shift] = new_span.strip()
                st.success(f"Saved {sel_shift} -> {new_span.strip()}")
                st.experimental_rerun()

        if st.button("Delete selected shift"):
            if sel_shift in st.session_state["named_shifts"]:
                del st.session_state["named_shifts"][sel_shift]
                st.success(f"Deleted shift {sel_shift}")
                st.experimental_rerun()


else:
    st.markdown("### Bulk Shift Operations")
    st.caption("Download the sample template, fill it, upload the file, then click Process to validate and generate the final result. No additional configuration is required on this page.")

    # Sample template generator
    def sample_bulk_template():
        df = pd.DataFrame([
            {"ERP ID": "12345678901", "Branch Name": "OIS Sample Branch", "Category": "Teacher", "First Bell Timing": "08:00", "Current Assigned Shift": "07:45-16:00"},
            {"ERP ID": "12345678902", "Branch Name": "OIS Sample Branch", "Category": "Teacher", "First Bell Timing": "09:00", "Current Assigned Shift": "08:45-17:00"},
            {"ERP ID": "12345678903", "Branch Name": "OIS Sample Branch", "Category": "Non-Teaching", "First Bell Timing": "08:30", "Current Assigned Shift": "08:15-16:30"},
        ])
        instr = pd.DataFrame({"Instructions": [
            "Fill the 'ERP ID', 'Branch Name', 'Category', 'First Bell Timing' (HH:MM), and 'Current Assigned Shift' (HH:MM-HH:MM).",
            "Do not change the column headers. The system auto-detects these columns case-insensitively.",
            "First Bell Timing is used to compute the expected start (First Bell - 15 minutes) and expected end based on the category working hours.",
        ]})
        return {"Sample": df, "Instructions": instr}

    if st.button("Download Sample File"):
        try:
            sheets = sample_bulk_template()
            b = to_excel_bytes(sheets)
            st.download_button("⬇️ Download sample bulk file", data=b, file_name="shift_bulk_sample.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(safe_error_message(e, context="creating sample file"))

    uploaded = st.file_uploader("Upload completed bulk file (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="bulk_shifts")
    if uploaded:
        try:
            df = read_any_table(uploaded)
            st.markdown("#### Preview of uploaded data (first 10 rows)")
            st.dataframe(df.head(10), use_container_width=True)

            def auto_detect_columns(df: pd.DataFrame):
                cols_lower = {c.lower(): c for c in df.columns}
                mapping = {}
                # ERP ID
                for candidate in ["erp id", "erpid", "employee id", "employee_id", "id", "employeeid"]:
                    if candidate in cols_lower:
                        mapping["erp_id"] = cols_lower[candidate]
                        break
                # Branch
                for candidate in ["branch name", "branch", "branch_name"]:
                    if candidate in cols_lower:
                        mapping["branch"] = cols_lower[candidate]
                        break
                # Category
                for candidate in ["category", "cat"]:
                    if candidate in cols_lower:
                        mapping["category"] = cols_lower[candidate]
                        break
                # First Bell Timing
                for candidate in ["first bell timing", "first bell", "first_bell_timing", "first_bell"]:
                    if candidate in cols_lower:
                        mapping["first_bell"] = cols_lower[candidate]
                        break
                # Assigned shift
                for candidate in ["current assigned shift", "assigned shift", "assigned_shift", "shift", "current shift", "current_assigned_shift"]:
                    if candidate in cols_lower:
                        mapping["assigned"] = cols_lower[candidate]
                        break
                return mapping

            mapping = auto_detect_columns(df)
            missing = [k for k in ("erp_id", "branch", "category", "first_bell", "assigned") if k not in mapping]
            if missing:
                st.error(f"Missing required columns (auto-detect failed): {missing}. Please ensure the uploaded sheet has these columns with recognizable names.")
            else:
                if st.button("Process / Validate"):
                    out_rows = []
                    for i, row in df.iterrows():
                        erp_val = row[mapping["erp_id"]]
                        erp_str = str(erp_val).strip() if pd.notna(erp_val) else ""
                        branch = row[mapping["branch"]]
                        category = row[mapping["category"]]
                        first_bell_raw = row[mapping["first_bell"]]
                        fb_time = _parse_time_to_time(first_bell_raw)
                        assigned = row[mapping["assigned"]]

                        a_start, a_end = split_assigned_shift(assigned)
                        current_working = None
                        if a_start and a_end:
                            dt_start = datetime.combine(date.today(), a_start)
                            dt_end = datetime.combine(date.today(), a_end)
                            if dt_end < dt_start:
                                dt_end += timedelta(days=1)
                            current_working = dt_end - dt_start

                        # System expected start = first_bell - BUFFER_MINUTES
                        expected_start_time = None
                        expected_end_time = None
                        expected_working = None
                        if fb_time:
                            expected_start_dt = datetime.combine(date.today(), fb_time) - timedelta(minutes=BUFFER_MINUTES)
                            expected_start_time = expected_start_dt.time()
                            # duration from category master
                            cat_hours = st.session_state["shift_categories"].get(str(category).strip(), None)
                            if cat_hours is None:
                                cat_hours = DEFAULT_WORKING_HOURS.get(str(category).strip(), None)
                            duration_td = parse_duration_to_timedelta(cat_hours) if cat_hours else None
                            if duration_td is not None:
                                ee_dt = expected_start_dt + duration_td
                                expected_end_time = ee_dt.time()
                                expected_working = duration_td

                        # compare
                        status = "To Be Checked"
                        notes = []
                        if a_start is None or a_end is None:
                            status = "To Be Checked"
                            notes.append("Bad Assigned Shift Format")
                        elif fb_time is None:
                            status = "To Be Checked"
                            notes.append("Missing/invalid First Bell Timing")
                        elif expected_working is None:
                            status = "To Be Checked"
                            notes.append("No duration configured for category")
                        else:
                            # compute diffs
                            diff_start = (datetime.combine(date.today(), a_start) - datetime.combine(date.today(), expected_start_time)).total_seconds() / 60.0
                            diff_end = (datetime.combine(date.today(), a_end) - datetime.combine(date.today(), expected_end_time)).total_seconds() / 60.0
                            tol_minutes = 5
                            if abs(diff_start) <= tol_minutes and abs(diff_end) <= tol_minutes:
                                status = "Shift Matching"
                            else:
                                status = "To Be Checked"
                                notes.append(f"Start Δ {diff_start:.0f} min, End Δ {diff_end:.0f} min")

                        out_rows.append({
                            "ERP ID": erp_str,
                            "Branch Name": branch,
                            "Category": category,
                            "First Bell Timing": fb_time.strftime("%H:%M") if fb_time else "",
                            "Current Assigned Shift": str(assigned),
                            "Current Start Time": a_start.strftime("%H:%M") if a_start else "",
                            "Current End Time": a_end.strftime("%H:%M") if a_end else "",
                            "Current Working Hours": format_timedelta(current_working) if current_working else "",
                            "System Calculated Shift": (f"{expected_start_time.strftime('%H:%M')}-{expected_end_time.strftime('%H:%M')}" if expected_start_time and expected_end_time else ""),
                            "System Start Time": expected_start_time.strftime("%H:%M") if expected_start_time else "",
                            "System End Time": expected_end_time.strftime("%H:%M") if expected_end_time else "",
                            "System Working Hours": format_timedelta(expected_working) if expected_working else "",
                            "Remarks": "; ".join(notes) if notes else status,
                        })

                    out_df = pd.DataFrame(out_rows)
                    st.markdown("#### Final Result")
                    st.dataframe(out_df, use_container_width=True, height=600)
                    # show counts
                    st.markdown(f"**Total rows:** {len(out_df)}; **To Be Checked:** {len(out_df[out_df['Remarks'] != 'Shift Matching'])}")

                    # final download
                    try:
                        b = to_excel_bytes({"Processed_Result": out_df})
                        st.download_button("⬇️ Download final processed file", data=b, file_name="shift_bulk_processed.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e:
                        st.error(safe_error_message(e, context="creating final download"))

        except Exception as e:
            st.error(safe_error_message(e, context="processing bulk upload"))


