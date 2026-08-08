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
    # try pandas
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
    elif " to " in text:
        a, b = text.split(" to ", 1)
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
        edited = st.data_editor(cat_df, use_container_width=True) if hasattr(st, "data_editor") else st.dataframe(cat_df)

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
        edited_shifts = st.dataframe(nsh_df, use_container_width=True)

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
    st.caption("Upload bulk assignments and validate/calculate shifts. The system will auto-detect required columns by name; no manual mapping is required.")

    st.markdown("#### Settings")
    col_a, col_b = st.columns(2)
    with col_a:
        first_bell = st.time_input("Default First Bell time (used to compute expected start)", value=time(8, 0), key="bulk_first_bell")
        buffer_minutes = st.number_input("Buffer minutes before first bell", min_value=0, value=BUFFER_MINUTES, key="bulk_buffer")
    with col_b:
        st.info("Upload must contain ERP ID, Branch Name, Category and Current Assigned Shift (HH:MM-HH:MM). Column names are matched case-insensitively.")

    uploaded = st.file_uploader("Upload bulk assignments (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="bulk_shifts")

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
        # Assigned shift
        for candidate in ["current assigned shift", "assigned shift", "assigned_shift", "shift", "current shift", "current_assigned_shift"]:
            if candidate in cols_lower:
                mapping["assigned"] = cols_lower[candidate]
                break
        return mapping

    if uploaded:
        try:
            df = read_any_table(uploaded)
            st.write(f"Loaded {len(df)} rows. Detected columns: {list(df.columns)}")
            mapping = auto_detect_columns(df)
            missing = [k for k in ("erp_id", "branch", "category", "assigned") if k not in mapping]
            if missing:
                st.error(f"Missing required columns (auto-detect failed): {missing}. Please ensure the uploaded sheet has these columns with recognizable names.")
            else:
                # validate ERP ID and parse assigned shift
                out_rows = []
                for i, row in df.iterrows():
                    raw_erp = row[mapping["erp_id"]]
                    erp_str = str(raw_erp).strip() if pd.notna(raw_erp) else ""
                    # ERP must be 11 digits
                    erp_valid = False
                    if erp_str.isdigit() and len(erp_str) == 11:
                        erp_valid = True
                    branch = row[mapping["branch"]]
                    category = row[mapping["category"]]
                    assigned = row[mapping["assigned"]]
                    a_start, a_end = split_assigned_shift(assigned)
                    assigned_working = None
                    if a_start and a_end:
                        dt_start = datetime.combine(date.today(), a_start)
                        dt_end = datetime.combine(date.today(), a_end)
                        if dt_end < dt_start:
                            dt_end += timedelta(days=1)
                        assigned_working = dt_end - dt_start
                    # system-calculated expected start = first_bell - buffer
                    expected_start_time = (datetime.combine(date.today(), first_bell) - timedelta(minutes=buffer_minutes)).time()
                    # determine duration from category master or default
                    cat_hours = st.session_state["shift_categories"].get(str(category).strip(), None)
                    if cat_hours is None:
                        cat_hours = DEFAULT_WORKING_HOURS.get(str(category).strip(), None)
                    duration_td = parse_duration_to_timedelta(cat_hours) if cat_hours else None
                    expected_end_time = None
                    expected_working = None
                    if duration_td is not None:
                        es_dt = datetime.combine(date.today(), expected_start_time)
                        ee_dt = es_dt + duration_td
                        expected_end_time = ee_dt.time()
                        expected_working = duration_td
                    # compare
                    status = "Unknown"
                    notes = []
                    if not erp_valid:
                        status = "Invalid ERP"
                        notes.append("ERP must be 11 digits")
                    elif a_start is None or a_end is None:
                        status = "Bad Assigned Shift Format"
                        notes.append("Assigned shift must be HH:MM-HH:MM")
                    else:
                        # compare assigned start vs expected start within buffer tolerance (allow +/- buffer?)
                        # we'll compute difference in minutes
                        diff_start = (datetime.combine(date.today(), a_start) - datetime.combine(date.today(), expected_start_time)).total_seconds() / 60.0
                        diff_end = None
                        if expected_end_time:
                            diff_end = (datetime.combine(date.today(), a_end) - datetime.combine(date.today(), expected_end_time)).total_seconds() / 60.0
                        # treat a match if both diffs are within 5 minutes
                        tol_minutes = 5
                        if expected_working is None:
                            status = "No Expected (no category duration)"
                            notes.append("No duration configured for category")
                        else:
                            if abs(diff_start) <= tol_minutes and abs(diff_end) <= tol_minutes:
                                status = "Match"
                            else:
                                status = "Mismatch"
                                notes.append(f"Start Δ {diff_start:.0f} min, End Δ {diff_end:.0f} min")
                    out_rows.append({
                        "ERP ID": erp_str,
                        "ERP Valid": erp_valid,
                        "Branch": branch,
                        "Category": category,
                        "Assigned Start": a_start.strftime("%H:%M") if a_start else "",
                        "Assigned End": a_end.strftime("%H:%M") if a_end else "",
                        "Assigned Working Hours": format_timedelta(assigned_working) if assigned_working else "",
                        "Expected Start": expected_start_time.strftime("%H:%M") if expected_start_time else "",
                        "Expected End": expected_end_time.strftime("%H:%M") if expected_end_time else "",
                        "Expected Working Hours": format_timedelta(expected_working) if expected_working else "",
                        "Status": status,
                        "Notes": "; ".join(notes),
                    })
                out_df = pd.DataFrame(out_rows)
                st.markdown("#### Validation Results")
                st.dataframe(out_df, use_container_width=True, height=400)
                # show exceptions
                exceptions = out_df[out_df["Status"] != "Match"]
                st.markdown(f"**Exceptions: {len(exceptions)} row(s)**")
                if not exceptions.empty:
                    st.dataframe(exceptions, use_container_width=True, height=300)
                # allow download
                download_button_for_df(out_df, "⬇️ Download validation results", "shift_bulk_validation_results.xlsx")
        except Exception as e:
            st.error(safe_error_message(e, context="processing bulk shifts"))


