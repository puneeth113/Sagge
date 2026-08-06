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
st.caption("Define shifts based on a First-Bell time, assign category-based working hours (PP/APP/L), and bulk-upload shift assignments.")

# Default working hours per category (HH:MM)
DEFAULT_WORKING_HOURS = {
    "PP": "6:15",  # Preprimary
    "APP": "6:45",  # Above Primary
    "L": "8:15",  # Leader
}


def parse_duration_to_timedelta(s: str) -> timedelta:
    """Parses a string like '6:15' or '06:15' into a timedelta."""
    if pd.isna(s):
        return None
    if isinstance(s, (int, float)):
        # treat as minutes if numeric
        return timedelta(minutes=int(s))
    text = str(s).strip()
    if not text:
        return None
    try:
        parts = text.split(":")
        if len(parts) == 1:
            # only hours provided
            hours = int(parts[0])
            minutes = 0
        else:
            hours = int(parts[0])
            minutes = int(parts[1])
        return timedelta(hours=hours, minutes=minutes)
    except Exception:
        # try parsing as pandas Timedelta
        try:
            td = pd.to_timedelta(text)
            return td
        except Exception:
            return None


def _parse_time_to_time(v) -> time:
    """Robustly parse a variety of time representations and return a datetime.time
    in 24-hour basis. Handles:
      - datetime.time or datetime.datetime inputs
      - strings like '06:00', '6:00 AM', '18:00', '6 PM'
      - Excel fractional day floats (e.g. 0.25) -> treated as fraction of 24h
      - pandas Timestamp or parsed datetimes
    Returns None if parsing fails.
    """
    if pd.isna(v):
        return None
    # direct time
    if isinstance(v, time):
        return v
    # datetime -> extract time
    if isinstance(v, datetime):
        return v.time()
    # pandas Timestamp
    try:
        import pandas as _pd

        if isinstance(v, _pd.Timestamp):
            return v.to_pydatetime().time()
    except Exception:
        pass

    # numeric: handle Excel-like fractional day (0.0 - 1.0) or epoch millis
    if isinstance(v, (int, float)):
        try:
            if 0.0 <= float(v) < 2.0:
                # treat as fraction of day
                seconds = int(float(v) * 24 * 3600)
                h = seconds // 3600
                m = (seconds % 3600) // 60
                s = seconds % 60
                return time(hour=h % 24, minute=m, second=s)
            # otherwise try as unix timestamp (seconds)
            try:
                dt = datetime.fromtimestamp(float(v))
                return dt.time()
            except Exception:
                pass
        except Exception:
            pass

    # string: try pandas to_datetime first (flexible)
    try:
        parsed = pd.to_datetime(v, errors="coerce")
        if not pd.isna(parsed):
            return parsed.time()
    except Exception:
        pass

    # fallback to dateutil (more flexible)
    try:
        from dateutil import parser

        parsed = parser.parse(str(v))
        return parsed.time()
    except Exception:
        return None


def format_timedelta(td: timedelta) -> str:
    if td is None or pd.isna(td):
        return ""
    total_seconds = int(td.total_seconds())
    hours = total_seconds // 3600
    minutes = (total_seconds % 3600) // 60
    return f"{hours:02d}:{minutes:02d}"


def sample_shift_template() -> pd.DataFrame:
    return pd.DataFrame([
        {
            "Employee ID": "E001",
            "Employee Name": "Ravi Kumar",
            "Designation": "Bus Maid",
            "Category": "PP",
            "First Bell Time": "06:00",
            "Category Working Hours (optional)": "",
        },
        {
            "Employee ID": "E002",
            "Employee Name": "Sita Sharma",
            "Designation": "Cleaner",
            "Category": "APP",
            "First Bell Time": "06:30",
            "Category Working Hours (optional)": "6:30",
        },
        {
            "Employee ID": "E003",
            "Employee Name": "Arjun",
            "Designation": "Supervisor",
            "Category": "L",
            "First Bell Time": "08:00",
            "Category Working Hours (optional)": "",
        },
    ])


def format_time_obj(dt) -> str:
    """Format a datetime.datetime or datetime.time into H:MM (no leading zero hour).
    Returns None if input is None/NaN.
    """
    if dt is None or pd.isna(dt):
        return None
    try:
        if isinstance(dt, datetime):
            h = dt.hour
            m = dt.minute
        elif isinstance(dt, time):
            h = dt.hour
            m = dt.minute
        else:
            # try parsing
            t = _parse_time_to_time(dt)
            if t is None:
                return None
            h = t.hour
            m = t.minute
        return f"{h}:{m:02d}"
    except Exception:
        return None


# --- New helpers: auto-detect columns and category management -----------------

def _normalize_col(c: str) -> str:
    return str(c).strip().lower()


def detect_columns(cols: list) -> dict:
    """Try to auto-detect common column names. Returns mapping or None for missing."""
    norm = {c: _normalize_col(c) for c in cols}

    def find(candidates):
        for orig, lower in norm.items():
            for cand in candidates:
                if cand in lower:
                    return orig
        return None

    id_candidates = ["employee id", "emp id", "id", "employeeid", "empid"]
    name_candidates = ["employee name", "name", "emp name", "employee_name"]
    desig_candidates = ["designation", "role", "job title", "designation"]
    first_bell_candidates = ["first bell time", "first bell", "bell time", "first_bell_time", "start time", "time", "start_time"]
    category_candidates = ["category", "cat", "shift category", "category"]
    override_candidates = ["category working hours", "working hours", "per-row working hours", "hours", "category_working_hours"]

    return {
        "id": find(id_candidates),
        "name": find(name_candidates),
        "designation": find(desig_candidates),
        "first_bell": find(first_bell_candidates),
        "category": find(category_candidates),
        "override_hours": find(override_candidates),
    }


# Initialize runtime defaults from session_state or defaults so the category
# subsection can show and update the live list.
if "runtime_defaults" not in st.session_state:
    st.session_state["runtime_defaults"] = DEFAULT_WORKING_HOURS.copy()

runtime_defaults = st.session_state["runtime_defaults"]


# --- UI: Category creation and Shift computation as separate subsections ----

# Category creation subsection
st.markdown("#### Category creation")
with st.expander("View active categories and add/update categories", expanded=True):
    st.caption("Current active categories and their default working hours. Add a new category or edit existing ones below.")
    # show current categories as a table
    try:
        df_cats = pd.DataFrame([{"Category": k, "Default Hours (HH:MM)": v} for k, v in runtime_defaults.items()])
        st.table(df_cats)
    except Exception:
        st.write(runtime_defaults)

    st.markdown("**Edit existing category default hours**")
    cols = st.columns(3)
    edited = {}
    keys = list(runtime_defaults.keys())
    for i, k in enumerate(keys):
        with cols[i % 3]:
            newv = st.text_input(f"{k} default hours (HH:MM)", value=runtime_defaults.get(k, ""), key=f"edit_cat_{k}")
            edited[k] = newv

    st.markdown("**Add a new category**")
    new_cat_col1, new_cat_col2, new_cat_col3 = st.columns([2, 2, 1])
    with new_cat_col1:
        new_cat_key = st.text_input("Category code (e.g. G1)", value="", max_chars=10, key="new_cat_key_main")
    with new_cat_col2:
        new_cat_hours = st.text_input("Default hours (HH:MM)", value="", key="new_cat_hours_main")
    with new_cat_col3:
        if st.button("Add / Update category", key="add_update_cat_main"):
            if not new_cat_key.strip():
                st.error("Provide a category code (non-empty).")
            else:
                code = new_cat_key.strip().upper()
                st.session_state["runtime_defaults"][code] = new_cat_hours.strip()
                st.success(f"Added/Updated category {code}")

    if st.button("Save category edits", key="save_cat_edits"):
        for k, v in edited.items():
            if v is not None:
                st.session_state["runtime_defaults"][k] = v.strip()
        st.success("Saved category timings.")


# Shift computation subsection
st.markdown("#### Shift computation")
st.caption("Compute shifts either via Bulk Upload or Single/manual entry.")

# --- Sub-navigation: two tabs to match requested layout ---------------------
sub_tabs = st.tabs(["📥 Bulk Upload", "✍️ Single Branch (manual)"])

# Bulk tab (index 0)
with sub_tabs[0]:
    st.markdown("##### Bulk upload (assign shifts & auto-compute start/end times)")
    st.caption("Upload a sheet with at least: First Bell Time, Category (PP/APP/L) and optional per-row working hours. The sheet can also include Employee ID/Name/Designation.")

    with st.container():
        with st.expander("Defaults (category working hours)", expanded=False):
            c1, c2, c3 = st.columns(3)
            def val_for(k):
                return runtime_defaults.get(k, "")
            with c1:
                pp_default = st.text_input("Preprimary (PP) default working hours (HH:MM)", value=val_for("PP"), key="pp_default_bulk")
            with c2:
                app_default = st.text_input("Above Primary (APP) default working hours (HH:MM)", value=val_for("APP"), key="app_default_bulk")
            with c3:
                l_default = st.text_input("Leader (L) default working hours (HH:MM)", value=val_for("L"), key="l_default_bulk")
            # update runtime defaults with any edits in this expander
            runtime_defaults.update({"PP": pp_default or runtime_defaults.get("PP"), "APP": app_default or runtime_defaults.get("APP"), "L": l_default or runtime_defaults.get("L")})

        uploaded = st.file_uploader("Upload shift assignments (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="shifts_bulk_tab")

        if st.button("Download sample template", key="download_sample_shift_bulk"):
            tmp = sample_shift_template()
            download_button_for_df(tmp, "⬇️ Download sample shift template", "sample_shift_template.xlsx")

        if uploaded:
            try:
                df = read_any_table(uploaded)
                st.dataframe(df.head(5), use_container_width=True)
                cols = list(df.columns)

                # attempt auto-detection
                detected = detect_columns(cols)
                required_found = detected.get("first_bell") and detected.get("category")

                st.markdown("###### Column mapping")
                if required_found:
                    st.success("Required columns auto-detected: First Bell and Category. Mapping will be applied automatically.")
                    # show detected mapping in compact way and allow override
                    with st.expander("Detected mapping (click to override)"):
                        id_col = st.selectbox("Employee ID column (optional)", ["(none)"] + cols, index=(1 + cols.index(detected["id"])) if detected.get("id") in cols else 0)
                        name_col = st.selectbox("Employee Name column (optional)", ["(none)"] + cols, index=(1 + cols.index(detected["name"])) if detected.get("name") in cols else 0)
                        designation_col = st.selectbox("Designation column (optional)", ["(none)"] + cols, index=(1 + cols.index(detected["designation"])) if detected.get("designation") in cols else 0)
                        first_bell_col = st.selectbox("First Bell Time column (required)", cols, index=cols.index(detected["first_bell"]))
                        category_col = st.selectbox("Category column (required: PP/APP/L)", cols, index=cols.index(detected["category"]))
                        override_hours_col = st.selectbox("Per-row Working Hours column (optional, HH:MM)", ["(none)"] + cols, index=(1 + cols.index(detected["override_hours"])) if detected.get("override_hours") in cols else 0)
                else:
                    st.info("Could not auto-detect required columns — please map them below.")
                    id_col = st.selectbox("Employee ID column (optional)", ["(none)"] + cols, index=0)
                    name_col = st.selectbox("Employee Name column (optional)", ["(none)"] + cols, index=0)
                    designation_col = st.selectbox("Designation column (optional)", ["(none)"] + cols, index=0)
                    first_bell_col = st.selectbox("First Bell Time column (required)", cols, key="first_bell_col_bulk")
                    category_col = st.selectbox("Category column (required: PP/APP/L)", cols, key="category_col_bulk")
                    override_hours_col = st.selectbox("Per-row Working Hours column (optional, HH:MM)", ["(none)"] + cols, index=0)

                if st.button("Compute shifts for uploaded sheet", key="compute_bulk"):
                    work = df.copy()

                    # parse first bell times to datetime.time (24-hour aware)
                    work["_first_bell_parsed"] = work[first_bell_col].apply(_parse_time_to_time)

                    # normalize category values
                    work["_category_norm"] = work[category_col].astype(str).str.strip().str.upper()

                    # compute durations
                    def choose_duration(row):
                        # per-row override column if provided
                        if override_hours_col != "(none)" and override_hours_col in work.columns:
                            val = row.get(override_hours_col, None)
                            if pd.notna(val) and str(val).strip() != "":
                                td = parse_duration_to_timedelta(val)
                                if td is not None:
                                    return td
                        # category-based default (runtime_defaults)
                        cat = row["_category_norm"] if pd.notna(row["_category_norm"]) else ""
                        # map common full words to keys
                        if cat.startswith("PRE"):
                            key = "PP"
                        elif cat.startswith("APP") or cat.startswith("ABOVE"):
                            key = "APP"
                        elif cat.startswith("L") or cat.startswith("LEAD") or cat.startswith("SUP"):
                            key = "L"
                        else:
                            # fallback - try exact match
                            key = cat if cat in runtime_defaults else None
                        if key and key in runtime_defaults and runtime_defaults[key]:
                            return parse_duration_to_timedelta(runtime_defaults[key])
                        # final fallback - None
                        return None

                    work["_working_duration_td"] = work.apply(choose_duration, axis=1)

                    # compute start & end datetimes
                    starts = []
                    ends = []
                    durations = []
                    for _, r in work.iterrows():
                        tb = r["_first_bell_parsed"]
                        td = r["_working_duration_td"]
                        if pd.isna(tb) or tb is None:
                            starts.append(None)
                            ends.append(None)
                            durations.append(None)
                            continue
                        today = date.today()
                        start_dt = datetime.combine(today, tb)
                        if td is None:
                            ends.append(None)
                        else:
                            end_dt = start_dt + td
                            ends.append(end_dt)
                        starts.append(start_dt)
                        durations.append(td)

                    # output in compact H:MM format and a combined Start-End column like "7:45-14:00"
                    start_strs = [format_time_obj(s) for s in starts]
                    end_strs = [format_time_obj(e) for e in ends]
                    work["Shift Start"] = start_strs
                    work["Shift End"] = end_strs
                    work["Shift (Start-End)"] = [f"{a}-{b}" if a and b else (a or b) for a, b in zip(start_strs, end_strs)]
                    work["Working Hours (HH:MM)"] = [format_timedelta(d) if d is not None else None for d in durations]

                    st.session_state["computed_shifts"] = work
                    st.success(f"Computed shifts for {len(work)} rows.")

            except Exception as e:
                st.error(safe_error_message(e, context="processing uploaded shift file"))

# Single/manual tab (index 1)
with sub_tabs[1]:
    st.markdown("##### Single / Manual shift compute")
    with st.expander("Compute a single shift from First Bell", expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            # time_input returns a datetime.time object; keep as-is but display in 24-hour on output
            first_bell_time = st.time_input("First Bell time (24-hour HH:MM)", value=datetime.now().time().replace(second=0, microsecond=0), key="first_bell_single")
        with col2:
            category = st.selectbox("Category", list(runtime_defaults.keys()), index=0, key="category_single")
        with col3:
            override_hours = st.text_input("Override working hours (optional, HH:MM)", value="", key="override_single")

        if st.button("Compute Shift (single)", key="compute_single"):
            try:
                dur = parse_duration_to_timedelta(override_hours) if override_hours else parse_duration_to_timedelta(runtime_defaults.get(category))
                if dur is None:
                    st.error("Could not parse working hours. Use HH:MM format, e.g. 6:15")
                else:
                    today = date.today()
                    start_dt = datetime.combine(today, first_bell_time)
                    end_dt = start_dt + dur
                    # display in compact H:MM-H:MM format
                    start_str = format_time_obj(start_dt)
                    end_str = format_time_obj(end_dt)
                    st.metric("Shift Start", start_str)
                    st.metric("Shift End", end_str)
                    st.markdown(f"**Shift:** {start_str}-{end_str}")
                    st.write(f"Working hours: {format_timedelta(dur)} (HH:MM)")
            except Exception as e:
                st.error(safe_error_message(e, context="computing single shift"))

# Final computed results area (shared)
if "computed_shifts" in st.session_state:
    out = st.session_state["computed_shifts"]
    st.divider()
    st.markdown("#### Computed shift results")
    st.dataframe(out, use_container_width=True, height=420)
    download_button_for_df(out, "⬇️ Download computed shifts", "computed_shifts.xlsx")
