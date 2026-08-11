import os
import sqlite3
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
st.caption("Category master and bulk shift operations. Use the masters to configure categories, branch start times and named shifts, then run bulk processing to validate and update assignments.")


# Default working hours per category (HH:MM). These figures already include
# whatever fixed daily allowance (e.g. break time) applies to that category -
# there is no separate buffer subtracted anywhere else in this file.
DEFAULT_WORKING_HOURS = {
    "PP": "6:15",  # Preprimary
    "APP": "6:45",  # Above Primary
    "L": "8:15",  # Leader
}

# Tolerance (in minutes) used only when comparing an employee's currently
# assigned shift against the system-computed shift, to decide whether they
# should be flagged as "matches" or "does not match". This has nothing to do
# with how the working hours / expected end time are computed.
MATCH_TOLERANCE_MINUTES = 5


def _rerun():
    """Compat wrapper: st.rerun() on newer Streamlit, falls back on older ones."""
    if hasattr(st, "rerun"):
        st.rerun()
    else:
        st.experimental_rerun()


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


# --- Persistent storage (SQLite) for branch start times ---------------------------
#
# Branch start times are the one master that needs to survive app restarts so
# users don't have to re-upload the sheet every session. Everything lives in a
# small local SQLite database file next to the app - no other files are
# touched or required.


def _get_db_path() -> str:
    this_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.dirname(this_dir)
    data_dir = os.path.join(root_dir, "data")
    os.makedirs(data_dir, exist_ok=True)
    return os.path.join(data_dir, "shift_management.db")


DB_PATH = _get_db_path()


def _get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def _migrate_old_single_key_schema(conn):
    """One-time upgrade for databases created by an earlier version of this
    page, where (branch_name) alone was the primary key. That bug meant
    saving a second category for the same branch silently overwrote the
    first one. This copies whatever survived into the new, correct table
    keyed on (branch_name, category).
    """
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(branch_start_times)")
    cols = cur.fetchall()
    if not cols:
        return  # table doesn't exist yet, nothing to migrate
    pk_cols = [c[1] for c in cols if c[5] > 0]  # column name where pk index > 0
    if pk_cols == ["branch_name", "category"]:
        return  # already on the new schema
    # Old schema detected (pk is just branch_name) - rename, recreate, copy over.
    cur.execute("ALTER TABLE branch_start_times RENAME TO branch_start_times_old")
    cur.execute(
        """
        CREATE TABLE branch_start_times (
            branch_name TEXT NOT NULL,
            category TEXT NOT NULL,
            start_time TEXT NOT NULL,
            PRIMARY KEY (branch_name, category)
        )
        """
    )
    cur.execute(
        """
        INSERT OR IGNORE INTO branch_start_times (branch_name, category, start_time)
        SELECT branch_name, category, start_time FROM branch_start_times_old
        """
    )
    cur.execute("DROP TABLE branch_start_times_old")
    conn.commit()


def init_db():
    conn = _get_conn()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS branch_start_times (
                branch_name TEXT NOT NULL,
                category TEXT NOT NULL,
                start_time TEXT NOT NULL,
                PRIMARY KEY (branch_name, category)
            )
            """
        )
        conn.commit()
        _migrate_old_single_key_schema(conn)
    finally:
        conn.close()


def load_branch_start_times() -> dict:
    """Returns a nested dict: {branch_name: {category: start_time_str}}.

    A branch can have a different start time per category (e.g. PP starts
    at 08:00 but L starts at 08:30 at the same branch), so both the branch
    name AND the category together identify a row.
    """
    conn = _get_conn()
    try:
        cur = conn.cursor()
        cur.execute("SELECT branch_name, category, start_time FROM branch_start_times ORDER BY branch_name, category")
        rows = cur.fetchall()
    finally:
        conn.close()
    result: dict = {}
    for branch_name, category, start_time in rows:
        result.setdefault(branch_name, {})[category] = start_time
    return result


def upsert_branch_start_time(branch_name: str, category: str, start_time: str):
    """Insert or update the start time for one (branch, category) pair.
    Saving a new category for a branch that already exists adds a new row
    rather than overwriting the branch's other categories.
    """
    conn = _get_conn()
    try:
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO branch_start_times (branch_name, category, start_time)
            VALUES (?, ?, ?)
            ON CONFLICT(branch_name, category) DO UPDATE SET
                start_time = excluded.start_time
            """,
            (branch_name, category, start_time),
        )
        conn.commit()
    finally:
        conn.close()


def delete_branch_start_time(branch_name: str, category: str = None):
    """Deletes one (branch, category) row, or every row for a branch if
    category is omitted."""
    conn = _get_conn()
    try:
        cur = conn.cursor()
        if category is None:
            cur.execute("DELETE FROM branch_start_times WHERE branch_name = ?", (branch_name,))
        else:
            cur.execute(
                "DELETE FROM branch_start_times WHERE branch_name = ? AND category = ?",
                (branch_name, category),
            )
        conn.commit()
    finally:
        conn.close()


init_db()


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
    # Branch start times always come from the database, so it survives
    # app restarts and doesn't need to be re-uploaded each session.
    st.session_state["branch_start_times"] = load_branch_start_times()


_ensure_masters()


# Top-level sub-navigation for the two required pages
page = st.radio("Shift Management — Section", options=["Category & Shift Master", "Bulk Shift Operations"], horizontal=True)

if page == "Category & Shift Master":
    st.markdown("### Category & Shift Master")
    st.caption("Create / edit categories, named shifts, and manage branch start times used by bulk operations. Branch start times are saved permanently to the database, so you only need to upload them once.")

    col1, col2 = st.columns([2, 1])
    with col1:
        st.subheader("Categories (code -> default working hours HH:MM)")
        cat_df = pd.DataFrame([{"Category": k, "Default Hours": v} for k, v in st.session_state["shift_categories"].items()])
        st.dataframe(cat_df, use_container_width=True)

        st.markdown("**Add new category**")
        new_code = st.text_input("Category code (e.g. G1)", value="", key="new_cat_code")
        new_hours = st.text_input("Default hours (HH:MM)", value="6:00", key="new_cat_hours")
        if st.button("Add category"):
            if not new_code.strip():
                st.error("Provide a category code.")
            else:
                st.session_state["shift_categories"][new_code.strip()] = new_hours.strip()
                st.success(f"Added category {new_code.strip()} -> {new_hours.strip()}")
                _rerun()

        st.markdown("**Edit existing category**")
        sel_cat = st.selectbox("Select category to edit", options=list(st.session_state["shift_categories"].keys()))
        if sel_cat:
            cur = st.session_state["shift_categories"][sel_cat]
            new_val = st.text_input("Default hours (HH:MM)", value=cur, key=f"edit_cat_{sel_cat}")
            if st.button("Save category", key=f"save_cat_{sel_cat}"):
                st.session_state["shift_categories"][sel_cat] = new_val.strip()
                st.success(f"Saved {sel_cat} -> {new_val.strip()}")
                _rerun()

        if st.button("Delete selected category"):
            if sel_cat in st.session_state["shift_categories"]:
                del st.session_state["shift_categories"][sel_cat]
                st.success(f"Deleted category {sel_cat}")
                _rerun()

        st.markdown("---")
        st.subheader("Branch start times (branch -> category, start time)")
        st.markdown(
            "Upload a sheet with: Branch Name, Category, Start Time (HH:MM), or add branches one at a time below. "
            "Entries are saved permanently to the database - you won't need to re-upload them next time you open the app. "
            "This master is used to compute the expected shift for employees in that branch."
        )

        if st.button("Download sample branch start template"):
            try:
                sample = pd.DataFrame([
                    {"Branch Name": "OIS Sample Branch", "Category": "PP", "Start Time": "08:00"},
                    {"Branch Name": "Another Branch", "Category": "APP", "Start Time": "08:30"},
                ])
                b = to_excel_bytes({"BranchStartSample": sample})
                st.download_button("⬇️ Download Branch Start sample", data=b, file_name="branch_start_sample.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(safe_error_message(e, context="creating branch sample"))

        upload_branch = st.file_uploader("Upload branch start times (.xlsx or .csv)", type=["xlsx", "xls", "csv"], key="upload_branch_start")
        if upload_branch:
            try:
                bdf = read_any_table(upload_branch)
                st.markdown("Preview of uploaded branch starts")
                st.dataframe(bdf.head(20), use_container_width=True)

                # auto-detect columns
                cols_lower = {c.lower(): c for c in bdf.columns}
                mapping = {}
                for candidate in ["branch name", "branch", "branch_name"]:
                    if candidate in cols_lower:
                        mapping["branch"] = cols_lower[candidate]
                        break
                for candidate in ["category", "cat"]:
                    if candidate in cols_lower:
                        mapping["category"] = cols_lower[candidate]
                        break
                for candidate in ["start time", "start_time", "starting time", "start"]:
                    if candidate in cols_lower:
                        mapping["start_time"] = cols_lower[candidate]
                        break

                missing = [k for k in ("branch", "category", "start_time") if k not in mapping]
                if missing:
                    st.error(f"Missing required columns in branch start upload: {missing}")
                else:
                    if st.button("Save branch starts to database"):
                        saved = 0
                        for i, r in bdf.iterrows():
                            br = str(r[mapping["branch"]]).strip()
                            cat = str(r[mapping["category"]]).strip()
                            stime = r[mapping["start_time"]]
                            t = _parse_time_to_time(stime)
                            start_time_str = t.strftime("%H:%M") if t else str(stime).strip()
                            if br:
                                upsert_branch_start_time(br, cat, start_time_str)
                                saved += 1
                        st.success(f"Saved {saved} branch start time(s) to the database")
                        _rerun()
            except Exception as e:
                st.error(safe_error_message(e, context="reading branch start upload"))

        st.markdown("**Add / update a single branch**")
        bc1, bc2, bc3 = st.columns(3)
        with bc1:
            single_branch = st.text_input("Branch name", value="", key="single_branch_name")
        with bc2:
            cat_options = list(st.session_state["shift_categories"].keys())
            single_category = st.selectbox("Category", options=cat_options, key="single_branch_category") if cat_options else st.text_input("Category", key="single_branch_category_txt")
        with bc3:
            single_start = st.text_input("Start time (HH:MM)", value="08:00", key="single_branch_start")
        if st.button("Save branch to database"):
            if not single_branch.strip():
                st.error("Provide a branch name.")
            else:
                upsert_branch_start_time(single_branch.strip(), str(single_category).strip(), single_start.strip())
                st.success(f"Saved {single_branch.strip()} to the database")
                _rerun()

        st.markdown("Current branch start master (stored in database)")
        bst = st.session_state["branch_start_times"]
        # bst is nested: {branch_name: {category: start_time}} - a branch can
        # have several rows, one per category, each with its own start time.
        bst_rows = [
            {"Branch": branch, "Category": category, "Start Time": start_time}
            for branch, cat_map in bst.items()
            for category, start_time in cat_map.items()
        ]
        bst_df = pd.DataFrame(bst_rows)
        st.dataframe(bst_df, use_container_width=True)

        if bst_rows:
            row_labels = [f"{r['Branch']} — {r['Category']} ({r['Start Time']})" for r in bst_rows]
            del_choice = st.selectbox("Select a branch + category entry to delete", options=row_labels, key="del_branch_select")
            if st.button("Delete selected entry"):
                chosen = bst_rows[row_labels.index(del_choice)]
                delete_branch_start_time(chosen["Branch"], chosen["Category"])
                st.success(f"Deleted {chosen['Branch']} ({chosen['Category']}) from the database")
                _rerun()

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
                _rerun()

        st.markdown("**Edit existing shift**")
        sel_shift = st.selectbox("Select shift to edit", options=list(nsh.keys()))
        if sel_shift:
            cur = st.session_state["named_shifts"][sel_shift]
            new_span = st.text_input("Shift span (HH:MM-HH:MM)", value=cur, key=f"edit_shift_{sel_shift}")
            if st.button("Save shift", key=f"save_shift_{sel_shift}"):
                st.session_state["named_shifts"][sel_shift] = new_span.strip()
                st.success(f"Saved {sel_shift} -> {new_span.strip()}")
                _rerun()

        if st.button("Delete selected shift"):
            if sel_shift in st.session_state["named_shifts"]:
                del st.session_state["named_shifts"][sel_shift]
                st.success(f"Deleted shift {sel_shift}")
                _rerun()


else:
    st.markdown("### Bulk Shift Operations")
    st.caption("Download the sample template, fill it, upload the file, then click Process to validate and generate the final result. This page uses the saved branch start master to compute expected shifts.")

    # Sample template generator (no First Bell column anymore)
    def sample_bulk_template():
        df = pd.DataFrame([
            {"ERP ID": "12345678901", "Branch Name": "OIS Sample Branch", "Category": "PP", "Current Assigned Shift": "08:00-14:15"},
            {"ERP ID": "12345678902", "Branch Name": "OIS Sample Branch", "Category": "PP", "Current Assigned Shift": "08:00-14:30"},
            {"ERP ID": "12345678903", "Branch Name": "Another Branch", "Category": "APP", "Current Assigned Shift": "08:30-15:15"},
        ])
        instr = pd.DataFrame({"Instructions": [
            "Fill the 'ERP ID', 'Branch Name', 'Category', and 'Current Assigned Shift' (HH:MM-HH:MM).",
            "Do not change the column headers. The system auto-detects these columns case-insensitively.",
            "Branch start times come from the 'Category & Shift Master' -> Branch start times master, saved permanently in the database (upload/add branches there once).",
            "Expected end time = branch start time + the category's default working hours, taken directly with no separate buffer subtracted.",
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
                # Assigned shift
                for candidate in ["current assigned shift", "assigned shift", "assigned_shift", "shift", "current shift", "current_assigned_shift"]:
                    if candidate in cols_lower:
                        mapping["assigned"] = cols_lower[candidate]
                        break
                return mapping

            mapping = auto_detect_columns(df)
            missing = [k for k in ("erp_id", "branch", "category", "assigned") if k not in mapping]
            if missing:
                st.error(f"Missing required columns (auto-detect failed): {missing}. Please ensure the uploaded sheet has these columns with recognizable names.")
            else:
                if st.button("Process / Validate"):
                    out_rows = []
                    for i, row in df.iterrows():
                        erp_val = row[mapping["erp_id"]]
                        erp_str = str(erp_val).strip() if pd.notna(erp_val) else ""
                        branch = str(row[mapping["branch"]]).strip()
                        category = str(row[mapping["category"]]).strip()
                        assigned = row[mapping["assigned"]]

                        a_start, a_end = split_assigned_shift(assigned)
                        current_working = None
                        if a_start and a_end:
                            dt_start = datetime.combine(date.today(), a_start)
                            dt_end = datetime.combine(date.today(), a_end)
                            if dt_end < dt_start:
                                dt_end += timedelta(days=1)
                            current_working = dt_end - dt_start

                        # System expected start = branch start (from the saved database master).
                        # Expected end = expected start + category's full default working hours.
                        # The category hours already include any built-in allowance (e.g. the
                        # ":15" in "6:15"), so no extra buffer is added or subtracted here.
                        expected_start_time = None
                        expected_end_time = None
                        expected_working = None
                        bst = st.session_state.get("branch_start_times", {})
                        # bst is nested: {branch_name: {category: start_time}} - look up
                        # the start time for THIS branch's THIS category specifically,
                        # since the same branch can have a different start time per category.
                        branch_categories = bst.get(branch, {})
                        raw_start_time = branch_categories.get(category)
                        if raw_start_time:
                            bst_time = _parse_time_to_time(raw_start_time)
                            if bst_time:
                                expected_start_time = bst_time
                                # duration from category master
                                cat_hours = st.session_state["shift_categories"].get(category, None)
                                if cat_hours is None:
                                    cat_hours = DEFAULT_WORKING_HOURS.get(category, None)
                                duration_td = parse_duration_to_timedelta(cat_hours) if cat_hours else None
                                if duration_td is not None:
                                    ee_dt = datetime.combine(date.today(), expected_start_time) + duration_td
                                    expected_end_time = ee_dt.time()
                                    expected_working = duration_td

                        # compare and build clear remarks
                        remarks = "To Be Checked"
                        if a_start is None or a_end is None:
                            remarks = "Assigned shift format invalid"
                        elif raw_start_time is None:
                            remarks = "Missing branch start master entry for this branch + category"
                        elif expected_working is None:
                            remarks = "No duration configured for category"
                        else:
                            # compute diffs
                            diff_start = (datetime.combine(date.today(), a_start) - datetime.combine(date.today(), expected_start_time)).total_seconds() / 60.0
                            diff_end = (datetime.combine(date.today(), a_end) - datetime.combine(date.today(), expected_end_time)).total_seconds() / 60.0
                            if abs(diff_start) <= MATCH_TOLERANCE_MINUTES and abs(diff_end) <= MATCH_TOLERANCE_MINUTES:
                                remarks = "Assigned shift matches computed shift"
                            else:
                                remarks = f"Assigned shift does not match computed shift (Start Δ {diff_start:.0f} min, End Δ {diff_end:.0f} min)"

                        out_rows.append({
                            "ERP ID": erp_str,
                            "Branch Name": branch,
                            "Category": category,
                            "Current Assigned Shift": str(assigned),
                            "Current Start Time": a_start.strftime("%H:%M") if a_start else "",
                            "Current End Time": a_end.strftime("%H:%M") if a_end else "",
                            "Current Working Hours": format_timedelta(current_working) if current_working else "",
                            "System Calculated Shift": (f"{expected_start_time.strftime('%H:%M')}-{expected_end_time.strftime('%H:%M')}" if expected_start_time and expected_end_time else ""),
                            "System Working Hours": format_timedelta(expected_working) if expected_working else "",
                            "Remarks": remarks,
                        })

                    out_df = pd.DataFrame(out_rows)
                    st.markdown("#### Final Result")
                    st.dataframe(out_df, use_container_width=True, height=600)
                    # show counts
                    st.markdown(f"**Total rows:** {len(out_df)}; **To Be Checked:** {len(out_df[out_df['Remarks'] != 'Assigned shift matches computed shift'])}")

                    # final download
                    try:
                        b = to_excel_bytes({"Processed_Result": out_df})
                        st.download_button("⬇️ Download final processed file", data=b, file_name="shift_bulk_processed.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except Exception as e:
                        st.error(safe_error_message(e, context="creating final download"))

        except Exception as e:
            st.error(safe_error_message(e, context="processing bulk upload"))
