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
safe_error_message = _u.safe_error_message

t = _load_module("tasks_db.py", "hr_tasks")
t.init_db()

st.set_page_config(page_title="Task Manager", page_icon="✅", layout="wide")
render_top_nav("Task Manager")

st.title("✅ Task Manager")
st.caption(
    "Tasks are stored in a local database and stick around until you mark them Completed or delete them — "
    "they don't disappear when you close the tab or reload the page."
)

# ---------------------------------------------------------------------- #
# Create
# ---------------------------------------------------------------------- #
with st.expander("➕ Add a new task", expanded=True):
    with st.form("new_task_form", clear_on_submit=True):
        c1, c2 = st.columns([2, 1])
        with c1:
            new_title = st.text_input("Title")
            new_description = st.text_area("Description (optional)", height=80)
        with c2:
            has_due_date = st.checkbox("Set a due date", value=True)
            new_due_date = st.date_input("Due date", value=date.today(), disabled=not has_due_date)
            new_status = st.selectbox("Status", t.VALID_STATUSES, index=0)

        submitted = st.form_submit_button("Add Task")
        if submitted:
            try:
                t.create_task(
                    new_title, new_description,
                    new_due_date if has_due_date else None, new_status,
                )
                st.success(f"Added: {new_title}")
                st.rerun()
            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(safe_error_message(e, context="adding the task"))

st.divider()

# ---------------------------------------------------------------------- #
# Read / filter
# ---------------------------------------------------------------------- #
filter_col, cleanup_col = st.columns([3, 1])
with filter_col:
    status_filter = st.radio("Show", ["All"] + list(t.VALID_STATUSES), horizontal=True, key="status_filter")
with cleanup_col:
    if st.button("🗑️ Clear all completed"):
        removed = t.delete_all_completed()
        st.success(f"Removed {removed} completed task(s).")
        st.rerun()

tasks = t.list_tasks(status_filter=status_filter)

if not tasks:
    st.info("No tasks here yet.")
else:
    today = date.today()
    for task in tasks:
        due = date.fromisoformat(task["due_date"]) if task["due_date"] else None
        overdue = due is not None and due < today and task["status"] != "Completed"

        with st.container(border=True):
            c1, c2, c3 = st.columns([3, 1, 1])

            with c1:
                title_display = f"~~{task['title']}~~" if task["status"] == "Completed" else f"**{task['title']}**"
                st.markdown(title_display)
                if task["description"]:
                    st.caption(task["description"])
                due_caption = f"Due: {due.isoformat()}" if due else "No due date"
                if overdue:
                    st.markdown(f"🔴 **Overdue** — {due_caption}")
                else:
                    st.caption(due_caption)

            with c2:
                new_status_for_row = st.selectbox(
                    "Status", t.VALID_STATUSES,
                    index=list(t.VALID_STATUSES).index(task["status"]),
                    key=f"status_{task['task_id']}",
                    label_visibility="collapsed",
                )
                if new_status_for_row != task["status"]:
                    t.update_task(task["task_id"], status=new_status_for_row)
                    st.rerun()

            with c3:
                edit_key = f"editing_{task['task_id']}"
                bc1, bc2 = st.columns(2)
                with bc1:
                    if st.button("✏️", key=f"edit_btn_{task['task_id']}", help="Edit"):
                        st.session_state[edit_key] = not st.session_state.get(edit_key, False)
                with bc2:
                    if st.button("🗑️", key=f"delete_btn_{task['task_id']}", help="Delete"):
                        t.delete_task(task["task_id"])
                        st.rerun()

            if st.session_state.get(f"editing_{task['task_id']}"):
                st.markdown("###### Edit task")
                with st.form(f"edit_form_{task['task_id']}"):
                    ec1, ec2 = st.columns([2, 1])
                    with ec1:
                        edit_title = st.text_input("Title", value=task["title"], key=f"et_{task['task_id']}")
                        edit_desc = st.text_area("Description", value=task["description"] or "", key=f"ed_{task['task_id']}")
                    with ec2:
                        edit_has_due = st.checkbox("Has due date", value=due is not None, key=f"ehd_{task['task_id']}")
                        edit_due = st.date_input("Due date", value=due or date.today(), key=f"edt_{task['task_id']}", disabled=not edit_has_due)

                    save_col, cancel_col = st.columns(2)
                    with save_col:
                        if st.form_submit_button("💾 Save"):
                            try:
                                t.update_task(
                                    task["task_id"], title=edit_title, description=edit_desc,
                                    due_date=edit_due if edit_has_due else None,
                                    _due_date_explicitly_cleared=not edit_has_due,
                                )
                                st.session_state[f"editing_{task['task_id']}"] = False
                                st.rerun()
                            except ValueError as e:
                                st.error(str(e))
                    with cancel_col:
                        if st.form_submit_button("Cancel"):
                            st.session_state[f"editing_{task['task_id']}"] = False
                            st.rerun()
