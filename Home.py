import os
import importlib.util

import streamlit as st


def _load_utils():
    """Loads utils.py by its exact file path (no sys.path / package
    resolution involved), so it works no matter how Streamlit was launched
    or what the current working directory is."""
    this_dir = os.path.dirname(os.path.abspath(__file__))
    for candidate in (
        os.path.join(this_dir, "utils.py"),
        os.path.join(this_dir, "Utils.py"),
    ):
        if os.path.exists(candidate):
            spec = importlib.util.spec_from_file_location("hr_utils", candidate)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise FileNotFoundError(
        "Could not find utils.py. Make sure it sits directly inside the "
        "same folder as Home.py."
    )


_u = _load_utils()
render_top_nav = _u.render_top_nav

# reuse simple json helpers (kept local to avoid circular imports)
import json

def _load_json(path: str, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path: str, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


st.set_page_config(
    page_title="HR Assistant",
    page_icon="🗂️",
    layout="wide",
)

render_top_nav("Home")

# --- require login before showing Home content ---
if not st.session_state.get("logged_in"):
    st.warning("You must log in to access the HR Assistant.")
    st.page_link("pages/0_Login.py", label="Go to Login →", icon="🔐")
    st.stop()

# At this point the user is logged in
st.title("🗂️ HR Assistant")
st.caption(
    "Branch HR operations toolkit — employee database, long-absence tracking, "
    "shift management, and payroll (PF/ESIC/TDS) calculations."
)

st.markdown("#### Navigate")
nav1, nav2, nav3, nav4 = st.columns(4)

with nav1:
    with st.container(border=True):
        st.markdown("##### 📅 Long Absence Tracker")
        st.caption("Upload attendance muster, auto-count `A` days, bucket employees into risk categories.")
        st.page_link("pages/1_Long_Absence_Tracker.py", label="Open →", icon="📅")

with nav2:
    with st.container(border=True):
        st.markdown("##### 🧾 Payroll Calculator")
        st.caption("Compute monthly gross, PF/ESIC deductions and net pay for full-time staff, plus TDS billing for gig workers.")
        st.page_link("pages/3_Payroll_Calculator.py", label="Open →", icon="🧾")

with nav3:
    with st.container(border=True):
        st.markdown("##### 👥 Employee Database")
        st.caption("Upload your employee master sheet, filter and search records, and download filtered views.")
        st.page_link("pages/4_Employee_Database.py", label="Open →", icon="👥")

with nav4:
    with st.container(border=True):
        st.markdown("##### 🕒 Shift Management")
        st.caption("Configure category & branch start time masters, then bulk-validate employee shift assignments.")
        st.page_link("pages/7_Shift_management.py", label="Open →", icon="🕒")

# ------------------------------------------------------------------
# Super-admin: request approval / revoke management
# This section is only visible to admin users and acts as the central
# approval/revoke UI (moved here per your request).
# ------------------------------------------------------------------

if st.session_state.get("role") == "admin":
    st.divider()
    st.header("Admin: Manage access requests and users")

    repo_root = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(repo_root, "..", "data")
    os.makedirs(data_dir, exist_ok=True)
    users_file = os.path.join(data_dir, "users.json")
    pending_file = os.path.join(data_dir, "pending_requests.json")

    users = _load_json(users_file, [])
    pending = _load_json(pending_file, [])

    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("Pending requests")
        if not pending:
            st.info("No pending requests.")
        else:
            for i, r in enumerate(list(pending)):
                st.markdown(f"**{r.get('username')}** — {r.get('reason')}")
                c_approve, c_reject = st.columns([1, 1])
                if c_approve.button("Approve", key=f"home_approve_{i}"):
                    temp_pw = r.get('username')
                    new_user = {
                        "username": r.get("username"),
                        "password_hash": None,
                        "password_hash_algo": "bcrypt",
                        "role": "user",
                        "active": True,
                        "must_change_password": True,
                    }
                    # set a bcrypt hash only if bcrypt available; otherwise keep None
                    try:
                        import bcrypt

                        new_user["password_hash"] = bcrypt.hashpw(temp_pw.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
                    except Exception:
                        # bcrypt not installed; keep password_hash None and require admin to set
                        new_user["password_hash"] = None

                    if not any(u.get("username") == new_user["username"] for u in users):
                        users.append(new_user)
                        _save_json(users_file, users)
                    pending = [p for p in pending if p.get("username") != r.get("username")]
                    _save_json(pending_file, pending)
                    st.success(f"Approved access for {new_user['username']}. User created and will be required to change password.")
                    st.experimental_rerun()
                if c_reject.button("Reject", key=f"home_reject_{i}"):
                    pending = [p for p in pending if p.get("username") != r.get("username")]
                    _save_json(pending_file, pending)
                    st.info(f"Rejected request from {r.get('username')}")
                    st.experimental_rerun()

    with col2:
        st.subheader("Users (revoke / activate)")
        if not users:
            st.info("No users configured.")
        else:
            for j, u in enumerate(list(users)):
                status = "active" if u.get("active", True) else "inactive"
                st.markdown(f"**{u.get('username')}** — role: {u.get('role','user')} — {status}")
                c_revoke, c_activate = st.columns([1, 1])
                if c_revoke.button("Revoke", key=f"revoke_{j}"):
                    for uu in users:
                        if uu.get("username") == u.get("username"):
                            uu["active"] = False
                    _save_json(users_file, users)
                    st.info(f"User {u.get('username')} revoked (inactive).")
                    st.experimental_rerun()
                if c_activate.button("Activate", key=f"activate_{j}"):
                    for uu in users:
                        if uu.get("username") == u.get("username"):
                            uu["active"] = True
                    _save_json(users_file, users)
                    st.success(f"User {u.get('username')} activated.")
                    st.experimental_rerun()

