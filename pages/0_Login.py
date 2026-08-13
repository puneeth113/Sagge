import os
import json
import hashlib
import streamlit as st

# Simple file-backed user store located under data/users.json
DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data")
os.makedirs(DATA_DIR, exist_ok=True)
USERS_FILE = os.path.join(DATA_DIR, "users.json")
PENDING_FILE = os.path.join(DATA_DIR, "pending_requests.json")


def _hash_password(pw: str) -> str:
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()


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


# Small helper to render a divider safely across Streamlit versions
def safe_divider():
    if hasattr(st, "divider"):
        st.divider()
    else:
        st.markdown("---")


# Hide the Streamlit sidebar (used on the login view to avoid the side page menu)
def hide_sidebar():
    css = """
    <style>
    section[data-testid="stSidebar"] {display: none;}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


# Ensure users file exists with the requested admin account
def ensure_initial_admin():
    users = _load_json(USERS_FILE, [])
    admin_username = "20250002367_OIS"
    admin_pw = "superadmin1132002"
    exists = any(u.get("username") == admin_username for u in users)
    if not exists:
        users.append({
            "username": admin_username,
            "password_hash": _hash_password(admin_pw),
            "role": "admin",
            "active": True,
        })
        _save_json(USERS_FILE, users)
    return users


ensure_initial_admin()


st.set_page_config(page_title="Login", page_icon="🔐")
st.title("🔐 Login")

# --- load users & pending requests ---
users = _load_json(USERS_FILE, [])
pending = _load_json(PENDING_FILE, [])

# If the user is not logged in, hide the sidebar menu so the login page is the main focus.
# This removes the side-bar page view while on the login screen.
if not st.session_state.get("logged_in"):
    hide_sidebar()

# Single form that can act as Login or Request Access to avoid ambiguity
with st.form(key="main_form"):
    action = st.radio("Action", ["Login", "Request access"], index=0, horizontal=True)

    # Login fields
    login_username = st.text_input("User name", key="login_user")
    login_password = st.text_input("Password", type="password", key="login_pw")

    # Request access fields
    req_username = st.text_input("Desired user name", key="req_user")
    req_reason = st.text_area("Reason for access (brief)", key="req_reason")

    submitted = st.form_submit_button("Submit")

if submitted:
    if action == "Login":
        if not login_username or not login_password:
            st.error("Please enter both username and password to log in.")
        else:
            found = None
            for u in users:
                if u.get("username") == login_username and u.get("active", True):
                    found = u
                    break
            if found and found.get("password_hash") == _hash_password(login_password):
                st.success(f"Welcome back, {login_username}!")
                st.session_state["logged_in"] = True
                st.session_state["user"] = login_username
                st.session_state["role"] = found.get("role", "user")
                st.experimental_rerun()
            else:
                st.error("Invalid username or password, or account inactive.")
    else:  # Request access
        if not req_username:
            st.error("Please provide a desired user name to request access.")
        else:
            # Record a simple pending request
            pending.append({"username": req_username, "reason": req_reason})
            _save_json(PENDING_FILE, pending)
            st.success("Your access request has been recorded and will be reviewed.")


safe_divider()

# --- Admin approval UI ---
if st.session_state.get("logged_in") and st.session_state.get("role") == "admin":
    safe_divider()
    st.header("Pending access requests")
    if not pending:
        st.info("No pending requests.")
    else:
        # We iterate over a copy of pending with index so we can mutate the real list safely
        for i, r in list(enumerate(pending)):
            st.markdown(f"**{r.get('username')}** — {r.get('reason')}")
            cols = st.columns([1, 1, 6])
            if cols[0].button("Approve", key=f"approve_{i}"):
                # Add user with a temporary password (same as username) — admin should inform them
                new_user = {
                    "username": r.get("username"),
                    "password_hash": _hash_password(r.get("username", "changeme")),
                    "role": "user",
                    "active": True,
                }
                # avoid duplicates
                if not any(u.get("username") == new_user["username"] for u in users):
                    users.append(new_user)
                    _save_json(USERS_FILE, users)
                # remove request
                pending.pop(i)
                _save_json(PENDING_FILE, pending)
                st.success(f"Approved access for {new_user['username']}. Temporary password set to the username — ask them to change it.")
                st.experimental_rerun()
            if cols[1].button("Reject", key=f"reject_{i}"):
                pending.pop(i)
                _save_json(PENDING_FILE, pending)
                st.info(f"Rejected request from {r.get('username')}")
                st.experimental_rerun()


# If logged in, show quick link to Home
if st.session_state.get("logged_in"):
    safe_divider()
    st.success(f"Logged in as {st.session_state.get('user')}")
    # page_link may not exist in all Streamlit versions; guard to avoid AttributeError
    if hasattr(st, "page_link"):
        st.page_link("Home.py", label="Go to Home →", icon="🗂️")
    else:
        if st.button("Go to Home →"):
            # best-effort: set query params if the API exists
            if hasattr(st, "experimental_set_query_params"):
                st.experimental_set_query_params(page="Home.py")
            else:
                # fallback: do nothing — avoid raising attribute errors
                pass
else:
    st.info("Not logged in. Use the form above or request access.")
