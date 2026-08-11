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

st.set_page_config(
    page_title="HR Assistant",
    page_icon="🗂️",
    layout="wide",
)

render_top_nav("Home")

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
