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
    {"path": "pages/2 incentive validator.py", "label": "Incentive Validator", "icon": "💰"},
    {"path": "pages/3_payroll calculator.py", "label": "Payroll Calculator", "icon": "🧾"},
    {"path": "pages/4_Employee_Database.py", "label": "Employee Database", "icon": "👥"},
    {"path": "pages/5_Paysheet_Operations.py", "label": "Paysheet Operations", "icon": "📋"},
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
    and return the raw bytes, ready for st.download_button."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets.items():
            safe_name = str(sheet_name)[:31]  # Excel sheet name limit
            df.to_excel(writer, index=False, sheet_name=safe_name)
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
# 1. Long Absence Tracking
# --------------------------------------------------------------------------

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


def sample_attendance_muster() -> pd.DataFrame:
    """Generate a sample attendance muster matching organization format.
    
    Columns: Employee ID, Employee Name, Designation, Location, then 25 date columns
    with day-of-week headers and attendance codes.
    
    Attendance codes:
    - P: Present
    - A: Absent (flagged for tracking)
    - H: Holiday
    - R: Rest/Off Day
    - VL: Vacation Leave
    - LWP: Leave Without Pay
    - CL: Casual Leave
    - CL:P: Casual Leave + Present
    - OFF: Off
    - H:CL: Holiday + Casual Leave
    """
    from datetime import datetime, timedelta
    
    # Generate 25-day sample from June 1-25, 2024
    start_date = datetime(2024, 6, 1)
    dates = [start_date + timedelta(days=i) for i in range(25)]
    day_names = ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri"]
    
    # Column headers: Employee info + dates with day names
    columns = ["Employee ID", "Employee Name", "Designation", "Location"]
    date_headers = [f"{d.strftime('%d')} {day_names[d.weekday()]}" for d in dates]
    columns.extend(date_headers)
    
    # Sample employees with varied attendance patterns
    employees_data = [
        ["EM20230060183", "A DEVARAJA", "Driver", "Bangalore"],
        ["EM20235150023", "A MARY ARPITHA", "Primary Teacher", "Maths, Bangalore"],
        ["EM20225080294", "A RADHAMMA", "Maid", "Bangalore"],
        ["EM20235060049", "A RAJESHWARI", "Bus Attendant", "Bangalore"],
        ["EM20235180026", "A S LOKESHWARI", "Teacher", "HortCulture, Bangalore"],
        ["EM20225060178", "A SANGEETA", "Teacher", "Social Studies, Bangalore"],
        ["EM20245060094", "A SATHYA", "Primary Teacher", "Maths, Bangalore"],
        ["EM20235180024", "A SELVI", "Maid", "Bangalore"],
        ["EM20235040168", "A SHABANA GULZAR", "Teacher", "Bangalore"],
    ]
    
    # Sample attendance patterns with variety
    sample_pattern_1 = ["P", "P", "P", "P", "P", "H", "R", "P", "VL", "VL", "VL", "LWP", "P", "P", "P", "P", "P", "H", "R", "P", "P", "P", "P", "P", "P"]
    sample_pattern_2 = ["P", "P", "LWP", "LWP", "LWP", "H", "R", "LWP", "LWP", "LWP", "P", "P", "P", "P", "P", "H", "R", "P", "P", "A", "A", "P", "P", "P", "P"]
    sample_pattern_3 = ["A", "A", "A", "P", "P", "H", "R", "P", "CL", "CL", "CL", "P", "P", "P", "P", "H", "R", "P", "P", "P", "P", "P", "A", "A", "A"]
    sample_pattern_4 = ["P", "P", "P", "P", "P", "H", "R", "P", "P", "P", "P", "P", "P", "P", "P", "H", "R", "P", "P", "P", "P", "P", "P", "P", "P"]
    
    patterns = [sample_pattern_1, sample_pattern_2, sample_pattern_3, sample_pattern_4, sample_pattern_1]
    
    # Assign patterns cyclically
    data = []
    for i, emp in enumerate(employees_data):
        row = emp + patterns[i % len(patterns)]
        data.append(row)
    
    df = pd.DataFrame(data, columns=columns)
    return df


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

    # Ensure HRA and Other are Series with matching length
    hra = out[hra_col] if hra_col and hra_col in out.columns else pd.Series([0.0] * len(out), index=out.index)
    other = out[other_allow_col] if other_allow_col and other_allow_col in out.columns else pd.Series([0.0] * len(out), index=out.index)

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
# Reverse Calculations (In-Hand → Gross / Billing ↔ Amount)
# --------------------------------------------------------------------------

def compute_fulltime_from_inhand(
    in_hand_salary: float,
    pf_wage_cap: float = 15000,
    apply_pf_cap: bool = True,
    pf_employee_pct: float = 12.0,
    esic_threshold: float = 21000,
    esic_employee_pct: float = 0.75,
) -> dict:
    """Reverse calculation: Given in-hand (net) salary, compute gross salary,
    PF, and ESIC deductions.
    
    Uses iterative approach because PF wage cap affects the calculation:
    - Net = Gross - PF_Deduction - ESIC_Deduction
    - PF deduction depends on min(Basic, wage_cap), not Gross directly
    - For simplicity, assumes Basic ≈ Gross (no HRA/Other deductions)
    """
    # Start with an estimate: net ≈ gross / 1.13 (assuming 12% PF + 0.75% ESIC)
    gross = in_hand_salary / 0.87  # Initial estimate
    
    # Iterate to find exact gross
    for _ in range(10):
        pf_wage = min(gross, pf_wage_cap) if apply_pf_cap else gross
        pf_deduction = pf_wage * pf_employee_pct / 100
        
        esic_applicable = gross <= esic_threshold
        esic_deduction = (gross * esic_employee_pct / 100) if esic_applicable else 0
        
        total_deductions = pf_deduction + esic_deduction
        computed_net = gross - total_deductions
        
        # Adjust gross based on difference
        if abs(computed_net - in_hand_salary) < 0.5:
            break
        gross = gross + (in_hand_salary - computed_net)
    
    pf_wage = min(gross, pf_wage_cap) if apply_pf_cap else gross
    pf_deduction = pf_wage * pf_employee_pct / 100
    esic_applicable = gross <= esic_threshold
    esic_deduction = (gross * esic_employee_pct / 100) if esic_applicable else 0
    
    return {
        "In-Hand Salary": in_hand_salary,
        "Computed Gross": round(gross, 2),
        "PF (Employee)": round(pf_deduction, 2),
        "ESIC (Employee)": round(esic_deduction, 2),
        "Total Deductions": round(pf_deduction + esic_deduction, 2),
        "ESIC Applicable": esic_applicable,
    }


def compute_gig_amount_from_inhand(
    in_hand: float,
    tds_pct: float = 1.0,
) -> dict:
    """Reverse: Given in-hand amount for gig worker, compute billing amount.
    Assumption: In-hand = Billing - TDS (worker keeps the net after TDS deduction)
    So: Billing = In-Hand / (1 - TDS%)
    """
    if tds_pct >= 100:
        return {"error": "TDS % must be less than 100%"}
    
    billing = in_hand / (1 - tds_pct / 100)
    tds_amount = billing * tds_pct / 100
    
    return {
        "In-Hand Amount": round(in_hand, 2),
        "Monthly Billing Amount": round(billing, 2),
        "TDS Amount": round(tds_amount, 2),
    }


def compute_gig_inhand_from_billing(
    billing: float,
    tds_pct: float = 1.0,
) -> dict:
    """Given monthly billing amount, compute in-hand amount and TDS.
    In-Hand = Billing - TDS
    TDS = Billing * TDS%
    """
    tds_amount = billing * tds_pct / 100
    in_hand = billing - tds_amount
    
    return {
        "Monthly Billing Amount": round(billing, 2),
        "TDS Amount": round(tds_amount, 2),
        "In-Hand Amount": round(in_hand, 2),
    }