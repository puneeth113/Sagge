"""
maternity_utils.py — Maternity payment calculation.

Standalone module, deliberately independent of utils.py so it can't affect
any existing calculation logic elsewhere in the app.

METHOD
    - Entitlement is a fixed number of calendar days (default 182, i.e. the
      26-week Maternity Benefit Act entitlement — pass a different value for
      other cases such as the 84-day/12-week entitlement for a third child).
    - Payment is calculated per calendar day using that day's own month:
      daily_rate = monthly_salary / (actual number of days in that month).
      This means a day in February is worth slightly more than a day in
      January for the same monthly salary — this is the "per day monthly
      basis" the requirement describes.
    - If any dates inside the entitlement window were already paid through
      another route (e.g. separately-paid leave that overlaps the
      maternity period), pass them in `already_paid_dates` — those specific
      days are excluded ("decremented") from the payable count and from
      the amount, without shifting the rest of the calculation.
"""

import calendar
from datetime import date, timedelta


def compute_maternity_payment(
    monthly_salary: float,
    leave_start_date: date,
    total_entitled_days: int = 182,
    already_paid_dates: set = None,
) -> dict:
    """Returns a dict with the leave window, decrement count, a month-by-month
    breakdown (since the daily rate differs per calendar month), and the
    total payable amount."""
    already_paid_dates = already_paid_dates or set()

    all_dates = [leave_start_date + timedelta(days=i) for i in range(total_entitled_days)]
    payable_dates = [d for d in all_dates if d not in already_paid_dates]
    decremented_days = total_entitled_days - len(payable_dates)

    by_month = {}
    for d in payable_dates:
        by_month.setdefault((d.year, d.month), []).append(d)

    breakdown = []
    total_amount = 0.0
    for (year, month), dates_in_month in sorted(by_month.items()):
        days_in_month = calendar.monthrange(year, month)[1]
        daily_rate = monthly_salary / days_in_month
        amount = daily_rate * len(dates_in_month)
        total_amount += amount
        breakdown.append({
            "Year": year,
            "Month": month,
            "Month Name": calendar.month_name[month],
            "Days in Month": days_in_month,
            "Payable Days This Month": len(dates_in_month),
            "Daily Rate (₹)": round(daily_rate, 2),
            "Amount (₹)": round(amount, 2),
        })

    return {
        "leave_start_date": leave_start_date,
        "leave_end_date": leave_start_date + timedelta(days=total_entitled_days - 1),
        "total_entitled_days": total_entitled_days,
        "decremented_days": decremented_days,
        "payable_days": len(payable_dates),
        "total_amount": round(total_amount, 2),
        "monthly_breakdown": breakdown,
    }


def parse_excluded_dates(raw_text: str) -> set:
    """Parses a free-text block of dates/ranges into a set of date objects.
    Accepts one entry per line, each either a single date (YYYY-MM-DD) or a
    range (YYYY-MM-DD:YYYY-MM-DD), inclusive on both ends. Blank lines and
    stray whitespace are ignored. Raises ValueError with a line-specific
    message on bad input, so the UI can show exactly what to fix."""
    excluded = set()
    if not raw_text:
        return excluded

    for line_no, raw_line in enumerate(raw_text.splitlines(), start=1):
        line = raw_line.strip()
        if not line:
            continue
        try:
            if ":" in line:
                start_str, end_str = [p.strip() for p in line.split(":", 1)]
                start_d = date.fromisoformat(start_str)
                end_d = date.fromisoformat(end_str)
                if end_d < start_d:
                    raise ValueError("range end is before range start")
                cur = start_d
                while cur <= end_d:
                    excluded.add(cur)
                    cur += timedelta(days=1)
            else:
                excluded.add(date.fromisoformat(line))
        except ValueError as e:
            raise ValueError(f"Line {line_no} ('{raw_line.strip()}') is not a valid date or range: {e}")

    return excluded
