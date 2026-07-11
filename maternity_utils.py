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

from datetime import date, timedelta


def compute_maternity_payment(
    monthly_salary: float,
    leave_start_date: date,
    leave_end_date: date,
    max_entitled_days: int = 182,
    already_paid_dates: set = None,
    daily_rate_divisor: int = 30,
) -> dict:
    """
    1. actual_days   = calendar days from leave_start_date to leave_end_date, inclusive.
    2. capped_days   = min(actual_days, max_entitled_days) — the leave span can never
                       be paid for more than the entitlement (182 days by default),
                       even if the requested end date is further out.
    3. payable_days  = capped_days minus any already-paid dates that fall inside the
                       capped window (dates outside the window don't count — they
                       weren't going to be paid by this calculation anyway).
    4. daily_rate    = monthly_salary / daily_rate_divisor — a flat "per day on a
                       monthly basis" rate (default divisor 30), not a rate that
                       varies by how many days each specific calendar month has.
    5. total_amount  = payable_days * daily_rate.
    """
    already_paid_dates = already_paid_dates or set()

    if leave_end_date < leave_start_date:
        raise ValueError("Leave End Date must be on or after Leave Start Date.")
    if daily_rate_divisor <= 0:
        raise ValueError("Daily rate divisor must be greater than zero.")

    actual_days = (leave_end_date - leave_start_date).days + 1
    was_capped = actual_days > max_entitled_days
    capped_days = min(actual_days, max_entitled_days)
    effective_end_date = leave_start_date + timedelta(days=capped_days - 1)

    # Only dates that actually fall within the (possibly capped) payable window count.
    relevant_excluded = {d for d in already_paid_dates if leave_start_date <= d <= effective_end_date}
    decremented_days = len(relevant_excluded)

    payable_days = max(capped_days - decremented_days, 0)

    daily_rate = monthly_salary / daily_rate_divisor
    total_amount = round(payable_days * daily_rate, 2)

    return {
        "leave_start_date": leave_start_date,
        "leave_end_date": leave_end_date,
        "actual_days_requested": actual_days,
        "max_entitled_days": max_entitled_days,
        "was_capped": was_capped,
        "effective_end_date": effective_end_date,
        "decremented_days": decremented_days,
        "payable_days": payable_days,
        "daily_rate": round(daily_rate, 2),
        "total_amount": total_amount,
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
