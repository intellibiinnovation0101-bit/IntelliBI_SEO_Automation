"""
Report period logic (common/report_periods.py).

Produces the current window and a like-for-like previous window that uses the
SAME number of elapsed days — so a partial current week/month is never compared
against a full previous one.

  Weekly  : current week start -> as_of ; previous week, same elapsed days.
  Monthly : current month start -> as_of ; previous month, same elapsed days
            (clamped if the previous month is shorter).
  Manual  : start -> end ; the immediately-preceding equal-length window.

week_starts_on: 'monday' (default) or 'sunday'.
Each Period carries labelled date ranges for the report header/filenames.
"""
from __future__ import annotations
import calendar
import datetime as _dt
from dataclasses import dataclass


@dataclass
class Period:
    label: str            # "Weekly" | "Monthly" | "Manual"
    cur_start: _dt.date
    cur_end: _dt.date
    prev_start: _dt.date
    prev_end: _dt.date

    @property
    def elapsed_days(self) -> int:
        return (self.cur_end - self.cur_start).days + 1

    def fname(self) -> str:
        return (f"IntelliBI SEO Walk-In Analysis - {self.label} - "
                f"{self.cur_start.strftime('%d-%b-%Y')} to {self.cur_end.strftime('%d-%b-%Y')}")

    def range_str(self) -> str:
        return (f"Current: {self.cur_start.strftime('%d-%b-%Y')} → {self.cur_end.strftime('%d-%b-%Y')}"
                f"   vs   Previous: {self.prev_start.strftime('%d-%b-%Y')} → {self.prev_end.strftime('%d-%b-%Y')}")


def _week_start(d: _dt.date, week_starts_on: str) -> _dt.date:
    if str(week_starts_on).lower().startswith("sun"):
        return d - _dt.timedelta(days=(d.weekday() + 1) % 7)
    return d - _dt.timedelta(days=d.weekday())        # Monday-based


def _minus_days(d: _dt.date, n: int) -> _dt.date:
    return d - _dt.timedelta(days=n)


def _add_months_start(d: _dt.date, months: int) -> _dt.date:
    """First day of the month that is `months` before/after d's month."""
    m = d.month - 1 + months
    y = d.year + m // 12
    m = m % 12 + 1
    return _dt.date(y, m, 1)


def weekly(as_of: _dt.date, week_starts_on: str = "monday") -> Period:
    cs = _week_start(as_of, week_starts_on)
    elapsed = (as_of - cs).days
    ps = _minus_days(cs, 7)
    pe = ps + _dt.timedelta(days=elapsed)
    return Period("Weekly", cs, as_of, ps, pe)


def monthly(as_of: _dt.date) -> Period:
    cs = as_of.replace(day=1)
    elapsed = (as_of - cs).days
    ps = _add_months_start(cs, -1)
    last_prev = calendar.monthrange(ps.year, ps.month)[1]
    pe_day = min(elapsed + 1, last_prev)
    pe = ps.replace(day=pe_day)
    return Period("Monthly", cs, as_of, ps, pe)


def manual(start: _dt.date, end: _dt.date) -> Period:
    if end < start:
        start, end = end, start
    length = (end - start).days
    pe = _minus_days(start, 1)
    ps = pe - _dt.timedelta(days=length)
    return Period("Manual", start, end, ps, pe)
