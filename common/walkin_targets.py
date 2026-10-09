"""
Walk-In lead performance vs target (common/walkin_targets.py).

Pure functions (no Google access) shared by the e-mail and the workbook, so the
"Weekly / Monthly Walk-In Lead Performance" e-mail section and the "Weekly /
Monthly Lead Trend" tabs always show the SAME figures.

TARGETS
  Monthly target  = MONTHLY_WALKIN_LEAD_TARGET in config/walkin_target.yaml
                    (read + validated by common/target_config.py), Walk-In
                    leads per calendar month. Passed in by the caller — this
                    module holds no target value of its own.
  Weekly target   = the monthly target spread evenly over the days of each
                    calendar month, summed over the 7 days (Mon–Sun) of the week:

                        weekly_target = Σ  monthly_target / days_in_month(d)
                                        d ∈ Mon..Sun

                    A week inside a 31-day month = 7 × T / 31; inside a 30-day
                    month = 7 × T / 30; inside February (28 days) = 7 × T / 28
                    (T = 100 -> 22.6 / 23.3 / 25.0); a week
                    that crosses a month end takes each day from its own month.
                    Summed over a whole month, the daily shares add back up to
                    exactly the monthly target, so weekly and monthly targets
                    reconcile.

COUNTING
  Uses the report's own normalised, de-duplicated Walk-In records
  (walkin_data.load_walkins): one record = one walk-in, counted in the week /
  month of its date. Records without a date are not counted (same rule as every
  other section of the report).

PERIODS
  The reference date is the report's own `period.cur_end` (today, or the chosen
  WEEKLY_REFERENCE_DATE / MONTHLY_MONTH). The week / month containing it is the
  "current" one and is counted up to the reference date — exactly the report's
  Current total. Earlier weeks / months are complete.

STATUS / COLOUR
  Actual > target  -> "Above Target"  (green)
  Actual <= target -> "Below Target"  (red)  — "at or below", as specified.
  The current, still-running week / month is marked "In progress" and carries
  its days elapsed and its pro-rata target to date (for context only — the
  colour still follows Actual vs the full target).
"""
from __future__ import annotations
import calendar
import datetime as _dt

ABOVE, BELOW = "Above Target", "Below Target"


# ── targets ──────────────────────────────────────────────────────────────────
def daily_target(d: _dt.date, monthly_target: float) -> float:
    return monthly_target / calendar.monthrange(d.year, d.month)[1]


def target_for_days(start: _dt.date, end: _dt.date, monthly_target: float) -> float:
    """Σ monthly_target / days_in_month(d) over start..end (inclusive)."""
    t, d = 0.0, start
    while d <= end:
        t += daily_target(d, monthly_target)
        d += _dt.timedelta(days=1)
    return t


def week_start(d: _dt.date) -> _dt.date:
    return d - _dt.timedelta(days=d.weekday())            # Monday


def _month_start(d: _dt.date, back: int = 0) -> _dt.date:
    m = d.month - 1 - back
    y = d.year + m // 12
    return _dt.date(y, m % 12 + 1, 1)


def _month_end(ms: _dt.date) -> _dt.date:
    return ms.replace(day=calendar.monthrange(ms.year, ms.month)[1])


# ── counting ─────────────────────────────────────────────────────────────────
def count_between(records, start: _dt.date, end: _dt.date) -> int:
    return sum(1 for r in records if r.get("date") and start <= r["date"] <= end)


def status(actual: float, target: float) -> str:
    return ABOVE if actual > target else BELOW


def _row(records, kind, start, end, asof, monthly_target):
    """One week / month: full-period target, actual counted up to min(end, asof)."""
    current = start <= asof <= end
    upto = min(end, asof)
    actual = count_between(records, start, upto)
    target = target_for_days(start, end, monthly_target)
    days_total = (end - start).days + 1
    days_elapsed = (upto - start).days + 1
    in_progress = current and asof < end
    return {
        "kind": kind, "start": start, "end": end, "counted_to": upto,
        "actual": actual, "target": round(target, 1), "target_exact": target,
        "achievement": (actual / target * 100.0) if target else None,
        "status": status(actual, target), "above": actual > target,
        "current": current, "in_progress": in_progress,
        "days_elapsed": days_elapsed, "days_total": days_total,
        "target_to_date": round(target_for_days(start, upto, monthly_target), 1),
    }


def weekly_trend(records, asof: _dt.date, n_completed: int,
                 monthly_target: float) -> list:
    """Last `n_completed` full Mon–Sun weeks + the week containing `asof`
    (oldest first)."""
    cws = week_start(asof)
    out = []
    for k in range(n_completed, -1, -1):
        ws = cws - _dt.timedelta(days=7 * k)
        out.append(_row(records, "week", ws, ws + _dt.timedelta(days=6), asof, monthly_target))
    return out


def monthly_trend(records, asof: _dt.date, n_completed: int,
                  monthly_target: float) -> list:
    """Last `n_completed` calendar months + the month containing `asof`
    (oldest first). Each month's target = the monthly target."""
    out = []
    for k in range(n_completed, -1, -1):
        ms = _month_start(asof, k)
        out.append(_row(records, "month", ms, _month_end(ms), asof, monthly_target))
    return out


def trend_for_period(records, period, monthly_target,
                     n_email=5, n_tab=11):
    """{'email': rows, 'tab': rows, 'kind': 'week'|'month', 'monthly_target'} for a
    Weekly / Monthly report period, else None (Manual has no trend)."""
    lab = str(getattr(period, "label", "")).lower()
    asof = period.cur_end
    if lab == "weekly":
        fn, kind = weekly_trend, "week"
    elif lab == "monthly":
        fn, kind = monthly_trend, "month"
    else:
        return None
    tab = fn(records, asof, n_tab, monthly_target)
    return {"kind": kind, "monthly_target": monthly_target, "asof": asof,
            "tab": tab, "email": tab[-(n_email + 1):]}


# ── Monthly performance bands (Red / Amber / Green) ──────────────────────────
# Share of the configured monthly target (config/walkin_target.yaml) reached:
#   Red    below 60 %            Amber  60 % to 80 % (both inclusive)
#   Green  above 80 %
# Used by the "Weekly Lead Trend" and "Monthly Lead Trend" tabs (table colours +
# chart bars) — a week against its own derived weekly target, a month against the
# monthly target. The Above / Below Target status above is unchanged (e-mail, the
# "completed periods above target" card).
RAG_RED_BELOW = 60.0
RAG_GREEN_ABOVE = 80.0
BAND_RED, BAND_AMBER, BAND_GREEN = "red", "amber", "green"
BAND_LABELS = {BAND_RED: "Below 60% of target",
               BAND_AMBER: "60–80% of target",
               BAND_GREEN: "Above 80% of target"}


def achievement_pct(actual, target):
    """actual / target in % (None when there is no target), rounded to 6 dp so a
    float target such as 100.00000000000003 never moves a month across a band."""
    if not target:
        return None
    return round(actual / float(target) * 100.0, 6)


def perf_band(actual, target) -> str:
    """'red' (< 60 %), 'amber' (60–80 % inclusive) or 'green' (> 80 %) of target."""
    pct = achievement_pct(actual, target)
    if pct is None or pct < RAG_RED_BELOW:
        return BAND_RED
    return BAND_AMBER if pct <= RAG_GREEN_ABOVE else BAND_GREEN


# ── labels ───────────────────────────────────────────────────────────────────
def period_label(row) -> str:
    """'Mon 29-Sep – Sun 05-Oct-2026' / 'Sep-2026'."""
    if row["kind"] == "month":
        return row["start"].strftime("%b-%Y")
    s, e = row["start"], row["end"]
    return f"{s.strftime('%a %d-%b')} – {e.strftime('%a %d-%b-%Y')}"


def short_label(row) -> str:
    """Compact chart axis label; '*' marks the current, still-running period."""
    base = row["start"].strftime("%b-%y") if row["kind"] == "month" else row["start"].strftime("%d-%b")
    return base + ("*" if row["in_progress"] else "")


def progress_text(row) -> str:
    unit = "day" if row["days_total"] == 1 else "days"
    return (f"In progress · {row['days_elapsed']} of {row['days_total']} {unit} · "
            f"counted to {row['counted_to'].strftime('%d-%b')} · "
            f"target to date {row['target_to_date']:g}")


def target_formula_text(kind, monthly_target) -> str:
    if kind == "month":
        return f"Monthly target: {monthly_target:g} Walk-In leads per calendar month."
    return (f"Weekly target = monthly target ({monthly_target:g}) ÷ days in the month, for each "
            f"day Mon–Sun, summed (e.g. 7 × {monthly_target:g}/31 = "
            f"{7 * monthly_target / 31:.1f} in a 31-day month, "
            f"{7 * monthly_target / 30:.1f} in a 30-day month).")
