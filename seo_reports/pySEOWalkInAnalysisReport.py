"""
IntelliBI SEO Walk-In Analysis — report generator.

Builds management-level Walk-In analytics reports (Summary + Lead Source Trend +
Technology Trend + Lead Type Trend) from the 'Walk-In New' and 'Walk-In Old'
tabs, comparing the current period against the SAME elapsed days of the previous
comparable period.

HOW TO RUN
----------
Just run the file — no command-line parameters needed. Control everything from the
USER SETTINGS block below (which reports, dates, email on/off):

    python seo_reports/pySEOWalkInAnalysisReport.py

Command-line flags are still accepted and OVERRIDE the settings for that one run
(handy for ad-hoc runs), e.g.:
    python seo_reports/pySEOWalkInAnalysisReport.py --mode manual --start 2026-08-01 --end 2026-08-15
    python seo_reports/pySEOWalkInAnalysisReport.py --no-upload --no-email

Everything else (source sheet, Drive parent folder / sub-folder names, email
recipients & sender, normalization, week start, timezone) lives in config/.
"""
from __future__ import annotations
import argparse
import calendar
import datetime as _dt
import os
import sys

# ── make common/ importable (portable; no machine-specific paths) ────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
_COMMON = os.path.join(os.path.dirname(_HERE), "common")
if _COMMON not in sys.path:
    sys.path.insert(0, _COMMON)

import paths                     # noqa: E402
import config_loader as cfg      # noqa: E402
import logging_utils             # noqa: E402
import report_periods            # noqa: E402
import report_builder            # noqa: E402
import walkin_data               # noqa: E402

log = logging_utils.get_logger("pySEOWalkInAnalysisReport")

# ═══════════════════════════════════════════════════════════════════════════
#  USER SETTINGS  — edit these; run the file with no parameters.
# ═══════════════════════════════════════════════════════════════════════════
GENERATE_WEEKLY  = True
GENERATE_MONTHLY = True
GENERATE_MANUAL  = False          # Manual = a custom start/end date range

# Optional reference periods (None -> use today / current week / current month).
WEEKLY_REFERENCE_DATE = None      # "YYYY-MM-DD" — any day in the wanted week
MONTHLY_MONTH         = None      # 1-12  (None -> current month, to date)
MONTHLY_YEAR          = None      # e.g. 2026 (None -> current year)
MANUAL_START_DATE     = None      # "YYYY-MM-DD"  (required when GENERATE_MANUAL)
MANUAL_END_DATE       = None      # "YYYY-MM-DD"  (required when GENERATE_MANUAL)

EMAIL_SEND = True                 # email the report(s)? (recipients live in config.yaml)
UPLOAD_TO_DRIVE = True            # upload the report(s) to Drive?
# ═══════════════════════════════════════════════════════════════════════════


def _now_stamp() -> str:
    tz = cfg.get("report.timezone", "Asia/Kolkata")
    try:
        from zoneinfo import ZoneInfo
        now = _dt.datetime.now(ZoneInfo(tz))
    except Exception:
        now = _dt.datetime.now()
    return now.strftime("%d-%b-%Y %I:%M %p")


def _today() -> _dt.date:
    tz = cfg.get("report.timezone", "Asia/Kolkata")
    try:
        from zoneinfo import ZoneInfo
        return _dt.datetime.now(ZoneInfo(tz)).date()
    except Exception:
        return _dt.date.today()


def _parse_opt(s):
    """Parse a 'YYYY-MM-DD' string (or already-a-date) into a date; else None."""
    if not s:
        return None
    if isinstance(s, _dt.date):
        return s
    try:
        return _dt.datetime.strptime(str(s).strip(), "%Y-%m-%d").date()
    except ValueError:
        log.warning("Ignoring invalid date setting %r (want YYYY-MM-DD).", s)
        return None


def _window(records, start, end):
    return [r for r in records if r["date"] and start <= r["date"] <= end]


def _build_one(period, records, gen_stamp, sample=False):
    cur = _window(records, period.cur_start, period.cur_end)
    prev = _window(records, period.prev_start, period.prev_end)
    gs = cfg.get("report.google_search_label", "Google Search")
    wb = report_builder.build_workbook(period, cur, prev, gs, gen_stamp, sample=sample)
    paths.ensure_dirs()
    out_path = os.path.join(str(paths.OUTPUT_DIR), period.fname() + ".xlsx")
    wb.save(out_path)
    log.info("%s | %s | current=%d previous=%d | saved %s",
             period.label, period.range_str(), len(cur), len(prev), os.path.basename(out_path))
    return out_path, cur, prev


_folder_cache = {}


def _folder_for(drive, period):
    """Resolve the Drive folder for this report type. An explicit folder ID in
    config wins; otherwise the sub-folder NAME is resolved (and created if missing)
    under drive.parent_folder_id. Result is cached per run."""
    key = period.label.lower()
    explicit = (cfg.get("drive.folders", {}) or {}).get(key, "")
    if explicit:
        return explicit
    parent = cfg.get("drive.parent_folder_id", "")
    if not (parent and drive is not None):
        return ""
    if key in _folder_cache:
        return _folder_cache[key]
    name = (cfg.get("drive.subfolders", {}) or {}).get(key, period.label)
    try:
        import google_utils
        fid = google_utils.find_or_create_folder(drive, parent, name)
        _folder_cache[key] = fid
        return fid
    except Exception as e:
        log.warning("  could not resolve Drive sub-folder '%s' under parent: %s", name, e)
        return ""


def _period_suffix(period):
    """Reporting-period tag appended to the Drive report name so each period is a
    distinct, self-describing file. Monthly -> 'Sep-2026'; Weekly/Manual -> the
    current window range 'DD-Mon-YYYY to DD-Mon-YYYY'."""
    if period.label.lower() == "monthly":
        return period.cur_start.strftime("%b-%Y")
    return (f"{period.cur_start.strftime('%d-%b-%Y')} to "
            f"{period.cur_end.strftime('%d-%b-%Y')}")


def _upload(drive, out_path, period):
    folder = _folder_for(drive, period)
    if not (cfg.get("drive.upload", True) and folder and drive is not None):
        return None
    try:
        import google_utils
        names = cfg.get("drive.report_names", {}) or {}
        base = (names.get(period.label.lower()) or "").strip()
        # Append the reporting period to the configured base name. If no base name
        # is configured, fall back to the already-dated output file name.
        name = f"{base} {_period_suffix(period)}" if base \
            else os.path.splitext(os.path.basename(out_path))[0]
        fid, link = google_utils.upload_xlsx(drive, out_path, name, folder)
        log.info("  uploaded to Drive as '%s': %s", name, link)
        return link
    except Exception as e:
        log.warning("  Drive upload skipped/failed: %s", e)
        return None


def _email(period, cur, prev, drive_link, gen_stamp, out_path, send_email):
    if not send_email:
        return
    try:
        import email_utils
        gs = cfg.get("report.google_search_label", "Google Search")
        email_utils.send_report(period, cur, prev, gs, drive_link, gen_stamp, out_path)
    except Exception as e:
        log.warning("  email step skipped/failed: %s", e)


def _monthly_asof(today):
    """as-of date that drives the Monthly report from the MONTHLY_MONTH/_YEAR
    settings: current month -> today (month-to-date); a chosen past/other month ->
    that month in full."""
    if not MONTHLY_MONTH:
        return today
    y = int(MONTHLY_YEAR) if MONTHLY_YEAR else today.year
    m = int(MONTHLY_MONTH)
    if y == today.year and m == today.month:
        return today
    last = calendar.monthrange(y, m)[1]
    return _dt.date(y, m, last)


def _resolve_periods(cli_mode, cli_as_of, cli_start, cli_end):
    """Build the list of Period objects from the USER SETTINGS, with any CLI flags
    overriding for this run."""
    week_start = cfg.get("report.week_starts_on", "monday")
    today = cli_as_of or _today()

    if cli_mode and cli_mode != "all":
        want = {cli_mode}
    else:
        want = set()
        if GENERATE_WEEKLY:
            want.add("weekly")
        if GENERATE_MONTHLY:
            want.add("monthly")
        if GENERATE_MANUAL:
            want.add("manual")

    periods = []
    if "weekly" in want:
        ref = cli_as_of or _parse_opt(WEEKLY_REFERENCE_DATE) or today
        periods.append(report_periods.weekly(ref, week_start))
    if "monthly" in want:
        periods.append(report_periods.monthly(cli_as_of or _monthly_asof(today)))
    if "manual" in want:
        ms = cli_start or _parse_opt(MANUAL_START_DATE)
        me = cli_end or _parse_opt(MANUAL_END_DATE)
        if ms and me:
            periods.append(report_periods.manual(ms, me))
        else:
            log.error("Manual is enabled but start/end dates are missing "
                      "(set MANUAL_START_DATE & MANUAL_END_DATE, or pass --start/--end).")
    return periods


def run(periods, upload=True, send_email=True):
    logging_utils.section(log, "IntelliBI SEO Walk-In Analysis")
    if not periods:
        log.warning("Nothing to generate — enable a report in USER SETTINGS "
                    "(GENERATE_WEEKLY / GENERATE_MONTHLY / GENERATE_MANUAL).")
        return True
    gen_stamp = _now_stamp()

    try:
        import google_utils
        sheets, drive = google_utils.get_services()
    except Exception as e:
        log.error("Google auth failed — cannot read the source sheet: %s", e)
        return False

    sid = cfg.get("source.spreadsheet_id", "")
    tabs = cfg.get("source.tabs", ["Walk-In New", "Walk-In Old"])
    records = walkin_data.load_walkins(sheets, sid, tabs)
    if not records:
        log.error("No Walk-In records loaded — aborting (nothing written).")
        return False

    ok = True
    for p in periods:
        try:
            out_path, cur, prev = _build_one(p, records, gen_stamp)
            link = _upload(drive, out_path, p) if upload else None
            _email(p, cur, prev, link, gen_stamp, out_path, send_email)
        except Exception as e:
            ok = False
            log.exception("%s report FAILED: %s", p.label, e)
    return ok


def _parse_date(s):
    return _dt.datetime.strptime(s, "%Y-%m-%d").date()


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="IntelliBI SEO Walk-In Analysis — edit USER SETTINGS at the top; "
                    "flags below are optional overrides for one run.")
    ap.add_argument("--mode", choices=["weekly", "monthly", "manual", "all"], default=None,
                    help="Override which report to generate (default: USER SETTINGS).")
    ap.add_argument("--as-of", type=_parse_date, default=None, help="Treat this date as 'today' (YYYY-MM-DD).")
    ap.add_argument("--start", type=_parse_date, default=None, help="Manual start date (YYYY-MM-DD).")
    ap.add_argument("--end", type=_parse_date, default=None, help="Manual end date (YYYY-MM-DD).")
    ap.add_argument("--no-upload", action="store_true", help="Do not upload to Google Drive.")
    ap.add_argument("--email", dest="email", action="store_true", default=None,
                    help="Force-send the report email (overrides EMAIL_SEND).")
    ap.add_argument("--no-email", dest="email", action="store_false",
                    help="Do not send any email (overrides EMAIL_SEND).")
    args = ap.parse_args(argv)

    periods = _resolve_periods(args.mode, args.as_of, args.start, args.end)
    upload = UPLOAD_TO_DRIVE and (not args.no_upload)
    send_email = args.email if args.email is not None else EMAIL_SEND
    ok = run(periods, upload=upload, send_email=send_email)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
