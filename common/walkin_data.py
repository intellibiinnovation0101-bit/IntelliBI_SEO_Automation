"""
Walk-In data layer (common/walkin_data.py).

Reads the 'Walk-In New' and 'Walk-In Old' tabs, maps their differently-worded
columns to ONE normalized schema, normalizes category values centrally, parses
dates robustly, and de-duplicates so the same walk-in is never double-counted.

Normalized record:
    { date, lead_source, technology, lead_type, mobile, name, tab }
"""
from __future__ import annotations
import datetime as _dt
import re
import config_loader as cfg
import seo_normalize as norm
import logging_utils

log = logging_utils.get_logger("seo_walkin_data")

_DEFAULT_MAP = {
    "new": {
        "date": "Timestamp",
        "lead_source": "How did you hear about IntelliBI?",
        "technology": "Which technology are you interested in learning?",
        "lead_type": "Current Status",
        "mobile": "Mobile Number",
        "name": "Full Name",
    },
    "old": {
        "date": "Timestamp",
        "lead_source": "How did you hear about us?",
        "technology": "Course Interested",
        "lead_type": "Background",
        "mobile": "Phone number",
        "name": "First and last name",
    },
}

_DATE_FORMATS = (
    "%m/%d/%Y %H:%M:%S", "%m/%d/%Y", "%d-%b-%Y %I:%M %p", "%d-%b-%Y",
    "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
    "%m/%d/%Y %H:%M", "%d-%m-%Y",
)


def _norm_header(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip()).lower()


def _col_index(header_row: list, wanted: str):
    w = _norm_header(wanted)
    for i, h in enumerate(header_row):
        if _norm_header(h) == w:
            return i
    return None


def parse_date(v):
    """Return a datetime.date from a string, datetime, or Excel serial; else None."""
    if v is None or v == "":
        return None
    if isinstance(v, _dt.datetime):
        return v.date()
    if isinstance(v, _dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return (_dt.datetime(1899, 12, 30) + _dt.timedelta(days=float(v))).date()
        except Exception:
            return None
    s = str(v).strip()
    if not s:
        return None
    for f in _DATE_FORMATS:
        try:
            return _dt.datetime.strptime(s, f).date()
        except ValueError:
            continue
    return None


def _digits10(v) -> str:
    d = re.sub(r"\D", "", str(v or ""))
    return d[-10:] if len(d) >= 10 else d


def _mapping():
    m = cfg.get("field_mapping", None)
    if isinstance(m, dict) and m.get("new") and m.get("old"):
        return m
    return _DEFAULT_MAP


def records_from_rows(tab_name: str, rows: list, which: str) -> list:
    """Convert raw tab rows to normalized records. `which` is 'new' or 'old'."""
    if not rows:
        return []
    header = rows[0]
    fmap = _mapping()[which]
    idx = {k: _col_index(header, v) for k, v in fmap.items()}
    missing = [k for k, i in idx.items() if i is None]
    if missing:
        log.warning("%s: columns not found for %s (check field_mapping)", tab_name, missing)
    out = []
    for r in rows[1:]:
        if not any(str(x).strip() for x in r if x is not None):
            continue

        def cell(key):
            i = idx.get(key)
            return r[i] if (i is not None and i < len(r)) else ""

        out.append({
            "date": parse_date(cell("date")),
            "lead_source": norm.norm_lead_source(cell("lead_source")),
            "technology": norm.norm_technology(cell("technology")),
            "lead_type": norm.norm_lead_type(cell("lead_type")),
            "mobile": _digits10(cell("mobile")),
            "name": norm.clean(cell("name")),
            "tab": tab_name,
        })
    return out


def dedupe(records: list) -> tuple:
    """Drop duplicates keyed by (mobile, date) when a mobile is present."""
    seen = set()
    kept = []
    removed = 0
    for rec in records:
        if rec["mobile"] and rec["date"]:
            key = (rec["mobile"], rec["date"])
            if key in seen:
                removed += 1
                continue
            seen.add(key)
        kept.append(rec)
    return kept, removed


def load_walkins(sheets, spreadsheet_id: str, tabs: list) -> list:
    """Read every configured tab and return combined, de-duplicated records."""
    import google_utils
    all_recs = []
    for tab in tabs:
        which = "new" if "new" in tab.lower() else "old"
        rows = google_utils.read_tab(sheets, spreadsheet_id, tab)
        recs = records_from_rows(tab, rows, which)
        log.info("  %-14s: %d data rows -> %d records", tab, max(len(rows) - 1, 0), len(recs))
        all_recs += recs
    deduped, removed = dedupe(all_recs)
    log.info("Combined %d records; de-duplicated to %d (removed %d)",
             len(all_recs), len(deduped), removed)
    return deduped
