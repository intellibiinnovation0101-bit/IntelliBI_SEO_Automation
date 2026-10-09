"""
Walk-In data layer (common/walkin_data.py).

Reads the 'Walk-In New' and 'Walk-In Old' tabs, maps their differently-worded
columns to ONE normalized schema, normalizes category values centrally, dates
every row by the explicit DATE RULE below (real date cells by value; text dates
with an evidence-based day/month rule), and de-duplicates so the same walk-in is
never double-counted.

Normalized record:
    { date, lead_source, technology, lead_type, mobile, name, tab }
"""
from __future__ import annotations
import collections
import datetime as _dt
import re
import config_loader as cfg
import seo_normalize as norm
import logging_utils

try:
    from dateutil import parser as _du          # robust, tolerant date parsing
except Exception:                               # pragma: no cover
    _du = None

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

# Fixed formats are the fast path; anything not matching falls through to the
# tolerant dateutil parser in parse_date() below so no real walk-in is dropped.
_DATE_FORMATS = (
    "%m/%d/%Y %H:%M:%S", "%m/%d/%Y", "%d-%b-%Y %I:%M %p", "%d-%b-%Y",
    "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
    "%m/%d/%Y %H:%M", "%d-%m-%Y", "%d.%m.%Y", "%d-%B-%Y",
)

_ORDINAL_RE = re.compile(r"(?<=\d)(st|nd|rd|th)\b", re.I)


def _norm_header(s) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip()).lower()


def _col_index(header_row: list, wanted: str):
    w = _norm_header(wanted)
    for i, h in enumerate(header_row):
        if _norm_header(h) == w:
            return i
    return None


def parse_date(v):
    """Return a datetime.date from a string, datetime, or Excel serial; else None.

    Tolerant by design: walk-in 'Timestamp' cells come from Google Forms
    (US 'M/D/YYYY h:mm:ss') AND from staff who hand-type Indian-style dates
    ('23-09-2026', '23.09.2026', '14-July-2025', '16-07-2025 : 2 :00', '3rd Sep
    2026', Excel serials). Previously any value outside a short fixed list was
    silently dropped, which removed real walk-ins from the report. We now:
      1) try the fast fixed formats, then
      2) fall back to dateutil with a SEPARATOR-AWARE day/month order so an
         ambiguous '05-09-2026' is read as 5-Sep (Indian), not 9-May, while a
         slash date '9/5/2026' stays US month-first (Google Form).
    """
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
    for f in _DATE_FORMATS:                       # fast path
        try:
            return _dt.datetime.strptime(s, f).date()
        except ValueError:
            continue

    if _du is None:                               # dateutil unavailable
        return None
    # normalise common dirt, then let dateutil handle the rest
    t = re.sub(r"\s*:\s*", ":", s)                # "2 :00" -> "2:00"
    t = t.replace("@", " ")
    t = _ORDINAL_RE.sub("", t)
    t = re.sub(r"\s+", " ", t).strip(" ,")
    if not t:
        return None
    dayfirst = "/" not in t                       # slash=US(M/D); dash/dot/text=Indian(D/M)
    default = _dt.datetime(_dt.date.today().year, 1, 1)
    for df in (dayfirst, not dayfirst):
        try:
            return _du.parse(t, dayfirst=df, fuzzy=True, default=default).date()
        except (ValueError, OverflowError, TypeError):
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


# ── DATE RULE (explicit) ──────────────────────────────────────────────────────
# 1. A real DATE cell — every Google Form "Timestamp" — is read by its VALUE
#    (load_walkins reads the tabs with dateTimeRenderOption=SERIAL_NUMBER), so the
#    sheet's locale / display format can never swap day and month.
# 2. A hand-typed TEXT date is read as follows:
#    a) ISO "2026-05-14", or a month written as a word ("14-May-2026", "3rd Sep
#       2026") -> unambiguous (parse_date).
#    b) all-numeric "A<sep>B<sep>YEAR" (sep / - .):  A > 12 -> A is the day;
#       B > 12 -> B is the day; A == B -> no question.  -> unambiguous.
#    c) both A and B <= 12 and different -> AMBIGUOUS, resolved in this order:
#       i.   the TAB'S OWN EVIDENCE: the day/month order used by that tab's
#            unambiguous all-numeric text dates with the same separator, when there
#            are >= AMBIG_MIN_EVIDENCE of them and >= AMBIG_MIN_SHARE agree;
#       ii.  ROW SEQUENCE: the reading that falls between the dates of the nearest
#            dated rows above and below (+/- AMBIG_NEIGHBOUR_DAYS), used only when
#            exactly one reading fits;
#       iii. DEFAULT by separator: "/" month/day (Google Sheets' US display),
#            "-" and "." day/month (Indian hand-typing) — logged as a default.
#    Every decision is counted in the log and, per row, in records_from_rows(diag=).
AMBIG_MIN_EVIDENCE = 5
AMBIG_MIN_SHARE = 0.9
AMBIG_NEIGHBOUR_DAYS = 7
_NUM_DATE_RE = re.compile(r"^\s*(\d{1,2})\s*([/\-.])\s*(\d{1,2})\s*\2\s*(\d{2}|\d{4})(?!\d)")


def _ymd(y, m, d):
    y = y + 2000 if y < 100 else y
    try:
        return _dt.date(y, m, d)
    except ValueError:
        return None


def classify_text_date(s):
    """('numeric', date, order, sep) for an unambiguous all-numeric text date
    (order 'DM' / 'MD' / 'same'); ('ambiguous', {'MD': date, 'DM': date}, None, sep);
    ('other', None, None, None) for anything else (parse_date handles it)."""
    m = _NUM_DATE_RE.match(str(s or ""))
    if not m:
        return "other", None, None, None
    a, sep, b, y = int(m.group(1)), m.group(2), int(m.group(3)), int(m.group(4))
    md, dm = _ymd(y, a, b), _ymd(y, b, a)
    if a == b:
        return ("numeric", md, "same", sep) if md else ("other", None, None, None)
    if md and not dm:
        return "numeric", md, "MD", sep
    if dm and not md:
        return "numeric", dm, "DM", sep
    if md and dm:
        return "ambiguous", {"MD": md, "DM": dm}, None, sep
    return "other", None, None, None


def _default_order(sep):
    return "MD" if sep == "/" else "DM"


def _resolve_dates(raw_values):
    """Apply the DATE RULE to one tab's Timestamp column (in sheet row order).
    Returns ([date|None], [rule text], stats)."""
    n = len(raw_values)
    dates, rules, pending = [None] * n, [""] * n, []
    evidence = collections.defaultdict(collections.Counter)      # sep -> Counter(DM/MD)
    stats = collections.Counter()
    for i, v in enumerate(raw_values):
        if v is None or (isinstance(v, str) and not v.strip()):
            rules[i] = "blank"; stats["blank"] += 1
            continue
        if isinstance(v, (int, float, _dt.date)) and not isinstance(v, bool):
            dates[i] = parse_date(v)
            rules[i] = "date value" if dates[i] else "unparsed"
            stats["date value" if dates[i] else "unparsed"] += 1
            continue
        kind, val, order, sep = classify_text_date(v)
        if kind == "numeric":
            dates[i] = val
            rules[i] = {"DM": "text, day/month (unambiguous)", "MD": "text, month/day (unambiguous)",
                        "same": "text, day = month"}[order]
            if order in ("DM", "MD"):
                evidence[sep][order] += 1
            stats["text unambiguous"] += 1
        elif kind == "ambiguous":
            pending.append((i, val, sep))
        else:
            dates[i] = parse_date(v)
            rules[i] = "text, parsed" if dates[i] else "unparsed"
            stats["text unambiguous" if dates[i] else "unparsed"] += 1
    fixed = list(dates)                                          # anchors for the row-sequence step
    for i, cand, sep in pending:
        ev = evidence.get(sep, collections.Counter())
        tot = ev["DM"] + ev["MD"]
        top = ev.most_common(1)[0] if tot else (None, 0)
        if tot >= AMBIG_MIN_EVIDENCE and top[1] / tot >= AMBIG_MIN_SHARE:
            dates[i] = cand[top[0]]
            rules[i] = (f"text, AMBIGUOUS -> {'day/month' if top[0] == 'DM' else 'month/day'} "
                        f"(tab evidence {top[1]}/{tot} '{sep}' dates)")
            stats["ambiguous: tab evidence"] += 1
            continue
        prev = next((fixed[j] for j in range(i - 1, -1, -1) if fixed[j]), None)
        nxt = next((fixed[j] for j in range(i + 1, n) if fixed[j]), None)
        if prev and nxt:
            lo_d = min(prev, nxt) - _dt.timedelta(days=AMBIG_NEIGHBOUR_DAYS)
            hi_d = max(prev, nxt) + _dt.timedelta(days=AMBIG_NEIGHBOUR_DAYS)
            fits = [k for k, d in cand.items() if lo_d <= d <= hi_d]
            if len(fits) == 1:
                dates[i] = cand[fits[0]]
                rules[i] = (f"text, AMBIGUOUS -> {'day/month' if fits[0] == 'DM' else 'month/day'} "
                            f"(rows above/below dated {prev:%d-%b-%Y} / {nxt:%d-%b-%Y})")
                stats["ambiguous: row sequence"] += 1
                continue
        k = _default_order(sep)
        dates[i] = cand[k]
        rules[i] = (f"text, AMBIGUOUS -> {'day/month' if k == 'DM' else 'month/day'} "
                    f"(DEFAULT for '{sep}' — no tab evidence or row sequence)")
        stats["ambiguous: default"] += 1
    return dates, rules, stats


def records_from_rows(tab_name: str, rows: list, which: str, diag: list = None) -> list:
    """Convert raw tab rows to normalized records. `which` is 'new' or 'old'.
    diag (optional list): receives one dict per returned record — sheet row, raw
    Timestamp and the DATE RULE step that dated it (used by the trace)."""
    if not rows:
        return []
    header = rows[0]
    fmap = _mapping()[which]
    idx = {k: _col_index(header, v) for k, v in fmap.items()}
    missing = [k for k, i in idx.items() if i is None]
    if missing:
        log.warning("%s: columns not found for %s (check field_mapping)", tab_name, missing)
    src = [(n, r) for n, r in enumerate(rows[1:], start=2)
           if any(str(x).strip() for x in r if x is not None)]

    def cell(r, key):
        i = idx.get(key)
        return r[i] if (i is not None and i < len(r)) else ""

    raw_dates = [cell(r, "date") for _, r in src]
    dates, rules, stats = _resolve_dates(raw_dates)
    if src:
        log.info("  %-14s: dates — %s", tab_name,
                 ", ".join(f"{k} {v}" for k, v in sorted(stats.items())) or "none")
    out = []
    for (n, r), d, rule, raw in zip(src, dates, rules, raw_dates):
        out.append({
            "date": d,
            "lead_source": norm.norm_lead_source(cell(r, "lead_source")),
            "technology": norm.norm_technology(cell(r, "technology")),
            "lead_type": norm.norm_lead_type(cell(r, "lead_type")),
            "mobile": _digits10(cell(r, "mobile")),
            "name": norm.clean(cell(r, "name")),
            "tab": tab_name,
        })
        if diag is not None:
            diag.append({"row": n, "raw_date": raw, "rule": rule})
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
        # real date cells by VALUE (see DATE RULE) — never by the locale's display text
        rows = google_utils.read_tab(sheets, spreadsheet_id, tab, date_render="SERIAL_NUMBER")
        recs = records_from_rows(tab, rows, which)
        log.info("  %-14s: %d data rows -> %d records", tab, max(len(rows) - 1, 0), len(recs))
        all_recs += recs
    deduped, removed = dedupe(all_recs)
    log.info("Combined %d records; de-duplicated to %d (removed %d)",
             len(all_recs), len(deduped), removed)
    return deduped
