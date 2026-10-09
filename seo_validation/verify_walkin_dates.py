"""
Verification of the Walk-In DATE RULE (common/walkin_data.py) — offline,
synthetic values only:
    python seo_validation\\verify_walkin_dates.py
"""
import datetime as dt
import os
import sys
import tempfile
import types

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "common"))
import paths                                                    # noqa: E402
paths.LOGS_DIR = paths.Path(tempfile.mkdtemp(prefix="seo_logs_"))
import walkin_data as WD                                        # noqa: E402

FAIL = []
D = dt.date


def check(label, got, want):
    ok = got == want
    print(f"[{'pass' if ok else 'FAIL'}] {label}: {got!r}" + ("" if ok else f"  (expected {want!r})"))
    if not ok:
        FAIL.append(label)


def ser(d, h=14, m=48, s=38):
    return (dt.datetime.combine(d, dt.time(h, m, s)) - dt.datetime(1899, 12, 30)).total_seconds() / 86400


print("== 1. Real date cells are read by VALUE (locale / display can't swap day and month) ==")
d, r, _ = WD._resolve_dates([ser(D(2026, 5, 2)), ser(D(2026, 5, 14)), ser(D(2026, 2, 5)),
                             (D(2026, 5, 13) - D(1899, 12, 30)).days])
check("serial values -> exact dates (2-May, 14-May, 5-Feb, whole-number serial 13-May)", d,
      [D(2026, 5, 2), D(2026, 5, 14), D(2026, 2, 5), D(2026, 5, 13)])
check("… rule 'date value'", set(r), {"date value"})

print("\n== 2. Text dates — unambiguous ==")
vals = ["14/05/2026", "5/14/2026 10:00:00", "14-05-2026", "14.05.2026", "2026-05-14", "14-May-2026",
        "3rd May 2026", "16-07-2025 : 2 :00", "5/5/26", "31/12/25"]
d, r, _ = WD._resolve_dates(vals)
check("day > 12 or month > 12 decides; ISO / month names parsed", d,
      [D(2026, 5, 14)] * 6 + [D(2026, 5, 3), D(2025, 7, 16), D(2026, 5, 5), D(2025, 12, 31)])
check("rules", r[:4] + r[8:9],
      ["text, day/month (unambiguous)", "text, month/day (unambiguous)", "text, day/month (unambiguous)",
       "text, day/month (unambiguous)", "text, day = month"])

print("\n== 3. Text dates — AMBIGUOUS (both parts <= 12) ==")
ev = ["13/01/2026", "14/02/2026", "15/03/2026", "16/04/2026", "17/05/2026", "05/06/2026"]
d, r, _ = WD._resolve_dates(ev)
check("i. tab evidence: 5 unambiguous '/' dates are day/month -> 05/06/2026 = 5-Jun", (d[-1], "tab evidence 5/5" in r[-1]),
      (D(2026, 6, 5), True))
ev_md = ["1/13/2026", "2/14/2026", "3/15/2026", "4/16/2026", "5/17/2026", "05/06/2026"]
d, r, _ = WD._resolve_dates(ev_md)
check("i. tab evidence month/day -> 05/06/2026 = 6-May", d[-1], D(2026, 5, 6))
d, r, _ = WD._resolve_dates(["13/01/2026", "1/14/2026", "15/01/2026", "1/16/2026", "17/01/2026", "1/18/2026", "05/06/2026"])
check("i. mixed evidence (< 90 % agree) is not used -> DEFAULT", ("(tab evidence" in r[-1], "DEFAULT" in r[-1]), (False, True))
d, r, _ = WD._resolve_dates([ser(D(2026, 6, 3)), "05/06/2026", ser(D(2026, 6, 8))])
check("ii. row sequence: neighbours 3-Jun / 8-Jun -> 5-Jun", (d[1], "rows above/below" in r[1]), (D(2026, 6, 5), True))
d, r, _ = WD._resolve_dates([ser(D(2026, 5, 3)), "05/06/2026", ser(D(2026, 6, 8))])
check("ii. both readings fit the neighbours -> not decided by sequence", "rows above/below" in r[1], False)
d, r, _ = WD._resolve_dates(["05/06/2026", "05-06-2026", "05.06.2026"])
check("iii. default by separator: '/' month/day, '-' and '.' day/month (flagged DEFAULT)",
      (d, all("DEFAULT" in x for x in r)), ([D(2026, 5, 6), D(2026, 6, 5), D(2026, 6, 5)], True))

print("\n== 4. records_from_rows: diagnostics + blanks ==")
H = ["Timestamp", "First and last name", "Phone number", "How did you hear about us?", "Course Interested", "Background"]
rows = [H, [ser(D(2026, 5, 14)), "A", "9000000001", "Google", "x", "y"], ["", "", "", "", "", ""],
        ["14/05/2026", "B", "9000000002", "Google", "x", "y"], ["", "C no date", "9000000003", "Google", "x", "y"]]
diag = []
recs = WD.records_from_rows("Walk-In Old", rows, "old", diag)
check("blank row skipped, undated row kept with no date",
      [(x["name"], x["date"]) for x in recs], [("A", D(2026, 5, 14)), ("B", D(2026, 5, 14)), ("C no date", None)])
check("diag = sheet row, raw value, rule (one per record)",
      [(g["row"], g["rule"]) for g in diag], [(2, "date value"), (4, "text, day/month (unambiguous)"), (5, "blank")])

print("\n== 5. The report loader reads the tabs BY VALUE ==")
calls = []
fake = types.ModuleType("google_utils")
fake.read_tab = lambda sheets, sid, tab, date_render="FORMATTED_STRING": (calls.append((tab, date_render)) or rows)
sys.modules["google_utils"] = fake
WD.load_walkins(None, "sid", ["Walk-In New", "Walk-In Old"])
check("load_walkins requests SERIAL_NUMBER for both tabs", calls,
      [("Walk-In New", "SERIAL_NUMBER"), ("Walk-In Old", "SERIAL_NUMBER")])

print("\nALL CHECKS PASSED" if not FAIL else f"\n{len(FAIL)} CHECK(S) FAILED: {FAIL}")
sys.exit(1 if FAIL else 0)
