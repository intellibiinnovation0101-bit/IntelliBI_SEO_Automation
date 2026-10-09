"""
Verification of the Weekly / Monthly Walk-In Lead Performance vs target
(common/walkin_targets.py, the e-mail section in common/email_utils.py and the
"Weekly / Monthly Lead Trend" tab in common/report_builder.py).

Run from the project root (no Google access, synthetic walk-ins only):
    python seo_validation\\verify_walkin_lead_trend.py
"""
import datetime as dt
import io
import os
import random
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "common"))
import paths                                                    # noqa: E402
paths.LOGS_DIR = paths.Path(tempfile.mkdtemp(prefix="seo_logs_"))   # never write the real logs/
import openpyxl                                                 # noqa: E402
import report_periods as RP                                     # noqa: E402
import walkin_targets as WT                                     # noqa: E402
import report_builder as RB                                     # noqa: E402
import email_utils as EU                                        # noqa: E402
import target_config as TC                                      # noqa: E402

# The target under test is read from config/walkin_target.yaml (the single
# source), so this check follows any future change to MONTHLY_WALKIN_LEAD_TARGET.
T = TC.load_monthly_target()
print(f"MONTHLY_WALKIN_LEAD_TARGET (from {TC.TARGET_FILE.name}) = {T}")

FAIL = []


def check(label, got, want):
    ok = got == want
    print(f"[{'pass' if ok else 'FAIL'}] {label}: {got!r}" + ("" if ok else f"  (expected {want!r})"))
    if not ok:
        FAIL.append(label)


D = dt.date
# ── 1. targets ───────────────────────────────────────────────────────────────
print("== 1. Targets ==")
check("31-day month week = 7 × T/31", round(WT.target_for_days(D(2026, 10, 5), D(2026, 10, 11), T), 3),
      round(7 * T / 31, 3))
check("30-day month week = 7 × T/30", round(WT.target_for_days(D(2026, 9, 7), D(2026, 9, 13), T), 3),
      round(7 * T / 30, 3))
check("February week = 7 × T/28", round(WT.target_for_days(D(2027, 2, 8), D(2027, 2, 14), T), 3), round(T / 4, 3))
check("weekly target derived from the script's monthly target (31 / 30 / 28-day months, 1 dp)",
      [round(WT.target_for_days(a, b, T), 1) for a, b in
       ((D(2026, 10, 5), D(2026, 10, 11)), (D(2026, 9, 7), D(2026, 9, 13)), (D(2027, 2, 8), D(2027, 2, 14)))],
      [round(7 * T / 31, 1), round(7 * T / 30, 1), round(7 * T / 28, 1)])
check("the Weekly target follows the monthly one (a different T gives a different week)",
      round(WT.target_for_days(D(2026, 10, 5), D(2026, 10, 11), T * 2), 6),
      round(2 * WT.target_for_days(D(2026, 10, 5), D(2026, 10, 11), T), 6))
check("week across a month end takes each day from its own month (Mon 28-Sep … Sun 04-Oct)",
      round(WT.target_for_days(D(2026, 9, 28), D(2026, 10, 4), T), 4),
      round(3 * T / 30 + 4 * T / 31, 4))
check("daily targets of a whole month add back to the monthly target (Oct, Feb, Sep)",
      [round(WT.target_for_days(D(2026, m, 1), D(2026, m, n), T), 6) for m, n in ((10, 31), (2, 28), (9, 30))],
      [float(T)] * 3)
check("status: > target green 'Above', = target red 'Below', < target red",
      [WT.status(17, 16.3), WT.status(16.3, 16.3), WT.status(10, 16.3)], [WT.ABOVE, WT.BELOW, WT.BELOW])

# ── 2. synthetic walk-ins ────────────────────────────────────────────────────
rnd = random.Random(7)
recs = []
d = D(2025, 9, 1)
while d <= D(2026, 10, 31):
    for _ in range(rnd.randint(0, 5)):
        recs.append({"date": d, "lead_source": "Google Search", "technology": "Power BI",
                     "lead_type": "Student", "mobile": "", "name": "x", "tab": "Walk-In New"})
    d += dt.timedelta(days=1)
recs.append({"date": None, "lead_source": "", "technology": "", "lead_type": "", "mobile": "",
             "name": "no date", "tab": "Walk-In New"})        # never counted (same as the report)
cnt = lambda a, b: sum(1 for r in recs if r["date"] and a <= r["date"] <= b)

print("\n== 2. Weekly trend (as of Thu 08-Oct-2026) ==")
pw = RP.weekly(D(2026, 10, 8))
tw = WT.trend_for_period(recs, pw, T, n_email=5, n_tab=11)
check("12 weeks in the tab, 6 in the e-mail, the e-mail = the tab's last 6",
      (len(tw["tab"]), len(tw["email"]), tw["email"] == tw["tab"][-6:]), (12, 6, True))
check("every week is Monday → Sunday, consecutive, oldest first",
      all(x["start"].weekday() == 0 and (x["end"] - x["start"]).days == 6 for x in tw["tab"])
      and all((b["start"] - a["start"]).days == 7 for a, b in zip(tw["tab"], tw["tab"][1:])), True)
check("first / current week", (tw["tab"][0]["start"], tw["tab"][-1]["start"]), (D(2026, 7, 20), D(2026, 10, 5)))
cur = tw["tab"][-1]
check("current week: counted Mon 05-Oct → Thu 08-Oct = the report's Current total",
      (cur["in_progress"], cur["days_elapsed"], cur["actual"]),
      (True, 4, cnt(pw.cur_start, pw.cur_end)))
check("… full-week target (not pro-rata) and target to date for context",
      (cur["target"], cur["target_to_date"]), (round(7 * T / 31, 1), round(4 * T / 31, 1)))
check("completed weeks counted in full, never twice (sum of 12 weeks = all walk-ins in the range)",
      sum(x["actual"] for x in tw["tab"]), cnt(D(2026, 7, 20), D(2026, 10, 8)))
check("historical reference date (WEEKLY_REFERENCE_DATE = Sun 27-Sep-2026): that week is complete",
      (lambda t: (t["tab"][-1]["start"], t["tab"][-1]["in_progress"], t["tab"][-1]["actual"]))(
          WT.trend_for_period(recs, RP.weekly(D(2026, 9, 27)), T)),
      (D(2026, 9, 21), False, cnt(D(2026, 9, 21), D(2026, 9, 27))))

print("\n== 3. Monthly trend ==")
pm = RP.monthly(D(2026, 10, 8))
tm = WT.trend_for_period(recs, pm, T, n_email=5, n_tab=11)
check("12 months (Nov-2025 … Oct-2026), 6 in the e-mail",
      ([WT.period_label(x) for x in (tm["tab"][0], tm["tab"][-1])], len(tm["email"])), (["Nov-2025", "Oct-2026"], 6))
check("every month's target = T, current month in progress (counted to 08-Oct)",
      ({x["target"] for x in tm["tab"]}, tm["tab"][-1]["in_progress"], tm["tab"][-1]["actual"]),
      ({float(T)}, True, cnt(D(2026, 10, 1), D(2026, 10, 8))))
check("a past month chosen (MONTHLY_MONTH = 9) -> Sep-2026 complete",
      (lambda t: (WT.period_label(t["tab"][-1]), t["tab"][-1]["in_progress"]))(
          WT.trend_for_period(recs, RP.monthly(D(2026, 9, 30)), T)), ("Sep-2026", False))
check("Manual report -> no trend", WT.trend_for_period(recs, RP.manual(D(2026, 9, 1), D(2026, 9, 15)), T), None)

# ── 4. workbook ──────────────────────────────────────────────────────────────
print("\n== 4. Workbook tab ==")
_win = lambda p: ([r for r in recs if r["date"] and p.cur_start <= r["date"] <= p.cur_end],
                  [r for r in recs if r["date"] and p.prev_start <= r["date"] <= p.prev_end])
for p, t, name in ((pw, tw, "Weekly Lead Trend"), (pm, tm, "Monthly Lead Trend")):
    c, pv = _win(p)
    wb = RB.build_workbook(p, c, pv, "Google Search", "08-Oct-2026 07:00 PM", trend=t)
    check(f"{p.label}: existing tabs unchanged, new tab last",
          wb.sheetnames, ["Summary", "Lead Source Trend", "Technology Trend", "Lead Type Trend", name])
    ws = wb[name]
    hdr = next(r for r in range(1, 40) if ws.cell(r, 1).value == "#")
    rows = [[ws.cell(r, c).value for c in range(1, 7)] for r in range(hdr + 1, hdr + 13)]
    Per = "Weekly" if p.label == "Weekly" else "Monthly"
    unit = "week" if p.label == "Weekly" else "month"
    bands = [WT.perf_band(x["actual"], x["target"]) for x in t["tab"]]
    check(f"{p.label}: 12 rows = the trend (actual, target, band text)",
          [(r[2], r[3], r[5].split(" · ")[0]) for r in rows],
          [(x["actual"], x["target"], WT.BAND_LABELS[b]) for x, b in zip(t["tab"], bands)])
    check(f"{p.label}: row colours Red / Amber / Green by % of each period's own target",
          [ws.cell(hdr + 1 + i, 3).fill.fgColor.rgb[-6:] for i in range(12)], [RB.F_RAG[b] for b in bands])
    check(f"{p.label}: Achievement % = Walk-In Leads / {Per} Target as shown (no float noise)",
          [round(r[4], 8) for r in rows], [round(x["actual"] / x["target"], 8) for x in t["tab"]])
    cf = {str(k.sqref): [(x.formula[0], x.dxf.fill.fgColor.rgb[-6:]) for x in v]
          for k, v in ws.conditional_formatting._cf_rules.items()}
    rng = f"A{hdr + 1}:F{hdr + 12}"
    check(f"{p.label}: real conditional formatting on the table rows (3 rules, red/amber/green fills)",
          (list(cf), [f for _, f in cf.get(rng, [])]), ([rng], [RB.F_RAG["red"], RB.F_RAG["amber"], RB.F_RAG["green"]]))
    check(f"{p.label}: conditional-format formulas use the row's own leads / target cells",
          [fm for fm, _ in cf[rng]],
          [f"AND(ISNUMBER($C{hdr+1}),ISNUMBER($D{hdr+1}),$D{hdr+1}>0,$C{hdr+1}/$D{hdr+1}<0.6)",
           f"AND(ISNUMBER($C{hdr+1}),ISNUMBER($D{hdr+1}),$D{hdr+1}>0,$C{hdr+1}/$D{hdr+1}>=0.6,$C{hdr+1}/$D{hdr+1}<=0.8)",
           f"AND(ISNUMBER($C{hdr+1}),ISNUMBER($D{hdr+1}),$D{hdr+1}>0,$C{hdr+1}/$D{hdr+1}>0.8)"])
    check(f"{p.label}: current row marked", ("▶ current" in rows[-1][1], "In progress" in rows[-1][5]), (True, True))
    tot = [ws.cell(hdr + 13, c).value for c in range(2, 7)]
    check(f"{p.label}: completed total row = sum of the 11 completed rows + colour counts",
          (tot[1], tot[4].split("  ·  ")[1]),
          (sum(x["actual"] for x in t["tab"][:-1]),
           f"Green {bands[:11].count('green')} · Amber {bands[:11].count('amber')} · Red {bands[:11].count('red')}"))
    ch = ws._charts
    cur_i = next(i for i, x in enumerate(t["tab"]) if x["in_progress"])
    _sol = lambda gp: gp.solidFill.srgbClr.val if hasattr(gp.solidFill.srgbClr, "val") else gp.solidFill.srgbClr
    check(f"{p.label}: one clean column chart — 3 bar series (red / amber / green) + dashed target line",
          (len(ch), ch[0].grouping, [_sol(s.graphicalProperties) for s in ch[0].series], len(ch[0]._charts),
           ch[0]._charts[1].series[0].graphicalProperties.line.dashStyle),
          (1, "clustered", [RB.C_RAG["red"], RB.C_RAG["amber"], RB.C_RAG["green"]], 2, "dash"))
    check(f"{p.label}: value labels centred in every bar (clear of the target line)",
          [(s.dLbls.showVal, s.dLbls.position) for s in ch[0].series], [(True, "ctr")] * 3)
    helper = [[ws.cell(hdr + 1 + i, 32 + j).value for j in range(5)] for i in range(12)]
    check(f"{p.label}: chart data — exactly one bar per {unit}, in its band's series, target alongside",
          all(sum(v is not None for v in h[1:4]) == 1 and h[4] == x["target"]
              and h[1 + ("red", "amber", "green").index(b)] == x["actual"]
              for h, x, b in zip(helper, t["tab"], bands)), True)
    dpts = [(i, len(s.dPt)) for i, s in enumerate(ch[0].series) if s.dPt]
    cser = ch[0].series[("red", "amber", "green").index(bands[-1])]
    check(f"{p.label}: current {unit} = lighter bar with dashed outline + dark label, labelled '(to date)'",
          (helper[-1][0].endswith("(to date)"), dpts, cser.dPt[0].idx if dpts else None,
           cur_i in [d.idx for d in cser.dLbls.dLbl]),
          (True, [(("red", "amber", "green").index(bands[-1]), 1)], cur_i, True))
    check(f"{p.label}: target line visible inside the value axis (max above target and every bar)",
          ch[0].y_axis.scaling.max > max([x["actual"] for x in t["tab"]] + [x["target"] for x in t["tab"]]), True)
    tt = sorted({x["target"] for x in t["tab"]})
    check(f"{p.label}: chart heading and legend text",
          (ws.cell(next(r for r in range(hdr, hdr + 30) if str(ws.cell(r, 1).value or "").startswith("Walk-In Leads vs")), 1).value,
           [str(ws.cell(hdr, 32 + j).value) for j in range(1, 5)]),
          (f"Walk-In Leads vs {Per} Target — Last 11 Completed {unit.title()}s + Current {unit.title()}",
           ["Below 60% of target", "60–80% of target", "Above 80% of target",
            "Weekly target" if Per == "Weekly" else f"Monthly target ({T:g})"]))
    buf = io.BytesIO(); wb.save(buf)
    check(f"{p.label}: workbook saves and re-opens", name in openpyxl.load_workbook(io.BytesIO(buf.getvalue())).sheetnames, True)
print("\n== 4b. Red / Amber / Green bands (share of the period's target) ==")
check("band edges: 59.9 red · 60 amber · 80 amber · 80.1 green (target 100)",
      [WT.perf_band(a, 100) for a in (0, 59.9, 60, 70, 80, 80.1, 100, 150)],
      ["red", "red", "amber", "amber", "amber", "green", "green", "green"])
check("bands follow the configured target (T=50: 29 red, 30 amber, 40 amber, 41 green)",
      [WT.perf_band(a, 50) for a in (29, 30, 40, 41)], ["red", "amber", "amber", "green"])
check("float target noise never flips a band (70 / 100.00000000000003 -> amber, 80 / 99.99999999999997 -> amber)",
      [WT.perf_band(70, 100.00000000000003), WT.perf_band(80, 99.99999999999997)], ["amber", "amber"])
check("no target -> red (never green by accident)", WT.perf_band(5, 0), "red")
_tm_alt = WT.trend_for_period(recs, pm, 50.0, n_email=5, n_tab=11)
_c, _pv = _win(pm)
_ws_alt = RB.build_workbook(pm, _c, _pv, "Google Search", "x", trend=_tm_alt)["Monthly Lead Trend"]
_h = next(r for r in range(1, 40) if _ws_alt.cell(r, 1).value == "#")
check("another target (50): table colours recomputed from it",
      [_ws_alt.cell(_h + 1 + i, 3).fill.fgColor.rgb[-6:] for i in range(12)],
      [RB.F_RAG[WT.perf_band(x["actual"], 50.0)] for x in _tm_alt["tab"]])
_tw_alt = WT.trend_for_period(recs, pw, 50.0, n_email=5, n_tab=11)
_ws_walt = RB.build_workbook(pw, *_win(pw), "Google Search", "x", trend=_tw_alt)["Weekly Lead Trend"]
_h = next(r for r in range(1, 40) if _ws_walt.cell(r, 1).value == "#")
check("another target (50): Weekly colours recomputed from each week's derived target",
      [_ws_walt.cell(_h + 1 + i, 3).fill.fgColor.rgb[-6:] for i in range(12)],
      [RB.F_RAG[WT.perf_band(x["actual"], x["target"])] for x in _tw_alt["tab"]])
check("a 22.6-lead week (31-day month, target 100): 13 red · 14 amber · 18 amber · 19 green",
      [WT.perf_band(a, 22.6) for a in (13, 14, 18, 19)], ["red", "amber", "amber", "green"])
_wn = RB.build_workbook(pw, *_win(pw), "Google Search", "x", trend=tw)["Weekly Lead Trend"].cell(4, 1).value
check("explanation line states the bands in whole leads (22.6-lead week: 13 or fewer · 14–18 · 19 or more)",
      ("(13 leads or fewer)" in _wn, "(14–18 leads)" in _wn, "(19 leads or more)" in _wn), (True, True, True))
_mn = RB.build_workbook(pm, *_win(pm), "Google Search", "x", trend=tm)["Monthly Lead Trend"].cell(4, 1).value
check("… and for a month at target 100: 59 or fewer · 60–80 · 81 or more",
      ("(59 leads or fewer)" in _mn, "(60–80 leads)" in _mn, "(81 leads or more)" in _mn), (True, True, True))

c, pv = _win(RP.manual(D(2026, 9, 1), D(2026, 9, 15)))
check("Manual workbook: no Lead Trend tab (unchanged)",
      RB.build_workbook(RP.manual(D(2026, 9, 1), D(2026, 9, 15)), c, pv, "Google Search", "x").sheetnames,
      ["Summary", "Lead Source Trend", "Technology Trend", "Lead Type Trend"])

# ── 5. e-mail ────────────────────────────────────────────────────────────────
print("\n== 5. E-mail ==")
c, pv = _win(pw)
html_w = EU.build_body(pw, c, pv, "Google Search", "https://x", "08-Oct-2026", trend=tw)
check("weekly e-mail: section title",
      "Weekly Walk-In Lead Performance &ndash; Last 5 Completed Weeks + Current Week" in html_w, True)
check("weekly e-mail: the 6 weeks' figures = the tab's last 6 rows",
      all(f"{x['actual']} / {x['target']:g}" in html_w for x in tw["tab"][-6:]), True)
check("weekly e-mail: current week flagged 'In progress', newest first",
      (html_w.index("In progress") < html_w.index(WT.period_label(tw["tab"][-2]))), True)
c, pv = _win(pm)
html_m = EU.build_body(pm, c, pv, "Google Search", "https://x", "08-Oct-2026", trend=tm)
check("monthly e-mail: section title + the script's target",
      ("Monthly Walk-In Lead Performance &ndash; Last 5 Completed Months + Current Month" in html_m,
       f"Monthly target {T:g}" in html_m), (True, True))
base_dir = os.environ.get("SEO_BASE_DIR")
if base_dir:                        # optional: compare with the pre-change module
    import importlib.util
    spec = importlib.util.spec_from_file_location("eu_old", os.path.join(base_dir, "common", "email_utils.py"))
    old = importlib.util.module_from_spec(spec); spec.loader.exec_module(old)
    check("no trend (Manual / old callers): e-mail body identical to before",
          EU.build_body(pw, c, pv, "Google Search", "https://x", "s"), old.build_body(pw, c, pv, "Google Search", "https://x", "s"))

print("\n== 6. Single source of the target ==")
import re                                                       # noqa: E402
_src = {p: open(os.path.join(ROOT, p), encoding="utf-8").read() for p in
        ("seo_reports/pySEOWalkInAnalysisReport.py", "common/walkin_targets.py", "common/report_builder.py",
         "common/email_utils.py", "config/config.yaml")}       # (target_config.py = the loader itself)
check("no other target setting (script variable / config.yaml key / module default) remains",
      [p for p, s in _src.items() if re.search(
          r"^\s*MONTHLY_WALKIN_LEAD_TARGET\s*=|^\s*monthly_walkin_(lead_)?target\s*:|DEFAULT_MONTHLY_TARGET", s, re.M | re.I)], [])
_tf = open(TC.TARGET_FILE, encoding="utf-8").read()
check("config/walkin_target.yaml holds the one target line",
      len(re.findall(r"^MONTHLY_WALKIN_LEAD_TARGET\s*:", _tf, re.M)), 1)

# ── 7. validation rules ──────────────────────────────────────────────────────
print("\n== 7. Target file validation (PyYAML and the built-in fallback parser) ==")
_tmp = tempfile.mkdtemp(prefix="seo_target_")


def _try(body, name="t.yaml"):
    p = os.path.join(_tmp, name)
    if body is not None:
        with open(p, "w", encoding="utf-8", newline="") as f:
            f.write(body)
    warns = []
    try:
        return TC.load_monthly_target(p, warn=warns.append), warns
    except TC.TargetConfigError as e:
        return "ERR: " + str(e), warns


GOOD = [("integer", "MONTHLY_WALKIN_LEAD_TARGET: 100\n", 100),
        ("CRLF + comments + trailing comment", "# c\r\nMONTHLY_WALKIN_LEAD_TARGET: 120   # per month\r\n", 120),
        ("Notepad BOM", "\ufeffMONTHLY_WALKIN_LEAD_TARGET: 90\n", 90),
        ("100.0 -> 100", "MONTHLY_WALKIN_LEAD_TARGET: 100.0\n", 100),
        ("quoted \"100\"", 'MONTHLY_WALKIN_LEAD_TARGET: "100"\n', 100),
        ("lower-case name", "monthly_walkin_lead_target: 75\n", 75),
        ("bounds 1 / 10000", "MONTHLY_WALKIN_LEAD_TARGET: 10000\n", 10000)]
BAD = [("missing file", None, "not found"),
       ("empty file", "", "has no"),
       ("only comments", "# nothing\n", "has no"),
       ("key missing (typo)", "MONTHLY_WALKIN_TARGET: 100\n", "does not contain"),
       ("blank value", "MONTHLY_WALKIN_LEAD_TARGET:\n", "no value"),
       ("null", "MONTHLY_WALKIN_LEAD_TARGET: null\n", "no value"),
       ("boolean", "MONTHLY_WALKIN_LEAD_TARGET: true\n", "not a number"),
       ("text", "MONTHLY_WALKIN_LEAD_TARGET: abc\n", "not a number"),
       ("thousands comma", "MONTHLY_WALKIN_LEAD_TARGET: 1,000\n", "not a number"),
       ("percent", "MONTHLY_WALKIN_LEAD_TARGET: 100%\n", "not a number"),
       ("decimal", "MONTHLY_WALKIN_LEAD_TARGET: 100.5\n", "whole number"),
       ("zero", "MONTHLY_WALKIN_LEAD_TARGET: 0\n", "out of range"),
       ("negative", "MONTHLY_WALKIN_LEAD_TARGET: -5\n", "out of range"),
       ("too large", "MONTHLY_WALKIN_LEAD_TARGET: 10001\n", "out of range"),
       ("NaN", "MONTHLY_WALKIN_LEAD_TARGET: .nan\n", "not"),
       ("duplicate line", "MONTHLY_WALKIN_LEAD_TARGET: 100\nMONTHLY_WALKIN_LEAD_TARGET: 120\n", "more than once"),
       ("list value", "MONTHLY_WALKIN_LEAD_TARGET:\n  - 100\n", "single number"),
       ("python syntax instead of YAML", "MONTHLY_WALKIN_LEAD_TARGET = 100\n", "")]
_real_yaml = sys.modules.get("yaml")
for parser in ("PyYAML", "fallback"):
    if parser == "fallback":
        sys.modules["yaml"] = None          # makes 'import yaml' raise ImportError
    try:
        for i, (lab, body, want) in enumerate(GOOD):
            check(f"[{parser}] valid: {lab}", _try(body, f"g{i}.yaml")[0], want)
        for i, (lab, body, frag) in enumerate(BAD):
            got = _try(body, f"b{i}.yaml" if body is not None else "absent.yaml")[0]
            check(f"[{parser}] rejected: {lab}", isinstance(got, str) and got.startswith("ERR") and frag in got, True)
        v, w = _try("MONTHLY_WALKIN_LEAD_TARGET: 100\nWEEKLY_TARGET: 20\n", "w.yaml")
        check(f"[{parser}] unknown name ignored with a warning", (v, len(w), "WEEKLY_TARGET" in (w or [""])[0]), (100, 1, True))
    finally:
        if parser == "fallback":
            if _real_yaml is None:
                sys.modules.pop("yaml", None)
            else:
                sys.modules["yaml"] = _real_yaml
print("   e.g. " + _try("MONTHLY_WALKIN_LEAD_TARGET: 100.5\n", "eg.yaml")[0])

# ── 8. report run behaviour ──────────────────────────────────────────────────
print("\n== 8. Report run uses the file; a bad file stops Weekly / Monthly safely ==")
import types                                                    # noqa: E402
sys.path.insert(0, os.path.join(ROOT, "seo_reports"))
import pySEOWalkInAnalysisReport as REP                         # noqa: E402
fake_g = types.ModuleType("google_utils"); fake_g.get_services = lambda: (None, None)
sys.modules["google_utils"] = fake_g
built, mailed = [], []
REP.walkin_data.load_walkins = lambda *a, **k: recs
REP._email = lambda *a, **k: mailed.append(a[0].label)
_orig_build = REP._build_one


def _spy_build(p, records, gen_stamp, sample=False, target=None):
    built.append((p.label, target))
    return _orig_build(p, records, gen_stamp, sample=sample, target=target)


REP._build_one = _spy_build
paths.OUTPUT_DIR = paths.Path(tempfile.mkdtemp(prefix="seo_out_"))
REP.paths.OUTPUT_DIR = paths.OUTPUT_DIR
_ok = REP.run([RP.weekly(D(2026, 10, 8)), RP.monthly(D(2026, 10, 8))], upload=False, send_email=False)
check("valid file: Weekly + Monthly built with the file's target, e-mail step reached",
      (_ok, built, mailed), (True, [("Weekly", T), ("Monthly", T)], ["Weekly", "Monthly"]))
_real_file = TC.TARGET_FILE
for lab, body in (("invalid value", "MONTHLY_WALKIN_LEAD_TARGET: abc\n"), ("missing file", None)):
    built.clear(); mailed.clear()
    TC.TARGET_FILE = paths.Path(_tmp) / ("run_bad.yaml" if body is not None else "run_absent.yaml")
    if body is not None:
        TC.TARGET_FILE.write_text(body, encoding="utf-8")
    try:
        _ok = REP.run([RP.weekly(D(2026, 10, 8)), RP.monthly(D(2026, 10, 8)),
                       RP.manual(D(2026, 9, 1), D(2026, 9, 15))], upload=False, send_email=False)
        check(f"{lab}: Weekly / Monthly skipped (no workbook / e-mail), Manual still runs, run reports failure",
              (_ok, [b[0] for b in built], mailed), (False, ["Manual"], ["Manual"]))
        check(f"{lab}: --check-target exits 1", REP.main(["--check-target"]), 1)
    finally:
        TC.TARGET_FILE = _real_file
check("--check-target with the real file exits 0", REP.main(["--check-target"]), 0)

print("\nALL CHECKS PASSED" if not FAIL else f"\n{len(FAIL)} CHECK(S) FAILED: {FAIL}")
sys.exit(1 if FAIL else 0)
