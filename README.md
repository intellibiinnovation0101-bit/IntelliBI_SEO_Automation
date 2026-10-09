# IntelliBI SEO Walk-In Analytics

Management-level **Weekly / Monthly / Manual** Walk-In analytics for the SEO team,
built from the *Student Inquiry Tracker* Google Sheet (`Walk-In New` + `Walk-In Old`
tabs). Each report answers, quickly:

1. Are Walk-Ins increasing or decreasing?
2. Is **Google Search** generating more or fewer Walk-Ins?
3. What % of Walk-Ins come from Google Search?
4. Which Lead Sources perform best?
5. Which technologies generate the most interest?
6. What background of candidates is visiting?
7. How does the current period compare *fairly* with the previous equivalent period?

## Reports

Each run produces a workbook with four tabs — **Summary**, **Lead Source Trend**,
**Technology Trend**, **Lead Type Trend** — with KPI cards, comparison tables and
charts. Google Search is highlighted throughout.

- **Weekly** — current week start → today, vs the previous week's *same elapsed days*.
- **Monthly** — current month start → today, vs the previous month's *same elapsed days*.
- **Manual** — an explicit start/end, vs the immediately-preceding equal-length window.

Comparisons are always like-for-like (a partial current period is never compared
against a full previous one).

## Walk-In lead performance vs target (Weekly / Monthly)
Added 09-Oct-2026. Manual reports and every existing tab / e-mail section are unchanged.

| | Weekly report | Monthly report |
|---|---|---|
| E-mail section | **Weekly Walk-In Lead Performance – Last 5 Completed Weeks + Current Week** | **Monthly Walk-In Lead Performance – Last 5 Completed Months + Current Month** |
| New tab (last) | **Weekly Lead Trend** — last 11 completed weeks + current week | **Monthly Lead Trend** — last 11 completed months + current month |
| Period | Monday – Sunday | calendar month |
| Target | see formula below | `MONTHLY_WALKIN_LEAD_TARGET` Walk-In leads per month (currently 100) |

Each row: period, Walk-In leads, target, Achievement % (= leads ÷ target), performance band. The current week / month is counted up to the report date (the same figure as the Summary's *Current* total), marked *In progress* ("▶ current" row, "(to date)" on the chart, lighter bar) with its days elapsed and its pro-rata target to date; its colour still compares against the full target. The tab adds three summary cards, a completed-periods total row and a chart.

**Red / Amber / Green (Weekly and Monthly Lead Trend tabs).** Each week / month is coloured by the share of ITS OWN target reached (the monthly target from `config/walkin_target.yaml` for a month; the derived weekly target for a week): **Red** below 60 %, **Amber** 60 – 80 % (both edges inclusive), **Green** above 80 %. At a target of 100 a month is Red at 59 leads or fewer, Amber at 60 – 80, Green at 81 or more; a 22.6-lead week is Red at 13 or fewer, Amber at 14 – 18, Green at 19 or more (the tab's explanation line states these whole-lead edges for the current period). The table rows carry these colours both as fills and as real conditional-formatting rules (leads ÷ target of the row), and the Performance Status column names the band (plus "target exceeded" above 100 %). The chart is one clean column per week / month in its band colour with the number in the bar, the target as a bold dashed line, and the current period as a lighter bar with a dashed outline labelled "(to date)". Thresholds: `RAG_RED_BELOW` / `RAG_GREEN_ABOVE` in `common/walkin_targets.py`. Unchanged: the e-mail section (green above / red at-or-below target), the "completed periods above target" card, and every other tab.

**Weekly target formula** (no fixed weekly number):

    weekly target = Σ over Mon..Sun of  monthly target ÷ days in that day's month

So with a monthly target of 100, a week inside a 31-day month = 7 × 100 / 31 = **22.6**, inside a 30-day month = **23.3**, inside February = **25.0**; a week crossing a month end takes each day from its own month. Over a whole month the daily shares add back up to exactly the monthly target, so weekly and monthly targets agree.

- Target setting: `config/walkin_target.yaml` — see *Walk-In lead target file* below. Number of periods: `TREND_WEEKS_EMAIL / TREND_WEEKS_TAB / TREND_MONTHS_EMAIL / TREND_MONTHS_TAB` (5 / 11) in the script's USER SETTINGS.
- Counting reuses the report's own normalised, de-duplicated Walk-In records (`walkin_data.load_walkins`); one walk-in counts once, in the week / month of its date; undated rows are not counted (as everywhere in the report).
- Reference date: the report's own period end — today, `WEEKLY_REFERENCE_DATE`, or `MONTHLY_MONTH` / `MONTHLY_YEAR` (a past month is shown complete).
- One calculation (`common/walkin_targets.py`) feeds both the e-mail and the tab, so they always reconcile (the e-mail = the tab's last 6 rows).
- Weeks always run Monday – Sunday (`report.week_starts_on` should stay `monday` so the Summary's current week is the same week).
- Verification (offline, synthetic data): `python seo_validation\verify_walkin_lead_trend.py`.

### Walk-In dates — how a Timestamp is read (`common/walkin_data.py`, DATE RULE)

1. **Real date cells** (every Google Form *Timestamp*) are read by their **value** — the tabs are read with `dateTimeRenderOption=SERIAL_NUMBER` — so the spreadsheet's locale or display format can never swap day and month.
2. **Hand-typed text dates**: ISO (`2026-05-14`) and month names (`14-May-2026`, `3rd May 2026`) are unambiguous; an all-numeric `A/B/YEAR` (or `-` / `.`) is unambiguous when A > 12 (A = day) or B > 12 (B = day).
3. **Ambiguous text dates** (both parts ≤ 12, e.g. `05/06/2026`) are resolved in this order and each decision is logged: (i) the **tab's own evidence** — the order used by its unambiguous numeric text dates with the same separator, if ≥ 5 of them and ≥ 90 % agree; (ii) **row sequence** — the reading that falls between the dates of the rows above and below (± 7 days), if exactly one does; (iii) **default by separator** — `/` month/day (Google Sheets' US display), `-` and `.` day/month — flagged *DEFAULT*. Settings: `AMBIG_MIN_EVIDENCE`, `AMBIG_MIN_SHARE`, `AMBIG_NEIGHBOUR_DAYS`.

The same records feed every Weekly / Monthly figure, e-mail section, trend table and chart. Each run logs per tab how many Timestamps were date values, text, ambiguous (and how each was resolved) or blank. Tests: `python seo_validation\verify_walkin_dates.py`.

**Trace one month** (read-only): `python seo_validation\trace_month_walkins.py --month 2026-05` → `output\_trace - Walk-Ins May-2026.xlsx`:
*Summary* (spreadsheet locale / time zone; per tab the Timestamp cell types — date value / text / blank — and formats; the month's count by the report, by the previous display-text logic, by date-value cells only, as read DD/MM from the displayed text, and by any other date-like column), *Counted* (each lead: name, mobile, visit date, what the sheet shows, what the cell holds, the date rule used, lead source, technology, lead type, tab, sheet row, flags), *Month by month* (corrected vs previous counts per tab — historical validation), *Text dates* (every text Timestamp and how it was read), *Other date columns* (rows whose other date column is in a different month).

### Walk-In lead target file — `config/walkin_target.yaml`

The single place the target is set. Edit the number, save — no Python changes:

```yaml
MONTHLY_WALKIN_LEAD_TARGET: 100
```

Weekly target, e-mail section, Lead Trend tab, chart target line, Achievement % and the
green / red status all follow it on the next run. Check the file without running the
reports: `python seo_reports\pySEOWalkInAnalysisReport.py --check-target`.

Validation (every run, before anything is read from Google):

| Rule | Accepted | Rejected |
|---|---|---|
| File present and valid YAML (`NAME: value`) | UTF-8, Notepad BOM OK | missing / empty / unreadable file |
| `MONTHLY_WALKIN_LEAD_TARGET` present exactly once (case-insensitive name) | | missing, duplicated line |
| A number | `100`, `100.0`, `"100"` | blank, `null`, `true`, `abc`, `1,000`, `100%`, lists, NaN |
| Whole number | `100.0` → 100 | `100.5` |
| Range 1 – 10000 | | `0`, `-5`, `10001` |

Any other name in the file is ignored with a WARNING in the log (usually a typo).
If the value is invalid, the Weekly / Monthly reports are **not** generated (no workbook,
Drive upload or e-mail — a report with a wrong target is worse than none); the log says
exactly what is wrong and the run exits with code 1. A Manual report (no target) still runs.

## Layout

```
IntelliBI_SEO_Automation/
  common/            paths, logging, config, Google auth, normalization, data, builder
  config/            config.yaml (settings) · walkin_target.yaml (Walk-In lead target) ·
                     normalization.json (category rules)
  credentials/       service_account.json  (NOT committed — copy from Operations project)
  seo_reports/       pySEOWalkInAnalysisReport.py  (report generator)
  scripts/           run_seo_reports.py (entry point) · setup_schedule.ps1 (Task Scheduler)
  output/            generated .xlsx (git-ignored)
  logs/              run logs (git-ignored)
```

## Setup (on the machine that runs the schedule)

1. **Python env**

   ```
   cd IntelliBI_SEO_Automation
   python -m venv .venv
   .\.venv\Scripts\pip install -r requirements.txt
   ```

2. **Credentials** — copy the IntelliBI service account key into `credentials\service_account.json`
   (the same key `IntelliBI_Operations_Automation` uses).

3. **Share access with the service-account email** (`client_email` in the key —
   `intellibi-data-pipeline@intellibi-mis.iam.gserviceaccount.com`):
   - the source Google Sheet → **Viewer**
   - the Drive output folder (`drive.output_folder_id` in `config.yaml`) → **Editor**

4. **Run once**

   ```
   .\.venv\Scripts\python scripts\run_seo_reports.py           # Weekly + Monthly (today)
   .\.venv\Scripts\python scripts\run_seo_reports.py --no-upload
   ```

5. **Schedule daily** (elevated PowerShell):

   ```
   powershell -ExecutionPolicy Bypass -File scripts\setup_schedule.ps1
   ```

## Manual report

```
python seo_reports\pySEOWalkInAnalysisReport.py --mode manual --start 2026-08-01 --end 2026-08-15
```

## Email

Each report can be emailed (workbook attached + KPI summary + Drive link), sent
from `info@intellibiinnovationstechnologies.in` via Gmail SMTP using the SAME
`credentials/email_config.py` as the other IntelliBI projects (copy that file in;
the app password is read from it and never stored in this project).

- Recipients and sender: `email:` block in `config.yaml`
  (default recipients: `intellibiseo@gmail.com`, `info@intellibiinnovationstechnologies.in`).
- Turn it on/off: `EMAIL_SEND` in the script's USER SETTINGS block.
- Override per run: `--email` (force send) or `--no-email` (never send).
- **Gmail Starred (★)** — Weekly, Monthly and Manual report e-mails are starred in the
  **Info mailbox only** (`STAR_MAILBOX`), through `common/gmail_star.py` (the same file as
  in the Sales / Operations projects). After Gmail has accepted the e-mail, the script logs
  in to Info over IMAP (same app password, nothing new to configure), finds that exact
  e-mail by its Message-ID and stars it. Other recipients' copies are never touched; the
  subject, body, recipients and attachment are unchanged. Best-effort: a starring problem
  (IMAP disabled, network) is only logged and never affects the send, the report or the
  exit code. On/off: `STAR_EMAIL_IN_GMAIL` in USER SETTINGS. Requires IMAP enabled for
  Info (Gmail ▸ Settings ▸ Forwarding and POP/IMAP). Check recent reports (read-only):
  `python common\gmail_star.py --check`. Test (offline): `python seo_validation\verify_gmail_star.py`.

## Drive folders (per report type)

Uploads go into a **sub-folder per report type** under one parent folder
(`drive.parent_folder_id` in `config.yaml`). The sub-folder is resolved by NAME
(`drive.subfolders`, default `weekly` / `monthly` / `Manual`) at run time and
created automatically if missing — so each run lands in its own folder.

Share **only the parent folder** with the service-account email as **Editor**
(the sub-folders inherit access). Current parent: `1nYZNDromsZEWBM6-JszACJKrfquxQPoV`.

To pin a specific folder instead of resolving by name, put its ID in
`drive.folders.<type>` (an explicit ID always wins). Set `drive.upload: false` to
skip uploading entirely.

## What to run — USER SETTINGS (top of the script)

You don't pass parameters. Open `seo_reports/pySEOWalkInAnalysisReport.py` and edit
the **USER SETTINGS** block at the top, then just run the file:

```
GENERATE_WEEKLY  = True
GENERATE_MONTHLY = True
GENERATE_MANUAL  = False          # Manual = a custom start/end range

WEEKLY_REFERENCE_DATE = None      # "YYYY-MM-DD" — any day in the wanted week
MONTHLY_MONTH         = None      # 1-12  (None -> current month, to date)
MONTHLY_YEAR          = None      # e.g. 2026
MANUAL_START_DATE     = None      # "YYYY-MM-DD"  (required when GENERATE_MANUAL)
MANUAL_END_DATE       = None      # "YYYY-MM-DD"

EMAIL_SEND      = True            # email the report(s)?
UPLOAD_TO_DRIVE = True            # upload the report(s) to Drive?
```

Command-line flags still work as one-off overrides
(`--mode weekly|monthly|manual`, `--start/--end`, `--as-of`, `--no-upload`,
`--email/--no-email`) but are optional.

## Schedule

`setup_schedule.ps1` registers the daily task at **19:00 (7:00 PM)** by default
(`-Time "HH:mm"` to change). Weekly + Monthly are generated every day.

## Changing behaviour

- **Field mapping / source / Drive folders / email / week start / timezone** → `config/config.yaml`
- **Walk-In lead target (Weekly / Monthly)** → `config/walkin_target.yaml`
- **Category normalization** (Lead Source, Technology, Lead Type; Google Search grouping) →
  `config/normalization.json` — top-down keyword rules, edit here only.
