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
| Target | see formula below | 70 Walk-In leads per month |

Each row: period, Walk-In leads, target, Achievement % (= leads ÷ target), status. **Green = leads above target, red = at or below target.** The current week / month is counted up to the report date (the same figure as the Summary's *Current* total), marked *In progress* (e-mail chip, "▶ current" row, `*` on the chart axis, lighter bar) with its days elapsed and its pro-rata target to date; its colour still compares against the full target. The tab adds three summary cards, a completed-periods total row and a chart: bars = actual leads (green / red, current lighter, value labels), dashed line = target.

**Weekly target formula** (no fixed weekly number):

    weekly target = Σ over Mon..Sun of  monthly target ÷ days in that day's month

So a week inside a 31-day month = 7 × 70 / 31 = **15.8**, inside a 30-day month = **16.3**, inside February = **17.5**; a week crossing a month end takes each day from its own month. Over a whole month the daily shares add back up to exactly 70, so weekly and monthly targets agree.

- Target setting: `report.monthly_walkin_target` in `config/config.yaml` (70). Number of periods: `TREND_WEEKS_EMAIL / TREND_WEEKS_TAB / TREND_MONTHS_EMAIL / TREND_MONTHS_TAB` (5 / 11) in the script's USER SETTINGS.
- Counting reuses the report's own normalised, de-duplicated Walk-In records (`walkin_data.load_walkins`); one walk-in counts once, in the week / month of its date; undated rows are not counted (as everywhere in the report).
- Reference date: the report's own period end — today, `WEEKLY_REFERENCE_DATE`, or `MONTHLY_MONTH` / `MONTHLY_YEAR` (a past month is shown complete).
- One calculation (`common/walkin_targets.py`) feeds both the e-mail and the tab, so they always reconcile (the e-mail = the tab's last 6 rows).
- Weeks always run Monday – Sunday (`report.week_starts_on` should stay `monday` so the Summary's current week is the same week).
- Verification (offline, synthetic data): `python seo_validation\verify_walkin_lead_trend.py`.

## Layout

```
IntelliBI_SEO_Automation/
  common/            paths, logging, config, Google auth, normalization, data, builder
  config/            config.yaml (settings) · normalization.json (category rules)
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
- **Category normalization** (Lead Source, Technology, Lead Type; Google Search grouping) →
  `config/normalization.json` — top-down keyword rules, edit here only.
