"""
Runner (scripts/run_seo_reports.py) — the entry point the Windows scheduled task
calls. Generates the daily Weekly + Monthly reports and uploads them to Drive.

    python scripts/run_seo_reports.py            # weekly + monthly for today
    python scripts/run_seo_reports.py --no-upload
    python scripts/run_seo_reports.py --mode manual --start 2026-08-01 --end 2026-08-15
"""
from __future__ import annotations
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(_HERE)
sys.path.insert(0, os.path.join(_ROOT, "common"))
sys.path.insert(0, os.path.join(_ROOT, "seo_reports"))

import pySEOWalkInAnalysisReport as report   # noqa: E402


if __name__ == "__main__":
    sys.exit(report.main(sys.argv[1:]))
