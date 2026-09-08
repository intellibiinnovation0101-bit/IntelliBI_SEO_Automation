"""
IntelliBI SEO Automation — central path resolver (common/paths.py).

Every path derives from PROJECT_ROOT, discovered at run time from this file's
location (common/paths.py -> parents[1]). Nothing machine-specific: copy the
whole IntelliBI_SEO_Automation folder anywhere and paths still resolve.
"""
from __future__ import annotations
from pathlib import Path

PROJECT_ROOT: Path = Path(__file__).resolve().parents[1]
COMMON_DIR      = PROJECT_ROOT / "common"
CONFIG_DIR      = PROJECT_ROOT / "config"
CREDENTIALS_DIR = PROJECT_ROOT / "credentials"
REPORTS_DIR     = PROJECT_ROOT / "seo_reports"
SCRIPTS_DIR     = PROJECT_ROOT / "scripts"
LOGS_DIR        = PROJECT_ROOT / "logs"
OUTPUT_DIR      = PROJECT_ROOT / "output"

_ENSURE = (LOGS_DIR, OUTPUT_DIR)


def ensure_dirs() -> None:
    for d in _ENSURE:
        d.mkdir(parents=True, exist_ok=True)
