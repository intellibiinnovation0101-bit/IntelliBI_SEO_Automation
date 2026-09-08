"""
Centralized normalization (common/seo_normalize.py).

All Lead Source / Technology / Lead Type categorisation is DATA-DRIVEN from
config/normalization.json — no category is hard-coded here. Rules are evaluated
top-down; the first rule whose 'any' keyword is a case-insensitive substring of
the cleaned value wins, else the section 'default'; a blank value -> blank_label.

This keeps the New/Old wording differences (case, spacing, spelling, equivalent
category names, multi-select course cells) reconciled in ONE maintainable place.
"""
from __future__ import annotations
import json, re
import paths

_RULES = None


def _load_rules() -> dict:
    global _RULES
    if _RULES is None:
        path = paths.CONFIG_DIR / "normalization.json"
        with open(path, "r", encoding="utf-8") as fh:
            _RULES = json.load(fh)
    return _RULES


def clean(value) -> str:
    """Collapse whitespace and strip surrounding quotes/spaces (does not alter case)."""
    s = re.sub(r"\s+", " ", str(value if value is not None else "")).strip()
    return s.strip("\"'“”‘’").strip()


def _apply(section: str, value) -> str:
    rules = _load_rules()
    blank = rules.get("blank_label", "Unknown")
    sec = rules.get(section, {})
    raw = clean(value)
    if not raw:
        return blank
    text = raw.lower()
    # Technology: many cells list several courses — categorise the PRIMARY (first) one.
    if sec.get("primary_only"):
        for sep in sec.get("split_on", [","]):
            if sep in text:
                text = text.split(sep)[0].strip()
                break
    for rule in sec.get("rules", []):
        for kw in rule.get("any", []):
            if kw.lower() in text:
                return rule["value"]
    # 'passthrough': keep the original (cleaned) value unchanged instead of
    # collapsing it into a generic bucket — so distinct business categories are
    # never merged, renamed or shortened.
    if sec.get("passthrough"):
        return raw
    return sec.get("default", "Other")


def norm_lead_source(value) -> str:
    return _apply("lead_source", value)


def norm_technology(value) -> str:
    return _apply("technology", value)


def norm_lead_type(value) -> str:
    return _apply("lead_type", value)


def google_search_label() -> str:
    """The canonical label used for the SEO headline metric."""
    return "Google Search"


def is_google_search(normalized_source: str) -> bool:
    return clean(normalized_source).lower() == google_search_label().lower()
