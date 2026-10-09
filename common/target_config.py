"""
Walk-In lead target configuration (common/target_config.py).

Reads the monthly Walk-In lead target from config/walkin_target.yaml — the single
place it is set — and validates it. Parsing reuses the project's config approach
(common/config_loader: PyYAML, or the built-in parser when PyYAML is missing).

FILE FORMAT (YAML, one line that matters):
    MONTHLY_WALKIN_LEAD_TARGET: 100

VALIDATION RULES (any failure raises TargetConfigError with a plain message):
  1. The file exists, is readable text (UTF-8, a Notepad BOM is fine) and parses
     as YAML with "name: value" lines at the top level.
  2. MONTHLY_WALKIN_LEAD_TARGET is present exactly once (the name is matched
     ignoring upper / lower case; a duplicate line is an error, because YAML
     would otherwise silently keep only the last one).
  3. The value is a number: 100, 100.0 or "100" are accepted; empty / null,
     true / false, text ("abc", "1,000", "100%"), lists and NaN / infinity are not.
  4. It is a whole number (100.0 -> 100; 100.5 is rejected).
  5. It is between MIN_TARGET (1) and MAX_TARGET (10000) inclusive — 0 or a
     negative target would make every period "Above Target"; a huge value is
     almost certainly a typo.
  Other names in the file are ignored with a WARNING (usually a misspelt key).
"""
from __future__ import annotations
import math
import re
from pathlib import Path

import paths
import config_loader

TARGET_FILE = paths.CONFIG_DIR / "walkin_target.yaml"
TARGET_KEY = "MONTHLY_WALKIN_LEAD_TARGET"
MIN_TARGET = 1
MAX_TARGET = 10000

_EXAMPLE = f"{TARGET_KEY}: 100"


class TargetConfigError(ValueError):
    """The Walk-In target file is missing or holds an invalid value."""


def validate_target(value, source: str = TARGET_KEY) -> int:
    """Rules 3-5: return the target as an int, or raise TargetConfigError."""
    shown = repr(value)
    if value is None or (isinstance(value, (str, dict, list)) and not (value.strip() if isinstance(value, str) else value)):
        raise TargetConfigError(f"{source} has no value. Write a whole number, e.g. '{_EXAMPLE}'.")
    if isinstance(value, bool):
        raise TargetConfigError(f"{source} = {shown} is not a number. Write a whole number, e.g. '{_EXAMPLE}'.")
    if isinstance(value, (int, float)):
        num = float(value)
    elif isinstance(value, str):
        try:
            num = float(value.strip())
        except ValueError:
            raise TargetConfigError(
                f"{source} = {shown} is not a number. Write digits only (no commas, % or words), "
                f"e.g. '{_EXAMPLE}'.") from None
    else:
        raise TargetConfigError(
            f"{source} = {shown} must be a single number, not a {type(value).__name__}. "
            f"e.g. '{_EXAMPLE}'.")
    if math.isnan(num) or math.isinf(num):
        raise TargetConfigError(f"{source} = {shown} is not a usable number. e.g. '{_EXAMPLE}'.")
    if num != int(num):
        raise TargetConfigError(f"{source} = {shown} must be a whole number of leads (e.g. 100, not 100.5).")
    num = int(num)
    if not MIN_TARGET <= num <= MAX_TARGET:
        raise TargetConfigError(
            f"{source} = {shown} is out of range — it must be between {MIN_TARGET} and {MAX_TARGET}.")
    return num


def load_monthly_target(path=None, warn=None) -> int:
    """Read + validate the target file (rules 1-5). `warn(msg)` receives
    non-fatal warnings (unknown names); defaults to print."""
    path = Path(path) if path else TARGET_FILE
    warn = warn or (lambda m: print(f"[target] WARNING: {m}"))
    where = f"{path}"
    if not path.is_file():
        raise TargetConfigError(
            f"Target file not found: {where}. Create it with the single line '{_EXAMPLE}'.")
    try:
        data = config_loader.read_yaml_file(path)
        text = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError:
        raise TargetConfigError(f"{where} is not a readable text file — save it as UTF-8 (plain text).") from None
    except OSError as e:
        raise TargetConfigError(f"Cannot read {where}: {e}") from None
    except ValueError as e:
        raise TargetConfigError(f"{where} could not be read: {e}. Expected a line like '{_EXAMPLE}'.") from None
    if not isinstance(data, dict) or not data:
        raise TargetConfigError(
            f"{where} has no '{TARGET_KEY}: <number>' line. Expected a line like '{_EXAMPLE}'.")

    matches = [k for k in data if str(k).strip().upper() == TARGET_KEY]
    if not matches:
        raise TargetConfigError(
            f"{where} does not contain {TARGET_KEY}. Expected a line like '{_EXAMPLE}'"
            + (f" (found: {', '.join(map(str, data))})." if data else "."))
    lines = re.findall(rf"^\s*['\"]?{TARGET_KEY}['\"]?\s*:", text, re.M | re.I)
    if len(matches) > 1 or len(lines) > 1:
        raise TargetConfigError(
            f"{where} sets {TARGET_KEY} more than once — keep exactly one line.")
    for k in data:
        if k not in matches:
            warn(f"{where}: ignoring unknown setting '{k}' (only {TARGET_KEY} is read).")
    return validate_target(data[matches[0]], f"{TARGET_KEY} in {path.name}")
