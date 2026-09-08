"""Centralized logging (common/logging_utils.py) — console + logs/<name>.log."""
from __future__ import annotations
import logging, os
from pathlib import Path
import paths

_LEVEL = os.environ.get("SEO_LOG_LEVEL", "INFO").upper()
_FMT = "%(asctime)s | %(levelname)-7s | %(name)s | %(message)s"
_DATEFMT = "%Y-%m-%d %H:%M:%S"
_configured: dict[str, logging.Logger] = {}


def log_file_for(name: str) -> Path:
    paths.LOGS_DIR.mkdir(parents=True, exist_ok=True)
    safe = "".join(c if (c.isalnum() or c in "._-") else "_" for c in name)
    return paths.LOGS_DIR / f"{safe}.log"


def get_logger(name: str) -> logging.Logger:
    if name in _configured:
        return _configured[name]
    lg = logging.getLogger(name)
    lg.setLevel(_LEVEL)
    lg.propagate = False
    fmt = logging.Formatter(_FMT, _DATEFMT)
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    lg.addHandler(ch)
    try:
        fh = logging.FileHandler(log_file_for(name), encoding="utf-8")
        fh.setFormatter(fmt)
        lg.addHandler(fh)
    except Exception:
        pass
    _configured[name] = lg
    return lg


def section(logger: logging.Logger, title: str) -> None:
    logger.info("-" * 64)
    logger.info(title)
    logger.info("-" * 64)
