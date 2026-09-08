"""Configuration loader (common/config_loader.py) — reads config/config.yaml.

get('a.b.c', default) reads a dotted key. PyYAML is used when available; if it is
not installed, a small built-in parser handles this project's config (nested maps,
scalar values, and simple string lists), so the tool never silently loses its
settings just because PyYAML is missing from the environment.
"""
from __future__ import annotations
import paths

CONFIG_YAML = paths.CONFIG_DIR / "config.yaml"
_cache = None


# ── minimal YAML subset parser (fallback when PyYAML is not installed) ────────
def _scalar(v):
    v = v.strip()
    if len(v) >= 2 and v[0] in "\"'" and v[-1] == v[0]:
        return v[1:-1]
    low = v.lower()
    if low in ("true", "yes", "on"):
        return True
    if low in ("false", "no", "off"):
        return False
    if low in ("null", "~"):
        return None
    if v and (v.lstrip("-").isdigit()):
        try:
            return int(v)
        except ValueError:
            pass
    return v


def _strip_comment(v):
    v = v.strip()
    if v and v[0] in "\"'":
        q = v[0]
        end = v.find(q, 1)
        return v[:end + 1] if end != -1 else v
    if "#" in v:
        return v.split("#", 1)[0].strip()
    return v


def _mini_yaml(text: str) -> dict:
    lines = []
    for raw in text.split("\n"):
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        indent = len(raw) - len(raw.lstrip(" "))
        lines.append((indent, raw.strip()))

    pos = [0]

    def parse_block(min_indent):
        result = None
        while pos[0] < len(lines):
            indent, content = lines[pos[0]]
            if indent < min_indent:
                break
            if content.startswith("- "):
                if result is None:
                    result = []
                result.append(_scalar(_strip_comment(content[2:])))
                pos[0] += 1
            else:
                if result is None:
                    result = {}
                key, _, val = content.partition(":")
                key = key.strip()
                val = _strip_comment(val)
                pos[0] += 1
                if val == "":
                    if pos[0] < len(lines) and lines[pos[0]][0] > indent:
                        result[key] = parse_block(indent + 1)
                    else:
                        result[key] = {}
                else:
                    result[key] = _scalar(val)
        return result if result is not None else {}

    return parse_block(0) or {}


def load() -> dict:
    global _cache
    if _cache is not None:
        return _cache
    data = {}
    if CONFIG_YAML.exists():
        text = None
        try:
            with open(CONFIG_YAML, "r", encoding="utf-8") as fh:
                text = fh.read()
        except Exception as e:
            print(f"[config] WARNING: could not read config.yaml ({e}); using defaults")
        if text is not None:
            try:
                import yaml
                data = yaml.safe_load(text) or {}
            except ImportError:
                # PyYAML not installed — use the built-in parser for this config.
                data = _mini_yaml(text) or {}
            except Exception as e:
                print(f"[config] WARNING: could not parse config.yaml ({e}); using defaults")
                data = {}
    _cache = data
    return data


def get(dotted: str, default=None):
    node = load()
    for part in dotted.split("."):
        if isinstance(node, dict) and part in node:
            node = node[part]
        else:
            return default
    return node
