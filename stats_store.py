from __future__ import annotations

import json
import os
import platform
import time
import uuid
from typing import Any

from stats_models import StatsEntry


def _app_data_dir() -> str:
    """%LOCALAPPDATA%\\MirlisMark (same as main.py app_data_dir)."""
    if os.name == "nt":
        root = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
    else:
        root = os.path.expanduser("~")
    path = os.path.join(root, "MirlisMark")
    os.makedirs(path, exist_ok=True)
    return path


def _settings_path() -> str:
    return os.path.join(_app_data_dir(), "settings.json")


def _store_path() -> str:
    return os.path.join(_app_data_dir(), "stats_journal.jsonl")


def ensure_stats_store() -> str:
    """Ensure stats_journal.jsonl exists; return its path."""
    path = _store_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    if not os.path.exists(path):
        # Create empty journal file.
        # On Windows, write a UTF-8 BOM once so common viewers don't mis-detect it as ANSI/cp1251.
        # JSON parsers should use utf-8-sig when reading.
        enc = "utf-8-sig" if os.name == "nt" else "utf-8"
        with open(path, "w", encoding=enc):
            pass
    return path


def _load_settings() -> dict:
    try:
        p = _settings_path()
        if os.path.isfile(p):
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f) or {}
    except Exception:
        return {}
    return {}


def _save_settings(data: dict) -> None:
    try:
        with open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(data or {}, f, ensure_ascii=False, indent=2)
    except Exception:
        # Never raise from stats storage layer.
        return


def ensure_station_identity() -> tuple[str, str]:
    """
    Ensure station_uuid and station_label exist in settings.json.
    - station_uuid: generated once and persisted.
    - station_label: best-effort; fallback to computer name.
    """
    settings = _load_settings()

    station_uuid = str(settings.get("station_uuid") or "").strip()
    if not station_uuid:
        station_uuid = uuid.uuid4().hex
        settings["station_uuid"] = station_uuid

    station_label = str(settings.get("station_label") or "").strip()
    if not station_label:
        station_label = (
            os.environ.get("COMPUTERNAME")
            or platform.node()
            or "Station"
        )
        settings["station_label"] = station_label

    _save_settings(settings)
    return station_uuid, station_label


def append_entry(
    *,
    product: str = "",
    qty: str = "",
    unit: str = "",
    made_by: str = "",
    workshop: str = "",
    batch: str = "",
    copies: int = 1,
    ts: float | None = None,
    record_id: str | None = None,
) -> str:
    """
    Append one JSONL entry to %LOCALAPPDATA%\\MirlisMark\\stats_journal.jsonl.
    Returns record_id on success; returns "" on failure (never raises).
    """
    try:
        ensure_stats_store()
        station_uuid, station_label = ensure_station_identity()
        entry = StatsEntry.create(
            station_uuid=station_uuid,
            station_label=station_label,
            product=product,
            qty=qty,
            unit=unit,
            made_by=made_by,
            workshop=workshop,
            batch=batch,
            copies=copies,
            ts=(time.time() if ts is None else ts),
            record_id=record_id,
        )
        line = json.dumps(entry.to_dict(), ensure_ascii=False)
        with open(_store_path(), "a", encoding="utf-8") as f:
            f.write(line + "\n")
        return entry.record_id
    except Exception as e:
        try:
            import sys

            sys.stderr.write(f"[MirlisMark] stats_journal append error: {e}\n")
        except Exception:
            pass
        return ""


def read_local_entries() -> list[dict[str, Any]]:
    """Read all local JSONL entries. Bad lines are skipped."""
    path = ensure_stats_store()
    out: list[dict[str, Any]] = []
    try:
        # utf-8-sig safely strips BOM if present; otherwise behaves like utf-8.
        with open(path, "r", encoding="utf-8-sig") as f:
            for line in f:
                s = (line or "").strip()
                if not s:
                    continue
                try:
                    obj = json.loads(s)
                    if isinstance(obj, dict):
                        out.append(obj)
                except Exception:
                    continue
    except Exception as e:
        try:
            import sys

            sys.stderr.write(f"[MirlisMark] stats_journal read error: {e}\n")
        except Exception:
            pass
    return out

