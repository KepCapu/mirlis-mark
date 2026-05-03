from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Iterable


@dataclass(frozen=True)
class PrintRecord:
    dt: datetime
    product: str
    made_by: str
    workshop: str
    copies: int


def compute_shift_totals(records: Iterable[PrintRecord]) -> tuple[int, int]:
    """Return (day_shift_labels, night_shift_labels) for provided records.

    Day shift: 08:00–19:59
    Night shift: 20:00–07:59
    """
    day_shift_total = 0
    night_shift_total = 0
    for r in (records or []):
        try:
            hour = int(r.dt.hour)
        except Exception:
            continue
        try:
            copies = int(getattr(r, "copies", 0) or 0)
        except Exception:
            copies = 0
        if copies <= 0:
            continue
        if 8 <= hour < 20:
            day_shift_total += copies
        else:
            night_shift_total += copies
    return day_shift_total, night_shift_total


def normalize_stat_key(value: object) -> str | None:
    """Normalize category key for stats aggregation.

    Returns a cleaned string, or None if the value is empty/service placeholder
    and must not be aggregated as a real category.
    """
    if value is None:
        return None
    try:
        s = str(value).strip()
    except Exception:
        return None
    if not s:
        return None
    s_l = s.lower()
    if s_l in {"-", "—", "–", "―", "не указано", "неизвестно", "n/a"}:
        return None
    # Sometimes users accidentally save placeholders like "---" etc.
    if all(ch in "-—–―" for ch in s_l):
        return None
    return s


def _parse_int(s: str, default: int = 0) -> int:
    try:
        return int(str(s).strip())
    except Exception:
        return default


def _parse_dt(s: str) -> datetime | None:
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%d.%m.%Y %H:%M", "%d.%m.%Y %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    return None


def parse_print_record_from_text(text: str, fallback_dt: datetime) -> PrintRecord:
    product = ""
    made_by = ""
    workshop = ""
    copies = 0
    dt = None

    for raw in (text or "").splitlines():
        line = raw.strip()
        if not line or ":" not in line:
            continue
        k, v = line.split(":", 1)
        key = k.strip().lower()
        val = v.strip()

        if key == "продукт":
            product = val
        elif key in ("дата/время", "дата", "дата-время", "дата время"):
            dt = _parse_dt(val) or dt
        elif key == "изготовил":
            made_by = val
        elif key == "цех":
            workshop = val
        elif key == "количество":
            copies = _parse_int(val, copies or 0)

    if dt is None:
        dt = fallback_dt

    product = (product or "").strip() or "—"
    made_by = (made_by or "").strip() or "—"
    workshop = (workshop or "").strip() or "—"
    copies = int(copies) if int(copies) > 0 else 1

    return PrintRecord(dt=dt, product=product, made_by=made_by, workshop=workshop, copies=copies)


def load_print_records_from_archive(archive_root: str) -> list[PrintRecord]:
    """Загрузить записи печати из архива txt-файлов.

    Ожидаемая структура:
      <archive_root>/<dd.mm.yyyy>/*.txt
    """
    records: list[PrintRecord] = []
    if not archive_root or not os.path.isdir(archive_root):
        return records

    for day_name in sorted(os.listdir(archive_root)):
        day_dir = os.path.join(archive_root, day_name)
        if not os.path.isdir(day_dir):
            continue
        try:
            day_dt = datetime.strptime(day_name.strip(), "%d.%m.%Y")
        except Exception:
            day_dt = None

        for name in os.listdir(day_dir):
            if not name.lower().endswith(".txt"):
                continue
            fpath = os.path.join(day_dir, name)
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    text = f.read()
            except Exception:
                continue

            fallback_dt = day_dt or datetime.fromtimestamp(os.path.getmtime(fpath))
            try:
                rec = parse_print_record_from_text(text, fallback_dt=fallback_dt)
            except Exception:
                continue
            records.append(rec)

    records.sort(key=lambda r: r.dt)
    return records


def filter_records_by_period(records: Iterable[PrintRecord], period: str, now: datetime | None = None) -> list[PrintRecord]:
    now = now or datetime.now()
    period = (period or "day").strip().lower()
    rs = list(records or [])

    if period == "day":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=1)
    elif period == "week":
        end = now
        start = (end - timedelta(days=6)).replace(hour=0, minute=0, second=0, microsecond=0)
    elif period == "month":
        # Последние 30 календарных дней до текущего момента (rolling), не текущий календарный месяц.
        end = now
        d0 = now.date() - timedelta(days=29)
        start = datetime.combine(d0, datetime.min.time())
    else:
        return rs

    return [r for r in rs if start <= r.dt <= end]


def filter_records_by_datetime_range(
    records: Iterable[PrintRecord],
    start_dt: datetime,
    end_dt: datetime,
) -> list[PrintRecord]:
    """Фильтрация записей по произвольному временному интервалу [start_dt, end_dt] включительно."""
    rs = list(records or [])
    if start_dt > end_dt:
        return []
    return [r for r in rs if start_dt <= r.dt <= end_dt]

