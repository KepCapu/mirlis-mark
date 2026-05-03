from __future__ import annotations

from dataclasses import dataclass
import time
import uuid


@dataclass(frozen=True)
class StatsEntry:
    record_id: str
    station_uuid: str
    station_label: str
    ts: float
    product: str
    qty: str
    unit: str
    made_by: str
    workshop: str
    batch: str
    copies: int

    @staticmethod
    def new_id() -> str:
        return uuid.uuid4().hex

    @classmethod
    def create(
        cls,
        *,
        station_uuid: str,
        station_label: str,
        product: str = "",
        qty: str = "",
        unit: str = "",
        made_by: str = "",
        workshop: str = "",
        batch: str = "",
        copies: int = 1,
        ts: float | None = None,
        record_id: str | None = None,
    ) -> "StatsEntry":
        rid = (record_id or "").strip() or cls.new_id()
        su = (station_uuid or "").strip()
        sl = (station_label or "").strip()
        return cls(
            record_id=rid,
            station_uuid=su,
            station_label=sl,
            ts=float(time.time() if ts is None else ts),
            product=str(product or ""),
            qty=str(qty or ""),
            unit=str(unit or ""),
            made_by=str(made_by or ""),
            workshop=str(workshop or ""),
            batch=str(batch or ""),
            copies=int(copies) if copies is not None else 1,
        )

    def to_dict(self) -> dict:
        return {
            "record_id": self.record_id,
            "station_uuid": self.station_uuid,
            "station_label": self.station_label,
            "ts": float(self.ts),
            "product": self.product,
            "qty": self.qty,
            "unit": self.unit,
            "made_by": self.made_by,
            "workshop": self.workshop,
            "batch": self.batch,
            "copies": int(self.copies),
        }

