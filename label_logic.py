from __future__ import annotations
from dataclasses import dataclass
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

MSK = ZoneInfo("Europe/Moscow")

WEEKDAYS_RU = {
    0: "ПОНЕДЕЛЬНИК",
    1: "ВТОРНИК",
    2: "СРЕДА",
    3: "ЧЕТВЕРГ",
    4: "ПЯТНИЦА",
    5: "СУББОТА",
    6: "ВОСКРЕСЕНЬЕ",
}

@dataclass
class LabelData:
    weekday: str
    product_name: str
    qty_value: str
    qty_unit_ru: str  # "кг" или "шт"
    produced_at: datetime
    batch: str
    expires_at: datetime
    made_by: str
    checked_by: str


def now_msk() -> datetime:
    return datetime.now(tz=MSK)


def format_dt(dt: datetime) -> str:
    return dt.strftime("%d.%m.%Y %H:%M")


def batch_from_dt(dt: datetime) -> str:
    # DDMMYY без точек
    return dt.strftime("%d%m%y")


def unit_ru(unit: str) -> str:
    return "кг" if unit == "kg" else "шт"


def build_label(
    product_name: str,
    shelf_life_hours: int,
    qty_value: str,
    unit: str,  # "kg" / "pcs"
    made_by: str = "",
    checked_by: str = "",
    produced_at: datetime | None = None,
) -> LabelData:
    dt = produced_at or now_msk()
    expires = dt + timedelta(hours=int(shelf_life_hours))
    weekday = WEEKDAYS_RU[dt.weekday()]
    return LabelData(
        weekday=weekday,
        product_name=product_name,
        qty_value=qty_value,
        qty_unit_ru=unit_ru(unit),
        produced_at=dt,
        batch=batch_from_dt(dt),
        expires_at=expires,
        made_by=made_by.strip(),
        checked_by=checked_by.strip(),
    )
def generate_tspl(label):
    tspl = f"""
SIZE 80 mm, 60 mm
GAP 2 mm, 0 mm
CLS
TEXT 20,20,"0",0,2,2,"{label.weekday}"
TEXT 20,70,"0",0,1,1,"Продукт: {label.product_name}"
TEXT 20,100,"0",0,1,1,"Вес/шт: {label.qty_value} {label.qty_unit_ru}"
TEXT 20,130,"0",0,1,1,"Дата/время: {format_dt(label.produced_at)}"
TEXT 20,160,"0",0,1,1,"№ партии: {label.batch}"
TEXT 20,190,"0",0,1,1,"Годен до: {format_dt(label.expires_at)}"
TEXT 20,220,"0",0,1,1,"Изготовил: {label.made_by}"
TEXT 20,250,"0",0,1,1,"Проверил: {label.checked_by}"
PRINT 1
"""
    return tspl.strip()