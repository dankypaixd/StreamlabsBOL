"""Service level por estado y fecha de entrega.

Reglas (confirmadas por el usuario):
* El service level se escribe "3D", "4D", etc. y depende del estado del Ship To.
* La fecha de entrega se mide desde el READY DATE (DATA!N2, "MM/DD/YY HHMM").
* Cut-off 08:00 - 17:00. Dentro del cut-off, ese mismo dia es el dia 0.
  Despues de las 17:00 el dia 0 es el siguiente dia habil.
* Los fines de semana no se cuentan.
* La fecha de entrega se escribe MM/DD/YY.

Supuestos que el usuario no detallo (ajustables aqui):
* Antes de las 08:00 el dia 0 sigue siendo ese mismo dia.
* Si el ready date cae en sabado o domingo, el dia 0 es el lunes siguiente.
* No se descuentan feriados (HOLIDAYS esta vacio a proposito).
"""
from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from functools import lru_cache
from pathlib import Path

DEFAULT_TABLE = Path(__file__).parent / "assets" / "SERVICE_LEVELS.csv"

CUTOFF_START = time(8, 0)
CUTOFF_END = time(17, 0)
HOLIDAYS: frozenset[date] = frozenset()  # p. ej. frozenset({date(2026, 11, 26)})


@dataclass(frozen=True)
class StateInfo:
    code: str
    name: str
    days: int
    note: str = ""


@lru_cache(maxsize=4)
def load_table(path: str | Path = DEFAULT_TABLE) -> dict[str, StateInfo]:
    table: dict[str, StateInfo] = {}
    with open(path, newline="", encoding="utf-8") as handle:
        for row in csv.DictReader(handle):
            code = row["code"].strip().upper()
            table[code] = StateInfo(code, row["name"].strip(), int(row["days"]), row.get("note", "").strip())
    return table


def normalize_state(value: str | None) -> str:
    """Devuelve el codigo de 2 letras (TX1/WI1 -> TX/WI) o '' si no existe."""
    text = re.sub(r"[^A-Za-z0-9]", "", str(value or "")).upper()
    if text[:2] in load_table() and (len(text) == 2 or text[2:].isdigit()):
        return text[:2]
    return ""


# Primeros 3 digitos del ZIP -> estado (rangos del USPS). Sirve cuando el Ship To no trae el
# codigo de estado escrito, o para cruzar el estado leido con el codigo postal.
_ZIP3_RANGES: tuple[tuple[int, int, str], ...] = (
    (5, 5, "NY"), (6, 7, "PR"), (9, 9, "PR"), (10, 27, "MA"), (28, 29, "RI"), (30, 38, "NH"),
    (39, 49, "ME"), (50, 59, "VT"), (60, 69, "CT"), (70, 89, "NJ"), (100, 149, "NY"),
    (150, 196, "PA"), (197, 199, "DE"), (200, 200, "DC"), (201, 201, "VA"), (202, 205, "DC"),
    (206, 219, "MD"), (220, 246, "VA"), (247, 268, "WV"), (270, 289, "NC"), (290, 299, "SC"),
    (300, 319, "GA"), (320, 349, "FL"), (350, 369, "AL"), (370, 385, "TN"), (386, 397, "MS"),
    (398, 399, "GA"), (400, 427, "KY"), (430, 459, "OH"), (460, 479, "IN"), (480, 499, "MI"),
    (500, 528, "IA"), (530, 549, "WI"), (550, 567, "MN"), (570, 577, "SD"), (580, 588, "ND"),
    (590, 599, "MT"), (600, 629, "IL"), (630, 658, "MO"), (660, 679, "KS"), (680, 693, "NE"),
    (700, 714, "LA"), (716, 729, "AR"), (730, 749, "OK"), (750, 799, "TX"), (800, 816, "CO"),
    (820, 831, "WY"), (832, 838, "ID"), (840, 847, "UT"), (850, 865, "AZ"), (870, 884, "NM"),
    (885, 885, "TX"), (889, 898, "NV"), (900, 961, "CA"), (967, 968, "HI"), (970, 979, "OR"),
    (980, 994, "WA"), (995, 999, "AK"),
)


def state_from_zip(zip_code: str | None) -> str:
    """Estado segun el ZIP de 5 digitos ('' si no se puede o no esta en la tabla de service level)."""
    digits = re.sub(r"\D", "", str(zip_code or ""))
    if len(digits) < 5:
        return ""
    prefix = int(digits[:3])
    for low, high, code in _ZIP3_RANGES:
        if low <= prefix <= high:
            return code if code in load_table() else ""
    return ""


def state_from_name(text: str | None) -> str:
    """Busca un nombre completo de estado o provincia ('Puerto Rico', 'Florida'...) en un texto."""
    haystack = " " + re.sub(r"[^A-Za-z ]", " ", str(text or "")).lower() + " "
    haystack = re.sub(r"\s+", " ", haystack)
    best = ""
    best_len = 0
    for info in load_table().values():
        name = re.sub(r"\s+", " ", info.name.lower()).strip()
        if len(name) > best_len and f" {name} " in haystack:
            best, best_len = info.code, len(name)
    return best


def service_days(state: str | None) -> int | None:
    info = load_table().get(normalize_state(state))
    return info.days if info else None


def service_label(state: str | None) -> str:
    days = service_days(state)
    return f"{days}D" if days is not None else ""


# ----------------------------------------------------------------- fechas


_READY_RE = re.compile(
    r"^\s*(\d{1,2})/(\d{1,2})/(\d{2}|\d{4})(?:\s+(\d{1,2}):?(\d{2}))?\s*$"
)


def parse_ready(text: str | datetime | None) -> datetime | None:
    """Acepta '11/19/26 1356', '11/19/26 13:56', '11/19/2026 0800' o '11/19/26'."""
    if isinstance(text, datetime):
        return text
    match = _READY_RE.match(str(text or ""))
    if not match:
        return None
    month, day, year, hour, minute = match.groups()
    year_i = int(year) + 2000 if len(year) == 2 else int(year)
    try:
        return datetime(year_i, int(month), int(day), int(hour or 8), int(minute or 0))
    except ValueError:
        return None


def is_business_day(day: date, holidays: frozenset[date] = HOLIDAYS) -> bool:
    return day.weekday() < 5 and day not in holidays


def next_business_day(day: date, holidays: frozenset[date] = HOLIDAYS) -> date:
    while not is_business_day(day, holidays):
        day += timedelta(days=1)
    return day


def day_zero(ready: datetime, holidays: frozenset[date] = HOLIDAYS) -> date:
    day = ready.date()
    if ready.time() > CUTOFF_END:
        day += timedelta(days=1)
    return next_business_day(day, holidays)


def add_business_days(start: date, days: int, holidays: frozenset[date] = HOLIDAYS) -> date:
    current = start
    remaining = days
    while remaining > 0:
        current += timedelta(days=1)
        if is_business_day(current, holidays):
            remaining -= 1
    return current


def delivery_date(ready: datetime, days: int, holidays: frozenset[date] = HOLIDAYS) -> date:
    return add_business_days(day_zero(ready, holidays), days, holidays)


def format_mmddyy(day: date) -> str:
    return day.strftime("%m/%d/%y")


def delivery_text(ready: datetime | str | None, state: str | None) -> str:
    """Fecha de entrega 'MM/DD/YY' o '' si falta el ready date o el estado."""
    parsed = parse_ready(ready)
    days = service_days(state)
    if parsed is None or days is None:
        return ""
    return format_mmddyy(delivery_date(parsed, days))


def explain(ready: datetime | str | None, state: str | None) -> str:
    parsed = parse_ready(ready)
    days = service_days(state)
    if parsed is None or days is None:
        return ""
    zero = day_zero(parsed)
    when = "dentro del cut-off" if CUTOFF_START <= parsed.time() <= CUTOFF_END else (
        "despues del cut-off" if parsed.time() > CUTOFF_END else "antes del cut-off"
    )
    return (
        f"Ready {parsed:%m/%d/%y %H:%M} ({when}) -> dia 0 = {format_mmddyy(zero)}, "
        f"{days}D -> {format_mmddyy(delivery_date(parsed, days))}"
    )
