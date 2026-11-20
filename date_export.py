"""Genera el Excel de importacion igual que el template (hojas DATE y DATA).

En el template, lo que se pega en DATA se refleja por formulas en DATE. Aqui se escriben
los valores ya calculados (sin formulas dinamicas) para que el archivo se importe tal cual.

Columnas de DATE que se llenan (el resto queda vacio, como en el template):
    A  Handling Station        "DFW" cuando hay Housebill
    B  Housebill               BOL (Barcode de DATA)
    E  Shipper Number          BEX -> 149570, VRP -> 144463
    Q  Pickup Date             primeros 8 caracteres del READY DATE (MM/DD/YY)
    R  Pickup Time             ultimos 5 caracteres del READY DATE (como RIGHT(N2,5))
    S  Pickup Close Time       "1700"
    AB Consignee Contact       nombre            <- Ship To Attention
    AC Consignee Phone         telefono          <- Ship To Attention
    AF Scheduled Delivery Date fecha de entrega (MM/DD/YY, segun service level y cut-off)
    AG Scheduled Delivery Time "1600"
    AK Service                 "3D", "4D"...     (segun el estado del Ship To)
    CO Product Code            "HPDOM"
    DJ BilltoRefType           "RAX"
    DK BilltoRef               "1" si algun renglon del BOL en DATA tiene HAZMAT = Y, si no "0"
    DU Consignee Email         correo            <- Ship To Attention
    FK Priority Code           "INBD" (solo pestana Inbonds)

El DATA del libro se arma con el DIMS que sube el usuario (cualquier orden de columnas, con o sin
TRACKER / HAZMAT / RACK) en el orden del template; se rellenan N2 (READY DATE), N3:N4 (POD DATE TIME),
Q (SERVICE LEVEL) y R (DELIVERY DATE) en cada renglon del BOL.
"""
from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass, field
from datetime import date, datetime, time as dtime
from pathlib import Path
from typing import Iterable, Mapping

from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from service_levels import delivery_text, parse_ready, service_label

HEADERS_FILE = Path(__file__).parent / "assets" / "DATE_HEADERS.json"

SHIPPER_NUMBER = {"BEX": "149570", "VRP": "144463"}
STATION = "DFW"
PICKUP_CLOSE = "1700"
DELIVERY_TIME = "1600"
PRODUCT_CODE = "HPDOM"
BILLTO_REF_TYPE = "RAX"

PRIORITY_INBOND = "INBD"

# Orden de las columnas de DATA en el template. TRACKER no existe en el template: va al final (S).
DATA_LAYOUT = [
    ("carrier", "A"), ("trailer", "B"), ("barcode", "C"), ("length", "D"), ("width", "E"), ("height", "F"),
    ("weight", "G"), ("box", "H"), ("rack", "I"), ("hazmat", "J"), ("scanned", "K"), ("time_in", "L"),
    ("time_out", "M"), ("ready", "N"), ("service", "Q"), ("delivery", "R"), ("tracker", "S"),
]
# Encabezados del DIMS (sin importar mayusculas, espacios ni signos) -> campo
DIMS_HEADERS = {
    "carrier": "carrier", "trailer": "trailer", "barcode": "barcode", "length": "length", "width": "width",
    "height": "height", "tracker": "tracker", "tracking": "tracker", "trackingno": "tracker",
    "actualweight": "weight", "weight": "weight", "box": "box", "boxes": "box", "rack": "rack",
    "hazmat": "hazmat", "scannedtime": "scanned", "timein": "time_in", "timeout": "time_out",
    "readydate": "ready",
}
YES_VALUES = {"Y", "YES", "SI", "SÍ", "X", "1", "TRUE", "HAZMAT"}


def load_headers() -> dict[str, list[str]]:
    return json.loads(HEADERS_FILE.read_text(encoding="utf-8"))


def format_ready(value: str | datetime | None) -> str:
    """Normaliza el READY DATE a 'MM/DD/YY HHMM' (como esta en DATA!N2)."""
    parsed = parse_ready(value)
    if parsed is None:
        return ""
    return f"{parsed:%m/%d/%y} {parsed:%H%M}"


BLANK = "-"  # en una fila: "dejar este campo vacio" (distinto de "usar el calculado")


def normalize_delivery_date(value: object) -> str:
    """Fecha de entrega escrita a mano -> 'MM/DD/YY' ('' si esta vacia o no se entiende).

    Acepta MM/DD/YY, MM/DD/YYYY, MM-DD-YY, YYYY-MM-DD y objetos date/datetime.
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return f"{value:%m/%d/%y}"
    if isinstance(value, date):
        return f"{value:%m/%d/%y}"
    text = str(value).strip()
    if not text:
        return ""
    iso = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?", text)
    if iso:
        year, month, day = (int(g) for g in iso.groups())
    else:
        match = re.fullmatch(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2}|\d{4})", text)
        if not match:
            return ""
        month, day, year = (int(g) for g in match.groups())
        if year < 100:
            year += 2000
    try:
        return f"{date(year, month, day):%m/%d/%y}"
    except ValueError:
        return ""


def normalize_time(value: object) -> str:
    """Hora escrita a mano -> 'HHMM' ('' si esta vacia o no se entiende). Acepta 1600, 16:00, 4pm, 4:30 PM."""
    if value is None:
        return ""
    if isinstance(value, (dtime, datetime)):
        return f"{value:%H%M}"
    text = str(value).strip().lower().replace(".", "")
    if not text:
        return ""
    match = re.fullmatch(r"(\d{1,2})(?::?(\d{2}))?\s*(am|pm|a|p)?", text)
    if not match:
        return ""
    hour, minute, meridian = int(match.group(1)), int(match.group(2) or 0), match.group(3)
    if meridian:
        if not 1 <= hour <= 12:
            return ""
        hour = hour % 12 + (12 if meridian.startswith("p") else 0)
    if hour > 23 or minute > 59:
        return ""
    return f"{hour:02d}{minute:02d}"


# ------------------------------------------------------------------ DIMS


@dataclass
class DimsSheet:
    """Las dimensiones que se reciben (DIMS), leidas por nombre de columna.

    Los archivos llegan con mas o menos columnas (TRACKER, HAZMAT, RACK a veces estan y a veces
    no), asi que nunca se usa la posicion: solo se exige Barcode.
    """

    rows: list[dict[str, object]]  # un dict por renglon (campos de DIMS_HEADERS)
    found: list[str] = field(default_factory=list)  # campos que traia el archivo
    ready: str = ""  # READY DATE, solo si el archivo trae esa columna
    carriers: dict[str, str] = field(default_factory=dict)  # barcode -> BEX/VRP
    hazmat: dict[str, bool] = field(default_factory=dict)  # barcode -> hay HAZMAT
    trackers: dict[str, list[str]] = field(default_factory=dict)  # barcode -> trackers
    sheet_name: str = "DIMS"

    @property
    def barcodes(self) -> list[str]:
        seen: dict[str, None] = {}
        for row in self.rows:
            seen.setdefault(barcode_text(row.get("barcode")), None)
        return list(seen)

    @property
    def missing(self) -> list[str]:
        """Campos opcionales que este archivo no traia (solo informativo)."""
        return [name for name in ("carrier", "hazmat", "tracker") if name not in self.found]


DataSheet = DimsSheet  # nombre anterior


def barcode_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\.0$", "", str(value).strip())


def _norm(text: object) -> str:
    return re.sub(r"\W+", "", str(text if text is not None else "")).lower()


def _clean(value: object) -> object:
    if isinstance(value, str):
        value = value.strip()
    return None if value in ("", None) else value


def _grid_from_file(content: bytes, filename: str) -> tuple[list[list[object]], str]:
    name = filename.lower()
    if name.endswith(".csv"):
        import csv

        text = content.decode("utf-8-sig", errors="replace")
        return [list(row) for row in csv.reader(io.StringIO(text))], "DIMS"
    if name.endswith(".xls"):
        import pandas as pd

        frame = pd.read_excel(io.BytesIO(content), header=None, dtype=object)
        return frame.where(frame.notna(), None).values.tolist(), "DIMS"
    book = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
    try:
        sheets = {sheet.title.strip().upper(): sheet for sheet in book.worksheets}
        order = [sheets[n] for n in ("DIMS", "DATA") if n in sheets] + list(book.worksheets)
        for sheet in order:
            grid = [list(row) for row in sheet.iter_rows(values_only=True)]
            if _header_row(grid) is not None:
                return grid, sheet.title
        return [list(row) for row in book.worksheets[0].iter_rows(values_only=True)], book.worksheets[0].title
    finally:
        book.close()


def _header_row(grid: list[list[object]]) -> int | None:
    for index, row in enumerate(grid[:15]):
        if any(_norm(cell) == "barcode" for cell in row):
            return index
    return None


def read_dims(content: bytes, filename: str = "DIMS.xlsx") -> DimsSheet:
    """Lee un DIMS (.xlsx/.xlsm/.xls/.csv). Solo exige la columna Barcode."""
    grid, sheet_name = _grid_from_file(content, filename)
    if not grid:
        raise ValueError("empty")
    header_at = _header_row(grid)
    if header_at is None:
        raise ValueError("barcode_header")

    columns: dict[str, int] = {}
    for position, cell in enumerate(grid[header_at]):
        key = DIMS_HEADERS.get(_norm(cell))
        if key and key not in columns:  # si un encabezado se repite, vale el primero
            columns[key] = position

    rows: list[dict[str, object]] = []
    for raw in grid[header_at + 1:]:
        row = {key: _clean(raw[pos]) if pos < len(raw) else None for key, pos in columns.items()}
        if row.get("barcode") is None:
            continue  # sin barcode no es una caja
        rows.append(row)

    dims = DimsSheet(rows=rows, found=list(columns), sheet_name=sheet_name)
    first_ready = next((r["ready"] for r in rows if r.get("ready")), None)
    dims.ready = format_ready(first_ready) if first_ready else ""
    for row in rows:
        code = barcode_text(row["barcode"])
        carrier = str(row.get("carrier") or "").strip().upper()
        if carrier and code not in dims.carriers:
            dims.carriers[code] = carrier
        if str(row.get("hazmat") or "").strip().upper() in YES_VALUES:
            dims.hazmat[code] = True
        tracker = str(row.get("tracker") or "").strip()
        if tracker:
            dims.trackers.setdefault(code, []).append(tracker)
    return dims


read_data_sheet = read_dims  # nombre anterior


# ------------------------------------------------------------------ DATE


@dataclass
class ShipmentRow:
    bol: str
    name: str = ""
    email: str = ""
    phone: str = ""
    state: str = ""
    carrier: str = ""  # BEX / VRP
    hazmat: bool = False
    delivery_date: str = ""  # a mano para este BOL ('MM/DD/YY'; BLANK = vacio); '' = el global o el calculado
    delivery_time: str = ""  # a mano para este BOL ('HHMM'; BLANK = vacio); '' = el global o el de siempre


@dataclass
class DateResult:
    bol: str
    service: str
    delivery: str  # MM/DD/YY (el que se escribe: a mano si lo hay, si no el calculado)
    values: dict[str, object]  # letra de columna -> valor
    warnings: list[str] = field(default_factory=list)
    time: str = ""  # HHMM de entrega que se escribe


def _pick(row_value: str, default_value: str, computed: str) -> str:
    """Fila > global > calculado. BLANK en la fila o en el global deja el campo vacio."""
    for value in (row_value, default_value):
        if value == BLANK:
            return ""
        if value:
            return value
    return computed


def build_date_row(
    row: ShipmentRow, ready: str, priority: str = "", default_date: str = "", default_time: str = "",
) -> DateResult:
    """Fila de DATE para un BOL. `ready` = 'MM/DD/YY HHMM' (ya normalizado; puede venir vacio).

    La entrega y su hora salen, en este orden, de: lo escrito para este BOL, lo elegido para todos
    (`default_date`, `default_time`) y, por ultimo, el calculo (service level + READY DATE; hora 1600).
    """
    warnings: list[str] = []
    service = service_label(row.state)
    computed = delivery_text(ready, row.state) if service and ready else ""
    delivery = _pick(row.delivery_date, default_date, computed)
    delivery_hour = _pick(row.delivery_time, default_time, DELIVERY_TIME)
    if not service:
        warnings.append("no_service")
    if not ready:
        warnings.append("no_ready")
    carrier = (row.carrier or "").strip().upper()
    if carrier not in SHIPPER_NUMBER:
        warnings.append("no_carrier")
    bol_value: object = int(row.bol) if row.bol.isdigit() else row.bol
    values: dict[str, object] = {
        "A": STATION if row.bol else "",
        "B": bol_value,
        "E": SHIPPER_NUMBER.get(carrier, ""),
        "Q": ready[:8] if ready else "",
        "R": ready[-5:] if ready else "",
        "S": PICKUP_CLOSE if ready else "",
        "AB": row.name,
        "AC": row.phone,
        "AF": delivery,
        "AG": delivery_hour,
        "AK": service,
        "CO": PRODUCT_CODE,
        "DJ": BILLTO_REF_TYPE,
        "DK": "1" if row.hazmat else "0",
        "DU": row.email,
        "FK": priority,
    }
    return DateResult(row.bol, service, delivery, values, warnings, delivery_hour)


def carrier_from_name(filename: str) -> str:
    upper = (filename or "").upper()
    for carrier in SHIPPER_NUMBER:
        if upper.startswith(carrier) or f"_{carrier}_" in upper:
            return carrier
    return ""


# ---------------------------------------------------------------- escritura


def _write_date_sheet(ws, results: Iterable[DateResult]) -> None:
    headers = load_headers()["DATE"]
    ws.append(headers)
    width = len(headers)
    for result in results:
        line: list[object] = [None] * width
        for letter, value in result.values.items():
            if value in ("", None):
                continue
            line[column_index_from_string(letter) - 1] = value
        ws.append(line)
    ws.freeze_panes = "A2"
    for index in range(1, width + 1):
        ws.column_dimensions[get_column_letter(index)].width = 14
    for letter in ("AB", "AC", "DU"):
        ws.column_dimensions[letter].width = 30


def _write_data_sheet(ws, dims: DimsSheet, results: Mapping[str, DateResult], ready: str, extras: Mapping[str, str]) -> None:
    """DATA con el orden del template, mas READY DATE, POD DATE TIME y los mensajes de cut off."""
    headers = list(load_headers()["DATA"]) + ["TRACKER"]
    ws.append(headers)
    letter_of = dict(DATA_LAYOUT)
    width = len(headers)
    for number, row in enumerate(dims.rows, start=2):
        line: list[object] = [None] * width
        for key, letter in DATA_LAYOUT:
            if key in ("ready", "service", "delivery"):
                continue
            value = row.get(key)
            if value is not None:
                line[column_index_from_string(letter) - 1] = value
        result = results.get(barcode_text(row.get("barcode")))
        if result:
            line[column_index_from_string(letter_of["service"]) - 1] = result.service or None
            line[column_index_from_string(letter_of["delivery"]) - 1] = result.delivery or None
        ws.append(line)
    ws["N2"] = ready or None
    if extras.get("pod"):
        ws["N3"] = "POD DATE TIME"
        ws["N4"] = extras["pod"]
    row_no = 18
    for label in ("both", "left", "arrived"):
        if extras.get(label):
            ws[f"N{row_no}"] = "CUT OFF MESSAGE"
            ws[f"O{row_no}"] = extras[label]
            row_no += 1
    ws.freeze_panes = "A2"
    for index in range(1, width + 1):
        ws.column_dimensions[get_column_letter(index)].width = 14
    ws.column_dimensions["O"].width = 48


def build_workbook(
    shipments: Iterable[ShipmentRow],
    ready: str,
    dims: DimsSheet | None = None,
    include_data: bool = True,
    priority: str = "",
    extras: Mapping[str, str] | None = None,
    default_date: str = "",
    default_time: str = "",
) -> tuple[bytes, list[DateResult]]:
    """Devuelve (xlsx, resultados). Con `dims` e `include_data`, el libro trae DATE y DATA.

    `ready` puede ir vacio: Pickup Date/Time/Close y la entrega calculada quedan en blanco.
    `default_date` / `default_time` se aplican a todos los BOL que no traen su propia entrega a mano.
    """
    results = [build_date_row(row, ready, priority, default_date, default_time) for row in shipments]
    book = Workbook()
    date_ws = book.active
    date_ws.title = "DATE"
    _write_date_sheet(date_ws, results)
    if dims is not None and include_data:
        _write_data_sheet(book.create_sheet("DATA"), dims, {r.bol: r for r in results}, ready, extras or {})
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue(), results


# ---------------------------------------------------------- mensajes y nombre


_DATE_IN_TEXT = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{2,4})")


def _weekday(text: str) -> str:
    match = _DATE_IN_TEXT.search(text or "")
    if not match:
        return ""
    month, day, year = (int(g) for g in match.groups())
    year += 2000 if year < 100 else 0
    try:
        return datetime(year, month, day).strftime("%A").upper()
    except ValueError:
        return ""


def _with_weekday(text: str, enabled: bool) -> str:
    """'11/07/26 10:44' -> 'SATURDAY 11/07/26 10:44' (como los ejemplos de tu hoja DATA)."""
    day = _weekday(text) if enabled else ""
    return f"{day} {text}" if day else text


def cutoff_messages(ready: str, pod: str, weekday: bool = False) -> dict[str, str]:
    """Los tres mensajes que armaba DATA!O18:O20 con READY DATE (N2) y POD DATE TIME (N4)."""
    ready = re.sub(r"\s+", " ", ready or "").strip()
    pod = re.sub(r"\s+", " ", pod or "").strip()
    left = f"LEFT PLANT {_with_weekday(ready, weekday)}" if ready else ""
    arrived = f"ARRIVED AT {_with_weekday(pod, weekday)}" if pod else ""
    both = f"{left} {arrived}" if left and arrived else ""
    return {"both": both, "left": left, "arrived": arrived}


def safe_filename(name: str, fallback: str, extension: str = ".xlsx") -> str:
    """Nombre de archivo valido. Sin nombre, usa `fallback` (el READY DATE)."""
    text = re.sub(r"\.(xlsx|xlsm|xls)$", "", (name or "").strip(), flags=re.IGNORECASE)
    text = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "-", text).strip(" .-")
    if not text:
        text = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "-", (fallback or "").strip()).strip(" .-")
    return f"{text or 'DATE_IMPORT'}{extension}"
