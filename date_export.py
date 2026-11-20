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

En DATA se rellenan Q (SERVICE LEVEL) y R (DELIVERY DATE) en cada renglon del BOL.
"""
from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass, field
from datetime import datetime
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

DATA_COLS = {"carrier": "A", "barcode": "C", "hazmat": "J", "ready": "N", "service": "Q", "delivery": "R"}


def load_headers() -> dict[str, list[str]]:
    return json.loads(HEADERS_FILE.read_text(encoding="utf-8"))


def format_ready(value: str | datetime | None) -> str:
    """Normaliza el READY DATE a 'MM/DD/YY HHMM' (como esta en DATA!N2)."""
    parsed = parse_ready(value)
    if parsed is None:
        return ""
    return f"{parsed:%m/%d/%y} {parsed:%H%M}"


# ------------------------------------------------------------------ DATA


@dataclass
class DataSheet:
    """Lo que el usuario pego en DATA (o la salida del convertidor de dimensiones)."""

    headers: list[str]
    rows: list[list[object]]  # columnas A..R
    ready: str = ""  # N2, tal cual
    carriers: dict[str, str] = field(default_factory=dict)  # barcode -> BEX/VRP
    hazmat: dict[str, bool] = field(default_factory=dict)  # barcode -> hay HAZMAT = Y
    sheet_name: str = "DATA"

    @property
    def barcodes(self) -> list[str]:
        seen: dict[str, None] = {}
        for code in self._codes():
            seen.setdefault(code, None)
        return list(seen)

    def _codes(self) -> list[str]:
        idx = column_index_from_string(DATA_COLS["barcode"]) - 1
        return [barcode_text(row[idx]) if idx < len(row) else "" for row in self.rows]


def barcode_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\.0$", "", str(value).strip())


def _clean(value: object) -> object:
    return None if value == "" else value


def read_data_sheet(content: bytes, filename: str = "DATA.xlsx") -> DataSheet:
    """Lee DATA desde .xlsx/.xlsm (hoja 'DATA' o la primera) o .csv."""
    name = filename.lower()
    if name.endswith(".csv"):
        import csv

        text = content.decode("utf-8-sig", errors="replace")
        grid = [list(row) for row in csv.reader(io.StringIO(text))]
        sheet_name = "DATA"
    elif name.endswith(".xls"):
        import pandas as pd

        frame = pd.read_excel(io.BytesIO(content), header=None, dtype=object)
        grid = frame.where(frame.notna(), None).values.tolist()
        sheet_name = "DATA"
    else:
        book = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
        sheet = book["DATA"] if "DATA" in book.sheetnames else book.worksheets[0]
        sheet_name = sheet.title
        grid = [list(row) for row in sheet.iter_rows(values_only=True)]
        book.close()
    if not grid:
        raise ValueError("empty")

    headers = [("" if h is None else str(h)) for h in grid[0][:18]]
    headers += [""] * (18 - len(headers))
    expected = load_headers()["DATA"]
    if _norm(headers[2]) != _norm(expected[2]):
        raise ValueError("barcode_header")

    rows: list[list[object]] = []
    for raw in grid[1:]:
        row = list(raw[:18]) + [None] * (18 - len(raw[:18]))
        row = [_clean(v) for v in row]
        if row[2] in (None, ""):
            continue  # sin barcode no es una caja
        rows.append(row)

    sheet_data = DataSheet(headers=expected, rows=rows, sheet_name=sheet_name)
    ready_col = column_index_from_string(DATA_COLS["ready"]) - 1
    first_ready = grid[1][ready_col] if len(grid) > 1 and ready_col < len(grid[1]) else None
    sheet_data.ready = format_ready(first_ready) if first_ready else ""
    carrier_col = column_index_from_string(DATA_COLS["carrier"]) - 1
    hazmat_col = column_index_from_string(DATA_COLS["hazmat"]) - 1
    for row in rows:
        code = barcode_text(row[2])
        carrier = str(row[carrier_col] or "").strip().upper()
        if carrier and code not in sheet_data.carriers:
            sheet_data.carriers[code] = carrier
        if str(row[hazmat_col] or "").strip().upper() == "Y":
            sheet_data.hazmat[code] = True
    return sheet_data


def _norm(text: str) -> str:
    return re.sub(r"\W+", "", text or "").lower()


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


@dataclass
class DateResult:
    bol: str
    service: str
    delivery: str  # MM/DD/YY
    values: dict[str, object]  # letra de columna -> valor
    warnings: list[str] = field(default_factory=list)


def build_date_row(row: ShipmentRow, ready: str) -> DateResult:
    """Fila de DATE para un BOL. `ready` = 'MM/DD/YY HHMM' (ya normalizado)."""
    warnings: list[str] = []
    service = service_label(row.state)
    delivery = delivery_text(ready, row.state) if service and ready else ""
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
        "AG": DELIVERY_TIME,
        "AK": service,
        "CO": PRODUCT_CODE,
        "DJ": BILLTO_REF_TYPE,
        "DK": "1" if row.hazmat else "0",
        "DU": row.email,
    }
    return DateResult(row.bol, service, delivery, values, warnings)


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


def _write_data_sheet(ws, data: DataSheet, results: Mapping[str, DateResult]) -> None:
    ws.append(data.headers)
    service_idx = column_index_from_string(DATA_COLS["service"]) - 1
    delivery_idx = column_index_from_string(DATA_COLS["delivery"]) - 1
    for row in data.rows:
        out = list(row)
        result = results.get(barcode_text(out[2]))
        if result:
            out[service_idx] = result.service or None
            out[delivery_idx] = result.delivery or None
        ws.append(out)
    ws.freeze_panes = "A2"
    for index in range(1, 19):
        ws.column_dimensions[get_column_letter(index)].width = 14


def build_workbook(
    shipments: Iterable[ShipmentRow],
    ready: str,
    data: DataSheet | None = None,
    include_data: bool = True,
) -> tuple[bytes, list[DateResult]]:
    """Devuelve (xlsx, resultados). Con `data` e `include_data`, el libro trae DATE y DATA."""
    results = [build_date_row(row, ready) for row in shipments]
    book = Workbook()
    date_ws = book.active
    date_ws.title = "DATE"
    _write_date_sheet(date_ws, results)
    if data is not None and include_data:
        _write_data_sheet(book.create_sheet("DATA"), data, {r.bol: r for r in results})
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue(), results
