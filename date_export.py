"""Genera el Excel de importacion (hoja DATE) y lee el DIMS para compararlo contra los BOL.

Solo se descarga la hoja DATE, con los valores ya calculados (sin formulas) y SIN formato: ni anchos de
columna, ni paneles inmovilizados, ni estilos, y las celdas vacias quedan realmente vacias. Asi se importa
tal cual; con el formato de antes habia que copiar la hoja y pegarla como "solo valores".

Columnas de DATE que se llenan (el resto queda vacio, como en el template):
    A  Handling Station        "DFW" cuando hay Housebill
    B  Housebill               BOL
    E  Shipper Number          BEX -> 149570, VRP -> 144463
    Q  Pickup Date             primeros 8 caracteres del READY DATE (MM/DD/YY)
    R  Pickup Time             ultimos 5 caracteres del READY DATE (como RIGHT(N2,5))
    S  Pickup Close Time       "1700"
    T  Consignee Number        147229 (Ship To 11905 Conly Road, Laredo) o 210475 (12430 NW 25th Street, Miami);
                               en cualquier otra direccion queda vacio
    AB Consignee Contact       nombre            <- Ship To Attention
    AC Consignee Phone         telefono          <- Ship To Attention
    AK Service                 "3D", "4D"...     (segun el estado del Ship To)
    CO Product Code            "HPDOM"
    DJ BilltoRefType           "RAX"
    DK BilltoRef               "1" si el BOL es HAZMAT, si no "0"
    DU Consignee Email         correo            <- Ship To Attention
    FK Priority Code           "INBD" (solo pestana Inbonds)

La fecha y hora de entrega (Scheduled Delivery Date / Time, AF y AG de DATE) NO se escriben en ningun
Excel: no se usan para importar. La entrega solo se calcula para mostrarla (DateResult.delivery).

El DIMS ya no aporta datos al Excel: solo se lee su columna Barcode para verificar que trae los mismos BOL
que los PDF (ver dims_compare.py).
"""
from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterable

import xlsxwriter
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from service_levels import delivery_text, parse_ready, service_label

HEADERS_FILE = Path(__file__).parent / "assets" / "DATE_HEADERS.json"

SHIPPER_NUMBER = {"BEX": "149570", "VRP": "144463"}
STATION = "DFW"
PICKUP_CLOSE = "1700"
PRODUCT_CODE = "HPDOM"
BILLTO_REF_TYPE = "RAX"

PRIORITY_INBOND = "INBD"

# Unico encabezado del DIMS que se usa (sin importar mayusculas, espacios ni signos)
BARCODE_HEADER = "barcode"


def load_headers() -> dict[str, list[str]]:
    return json.loads(HEADERS_FILE.read_text(encoding="utf-8"))


def format_ready(value: str | datetime | None) -> str:
    """Normaliza el READY DATE a 'MM/DD/YY HHMM' (como esta en DATA!N2)."""
    parsed = parse_ready(value)
    if parsed is None:
        return ""
    return f"{parsed:%m/%d/%y} {parsed:%H%M}"


# ------------------------------------------------------------------ DIMS


@dataclass
class DimsSheet:
    """Los BOL (columna Barcode) del DIMS que se recibe, solo para compararlos contra los PDF.

    Los archivos llegan con mas o menos columnas y en cualquier orden, asi que nunca se usa la posicion:
    solo se exige el encabezado Barcode. Una caja por renglon, asi que un BOL puede repetirse.
    """

    barcodes: list[str]  # BOL unicos, en el orden en que aparecen
    boxes: dict[str, int] = field(default_factory=dict)  # BOL -> renglones (cajas) en el DIMS
    sheet_name: str = "DIMS"


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
    position = next(i for i, cell in enumerate(grid[header_at]) if _norm(cell) == BARCODE_HEADER)

    boxes: dict[str, int] = {}
    for raw in grid[header_at + 1:]:
        code = barcode_text(_clean(raw[position]) if position < len(raw) else None)
        if code:  # sin barcode no es una caja
            boxes[code] = boxes.get(code, 0) + 1
    return DimsSheet(barcodes=list(boxes), boxes=boxes, sheet_name=sheet_name)


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
    consignee: str = ""  # Consignee Number (DATE!T): solo para Ship To de Laredo / Miami


@dataclass
class DateResult:
    bol: str
    service: str
    delivery: str  # MM/DD/YY
    values: dict[str, object]  # letra de columna -> valor
    warnings: list[str] = field(default_factory=list)


def build_date_row(row: ShipmentRow, ready: str, priority: str = "") -> DateResult:
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
        "T": row.consignee,
        "AB": row.name,
        "AC": row.phone,
        "AK": service,
        "CO": PRODUCT_CODE,
        "DJ": BILLTO_REF_TYPE,
        "DK": "1" if row.hazmat else "0",
        "DU": row.email,
        "FK": priority,
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
    """Escribe DATE sin ningun formato: sin anchos, sin paneles y sin celdas vacias con texto ''."""
    headers = load_headers()["DATE"]
    for column, header in enumerate(headers):
        ws.write_string(0, column, header)
    for row_no, result in enumerate(results, start=1):
        for letter, value in result.values.items():
            if value in ("", None):
                continue  # la celda queda vacia (no se escribe)
            column = column_index_from_string(letter) - 1
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                ws.write_number(row_no, column, value)
            else:
                ws.write_string(row_no, column, str(value))


def build_workbook(
    shipments: Iterable[ShipmentRow],
    ready: str,
    priority: str = "",
) -> tuple[bytes, list[DateResult]]:
    """Devuelve (xlsx, resultados): un libro con la sola hoja DATE, sin formato."""
    results = [build_date_row(row, ready, priority) for row in shipments]
    buffer = io.BytesIO()
    book = xlsxwriter.Workbook(
        buffer, {"in_memory": True, "strings_to_numbers": False, "strings_to_formulas": False, "strings_to_urls": False}
    )
    _write_date_sheet(book.add_worksheet("DATE"), results)
    book.close()
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
