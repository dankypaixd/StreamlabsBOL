"""Direcciones de otros paises -> direccion de entrega, con la misma regla de la pestana 1.

En la pestana 1 (`bol_generator.apply_destination_rule`) el Ship To cambia segun el pais:

* Puerto Rico                  -> se queda igual (estado PR);
* Mexico                       -> Laredo, TX (la direccion de `MEXICO_ADDRESS`);
* Latinoamerica (lista fija)   -> Miami, FL (la direccion de `LATAM_ADDRESS`);
* cualquier otro (EE. UU., Canada...) -> se queda igual.

El service level sale del estado de la direccion *a la que realmente se entrega*, asi que un envio a
Mexico se mide como Texas y uno a Brasil como Florida. Aqui se reutilizan la funcion y las direcciones de
la pestana 1 (si cambian alla, cambian aqui), con una precaucion: si el Ship To termina en "US", nunca se
toma como extranjero (evita que "Panama City, FL" o "Mexico, MO" se confundan con un pais).
"""
from __future__ import annotations

import re
from typing import Iterable

import bol_generator as bg
import service_levels as sl

_US_COUNTRY = {"us", "usa", "unitedstates", "unitedstatesofamerica"}
_CA_PROVINCES = {"AB", "BC", "MB", "NB", "NL", "NS", "NT", "ON", "PE", "QC", "SK", "YT"}
_STATE_ZIP = re.compile(r"(?:^|[\s,])([A-Z]{2})\s*,?\s*\d{5}(?:[\s\-]?\d{4})?(?!\d)")
_TRAILING_LABEL = re.compile(r"(?i)^(heart\s+order|order\s+no|packing\s+id|ship\s+via|sold\s+to|ship\s+from)\b")
_ADDRESS_STATE = re.compile(r"\b([A-Z]{2})\s+\d{5}\b")


def _state_of(address: Iterable[str]) -> str:
    match = _ADDRESS_STATE.search(" ".join(address).upper())
    return sl.normalize_state(match.group(1)) if match else ""


# Estado de la direccion de entrega de cada regla (sale de las direcciones de la pestana 1).
DESTINATION_STATE = {"mexico": _state_of(bg.MEXICO_ADDRESS), "latam": _state_of(bg.LATAM_ADDRESS)}
DESTINATION_CITY = {"mexico": "Laredo, TX", "latam": "Miami, FL"}


def _looks_us(rows: list[str]) -> bool:
    """Hay "ST 12345" con un estado de EE. UU. (no provincia de Canada)."""
    for row in rows:
        for match in _STATE_ZIP.finditer(row.upper()):
            code = sl.normalize_state(match.group(1))
            if code and code not in _CA_PROVINCES:
                return True
    return False


def destination_rule(lines: Iterable[str]) -> str:
    """"mexico", "latam" o "" (sin cambio) para las lineas del Ship To."""
    rows = [re.sub(r"\s+", " ", line).strip() for line in lines if line and line.strip()]
    while rows and _TRAILING_LABEL.match(rows[-1]):
        rows.pop()  # la lectura del bloque a veces arrastra el titulo de la seccion que sigue
    if not rows:
        return ""
    if re.sub(r"[^a-z]", "", rows[-1].lower()) in _US_COUNTRY:
        return ""  # termina en US: es una direccion de EE. UU.
    record = bg.apply_destination_rule(bg.BolRecord("", "", "", tuple(rows), 0, "VRP"))
    rule = record.destination_rule
    if rule.startswith("Mexico"):
        key = "mexico"
    elif rule.startswith("Latino"):
        key = "latam"
    else:
        return ""
    last_raw = rows[-1].upper()
    explicit = last_raw == "MX" or last_raw in bg.LATAM_COUNTRY_CODES or re.sub(r"[^a-z ]", "", rows[-1].lower()).strip() in (
        {"mexico"} | bg.LATAM_COUNTRY_NAMES
    )
    if not explicit and _looks_us(rows):
        return ""  # el pais no esta claro y hay un estado de EE. UU. con ZIP: se respeta
    return key


def destination_state(lines: Iterable[str]) -> tuple[str, str]:
    """(regla, estado de entrega) o ("", "") si la direccion no cambia."""
    rule = destination_rule(lines)
    return (rule, DESTINATION_STATE[rule]) if rule else ("", "")


# ---------------------------------------------------------------- Consignee Number (DATE!T)

# Numero de consignatario que lleva cada una de las dos direcciones de entrega de la pestana 1.
CONSIGNEE_NUMBER = {"mexico": "147229", "latam": "210475"}


def address_key(lines: Iterable[str]) -> str:
    """"mexico" / "latam" si el Ship To *ya es* una de las dos direcciones de entrega (por su texto), si no "".

    Se buscan las palabras distintivas, no la direccion completa, porque el OCR puede cambiar alguna letra.
    """
    text = re.sub(r"\s+", " ", " ".join(line for line in lines if line)).lower()
    if "conly" in text or ("11905" in text and ("laredo" in text or "78045" in text)):
        return "mexico"  # 11905 Conly Road, Laredo TX 78045
    if ("12430" in text and ("25th" in text or "33182" in text or "miami" in text)) or ("dsv" in text and "miami" in text):
        return "latam"  # 12430 NW 25th Street, Miami FL 33182 (US_MIAMI DSV Inc.)
    return ""


def consignee_number(lines: Iterable[str]) -> str:
    """Consignee Number del Ship To: el de Laredo o el de Miami, ya sea porque la direccion se convierte
    (pais de Mexico / Latinoamerica) o porque ya viene escrita asi. Cualquier otra direccion: ""."""
    rows = list(lines)
    return CONSIGNEE_NUMBER.get(destination_rule(rows) or address_key(rows), "")
