"""Une las paginas de un mismo BOL, aplica la memoria y clasifica cada contacto.

Estados:
* ``ok``    (verde)    todo leido, formato valido y lecturas que coinciden entre si.
* ``doubt`` (amarillo) todo presente pero con alguna senal de duda (lecturas que no
                       coinciden, correccion con memoria, BOL distinto en otra pagina...).
* ``fail``  (rojo)     falta un dato o el formato no es valido: hay que completarlo a mano.

Una lectura que coincide exactamente con un contacto ya confirmado antes (memoria)
pasa a verde aunque el OCR tenga poco margen: el valor leido es *igual* a uno que
ya se verifico, no una suposicion.
"""
from __future__ import annotations

import difflib
import re
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Iterable, Protocol

from contact_extractor import (
    FieldRead,
    PageRead,
    PdfRead,
    _norm_name,
    capture_bytes,
    normalize_phone,
    phone_key,
    prefer_name_coherent_email,
    valid_email,
    valid_name,
    valid_phone,
    vote,
    vote_phone,
)
from service_levels import normalize_state, state_from_zip

# Umbrales de coincidencia entre lecturas (fraccion del peso total a favor).
AGREE_NAME = 0.75
AGREE_EMAIL = 0.80
AGREE_PHONE = 0.80
AGREE_BOL = 0.99
AGREE_STATE = 0.99
SUPPORT_PHONE_PREFIX = 0.5  # el valor exacto (con prefijo) debe tener al menos la mitad del peso

FIELDS = ("name", "email", "phone", "state")

# En letra sin remate, "I" (i mayuscula) y "l" (ele minuscula) son casi iguales. Una "I" pegada a
# minusculas o una "l" pegada a mayusculas es sospechosa (p. ej. "Imco.com" por "lmco.com").
_AMBIGUOUS_GLYPH = re.compile(r"(?<=[a-z])I|I(?=[a-z])|(?<=[A-Z])l|l(?=[A-Z])")


class Memory(Protocol):  # lo que `contact_memory.ContactMemory` expone
    def lookup(self, email: str) -> dict | None: ...
    def closest_email(self, email: str, cutoff: float = 0.92) -> str | None: ...
    def closest_domain(self, domain: str, cutoff: float = 0.9) -> str | None: ...
    def state_for_zip(self, zip_code: str) -> str: ...


@dataclass
class Contact:
    key: str
    bol: str
    source_file: str
    pages: list[int]
    carrier: str = ""  # resultado final (confirmado entre las fuentes de abajo)
    carrier_sources: dict[str, str] = field(default_factory=dict)  # pdf / dims / file -> BEX|VRP
    name: str = ""
    email: str = ""
    phone: str = ""
    state: str = ""
    zip_code: str = ""
    status: str = "fail"
    reasons: list[str] = field(default_factory=list)
    capture: bytes = b""
    ocr: dict[str, str] = field(default_factory=dict)  # lo que leyo el OCR, antes de memoria
    agreement: dict[str, float] = field(default_factory=dict)
    fixes: dict[str, str] = field(default_factory=dict)  # campo -> valor leido por OCR antes de corregir
    confirmed: bool = False  # una persona lo reviso y lo acepto
    learned: bool = False  # ya se envio a la memoria
    error: str = ""  # mensaje de la ultima accion fallida (clave de traduccion)

    def values(self) -> dict[str, str]:
        return {f: getattr(self, f) for f in FIELDS}


def carrier_from_filename(filename: str) -> str:
    stem = filename.upper()
    for carrier in ("BEX", "VRP"):
        if stem.startswith(carrier):
            return carrier
    return ""


def resolve_carrier(pdf: str = "", dims: str = "", file: str = "") -> tuple[str, str]:
    """Carrier final y motivo de duda ("" si todo coincide).

    Se identifica por el contenido del PDF (Ship From, como la pestana 1) y se confirma con la
    columna CARRIER del DIMS y con el nombre del archivo. Si no coinciden manda la mayoria y,
    en empate, lo que dice el propio BOL; en ese caso se marca para revisar.
    """
    sources = {"pdf": pdf, "dims": dims, "file": file}
    found = {k: v for k, v in sources.items() if v in ("BEX", "VRP")}
    if not found:
        return "", "carrier_unknown"
    values = list(found.values())
    if len(set(values)) == 1:
        return values[0], ""
    ranking = sorted(set(values), key=lambda c: (-values.count(c), "pdf" not in found or found["pdf"] != c))
    return ranking[0], "carrier_mismatch"


def set_carrier(contact: "Contact", dims: str = "") -> None:
    """Calcula contact.carrier con todas las fuentes y deja (o quita) la bandera de duda."""
    sources = dict(contact.carrier_sources)
    if dims:
        sources["dims"] = dims
    contact.carrier_sources = sources
    carrier, reason = resolve_carrier(sources.get("pdf", ""), sources.get("dims", ""), sources.get("file", ""))
    contact.carrier = carrier
    contact.reasons = [r for r in contact.reasons if not r.startswith("carrier_")]
    if reason:
        contact.reasons.append(reason)


def _merge(fields: list[FieldRead], normalizer=lambda s: s, prefer=None) -> FieldRead:
    reads: list[tuple[str, float | None]] = []
    weights: list[float] = []
    for item in fields:
        reads.extend(item.reads)
        weights.extend(item.weights or [1.0] * len(item.reads))
    return vote(reads, normalizer, weights, prefer)


def _merge_phone(fields: list[FieldRead]) -> FieldRead:
    reads: list[tuple[str, float | None]] = []
    weights: list[float] = []
    for item in fields:
        reads.extend(item.reads)
        weights.extend(item.weights or [1.0] * len(item.reads))
    return vote_phone(reads, weights)


def _majority(values: list[str]) -> str:
    values = [v for v in values if v]
    return max(set(values), key=values.count) if values else ""


def _edit_distance(a: str, b: str) -> int:
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def _ratio(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()


# ------------------------------------------------------------ unir paginas


def merge_pages(pdfs: Iterable[PdfRead]) -> list[Contact]:
    """Una contacto por BOL. Las paginas 1 OF 2 / 2 OF 2 se leen como dos lecturas."""
    groups: dict[str, list[tuple[str, PageRead]]] = defaultdict(list)
    for pdf in pdfs:
        for page in pdf.pages:
            if page.kind != "bol":
                continue
            key = page.bol.value or f"?{pdf.filename}#{page.page_no}"
            groups[key].append((pdf.filename, page))

    contacts: list[Contact] = []
    for key, items in groups.items():
        pages = [p for _, p in items]
        filename = items[0][0]
        name = _merge([p.name for p in pages], _norm_name)
        email = _merge([p.email for p in pages], str.lower)
        email, picked = prefer_name_coherent_email(email, name.value)
        phone = _merge_phone([p.phone for p in pages])
        state = _merge([p.state for p in pages])
        bol = _merge([p.bol for p in pages])

        contact = Contact(
            key=key,
            bol=bol.value,
            source_file=filename,
            pages=sorted(p.page_no for p in pages),
            carrier_sources={
                k: v for k, v in (
                    ("pdf", _majority([p.carrier for p in pages])),
                    ("file", carrier_from_filename(filename)),
                ) if v
            },
            name=name.value,
            email=email.value,
            phone=phone.value,
            state=state.value,
            zip_code=next((p.zip_code for p in pages if p.zip_code), ""),
            agreement={
                "name": name.agreement, "email": email.agreement, "phone": phone.agreement,
                "state": state.agreement, "bol": bol.agreement,
                "phone_prefix": phone.support,
            },
        )
        # La mejor captura: la pagina donde el contacto esta completo y mas consistente.
        best = max(
            pages,
            key=lambda p: (
                p.name.present + p.email.present + p.phone.present + p.state.present,
                p.name.agreement + p.email.agreement + p.phone.agreement,
                -p.page_no,
            ),
        )
        contact.capture = getattr(best, "capture_jpeg", b"") or capture_bytes(best.capture)

        # Banderas por pagina
        if picked or any("email_by_name" in p.reasons for p in pages):
            contact.reasons.append("email_by_name")
        for flag in (
            "state_fixed", "state_from_zip", "state_from_name",
            "destination_mexico", "destination_latam", "destination_unsure",
        ):
            if any(flag in p.reasons for p in pages):
                contact.reasons.append(flag)
        if any(p.email_near_edge for p in pages):
            contact.reasons.append("email_near_edge")
        if all(p.empty_block for p in pages):
            contact.reasons.append("empty_block")
        tracking = {p.tracking_value for p in pages if p.tracking_value}
        if tracking and tracking != {contact.bol}:
            contact.reasons.append("tracking_mismatch")
        if len(pages) > 1:
            for fname, norm in (("name", _norm_name), ("email", str.lower), ("phone", phone_key), ("state", lambda s: s)):
                values = {norm(getattr(p, fname).value) for p in pages if getattr(p, fname).value}
                if len(values) > 1:
                    contact.reasons.append(f"pages_disagree:{fname}")
        set_carrier(contact)
        contact.ocr = {**contact.values(), "bol": contact.bol}
        contacts.append(contact)
    return contacts


# --------------------------------------------------------------- memoria


def apply_memory(contact: Contact, memory: Memory | None) -> None:
    """Corrige con lo ya confirmado antes. Cada correccion queda marcada para revision."""
    if memory is None or not contact.email:
        return

    entry = memory.lookup(contact.email)
    if entry is None:
        near = memory.closest_email(contact.email)
        if near and near.lower() != contact.email.lower():
            contact.fixes["email"] = contact.email
            contact.email = near
            contact.reasons.append("memory_fix:email")
            entry = memory.lookup(near)
        else:
            local, _, domain = contact.email.partition("@")
            known = memory.closest_domain(domain.lower())
            if known and known != domain.lower():
                contact.fixes["email"] = contact.email
                contact.email = f"{local}@{known}"
                contact.reasons.append("memory_fix:domain")
    if entry:
        known_name = entry.get("name", "")
        if known_name and _norm_name(known_name) != _norm_name(contact.name) and _ratio(known_name, contact.name) >= 0.8:
            contact.fixes["name"] = contact.name
            contact.name = known_name
            contact.reasons.append("memory_fix:name")
        phones = list((entry.get("phones") or {}).keys())
        if phones and contact.phone not in phones:
            near_phone = next(
                (p for p in phones if len(p) == len(contact.phone) and _edit_distance(p, contact.phone) <= 2), None
            )
            if near_phone:
                contact.fixes["phone"] = contact.phone
                contact.phone = near_phone
                contact.reasons.append("memory_fix:phone")

    if contact.zip_code:
        known_state = memory.state_for_zip(contact.zip_code)
        if known_state and known_state != contact.state:
            contact.fixes["state"] = contact.state
            contact.state = known_state
            contact.reasons.append("memory_fix:state")


def is_memory_exact(contact: Contact, memory: Memory | None) -> bool:
    """El contacto leido es identico a uno ya confirmado (correo + nombre + telefono)."""
    if memory is None or not (contact.email and contact.name and contact.phone):
        return False
    entry = memory.lookup(contact.email)
    return bool(
        entry
        and _norm_name(entry.get("name", "")) == _norm_name(contact.name)
        and contact.phone in (entry.get("phones") or {})
    )


# --------------------------------------------------------- BOL contra DATA


def snap_to_data(contact: Contact, barcodes: set[str]) -> None:
    """Si el BOL leido (con dudas) no esta en DATA pero hay uno a 1 digito, lo propone."""
    if not barcodes or not contact.bol or contact.bol in barcodes:
        return
    # Los BOL son consecutivos (3049111, 3049211...): solo se ajusta si la lectura del propio
    # BOL fue dudosa; si todas las lecturas coinciden, se respeta y se avisa que no esta en DATA.
    unsure = contact.agreement.get("bol", 0.0) < AGREE_BOL or "tracking_mismatch" in contact.reasons
    near = [b for b in barcodes if len(b) == len(contact.bol) and _edit_distance(b, contact.bol) == 1] if unsure else []
    if len(near) == 1:
        contact.fixes["bol"] = contact.bol
        contact.bol = near[0]
        contact.reasons.append("bol_snapped")
    else:
        contact.reasons.append("bol_not_in_data")


# ---------------------------------------------------------- clasificacion


def classify(
    contact: Contact, memory: Memory | None = None, check_zip: bool = True, need_contact: bool = True,
) -> Contact:
    """Asigna status y razones. No cambia los valores (eso lo hizo apply_memory).

    Con ``need_contact=False`` (pestana Inbonds) no se exige nombre, correo ni telefono.
    """
    hard: list[str] = []
    if need_contact and "empty_block" in contact.reasons:
        hard.append("empty_block")
    if not contact.bol or contact.key.startswith("?"):
        hard.append("missing:bol")
    for fname in (("name", "email", "phone") if need_contact else ()):
        value = getattr(contact, fname)
        if not value:
            if "empty_block" not in contact.reasons:
                hard.append(f"missing:{fname}")
        elif not {"name": valid_name, "email": valid_email, "phone": valid_phone}[fname](value):
            hard.append(f"invalid:{fname}")
    if not contact.state:
        hard.append("missing:state")
    elif not normalize_state(contact.state):
        hard.append("invalid:state")

    if check_zip and contact.zip_code and contact.state and "state_from_zip" not in contact.reasons:
        by_zip = state_from_zip(contact.zip_code)
        if by_zip and by_zip != normalize_state(contact.state):
            contact.reasons.append("state_zip_mismatch")

    if (
        need_contact
        and contact.email
        and _AMBIGUOUS_GLYPH.search(contact.email)
        and not (memory and memory.lookup(contact.email))
        and "ambiguous_glyph" not in contact.reasons
    ):
        contact.reasons.append("ambiguous_glyph")

    if need_contact and contact.name and re.search(r"[\d/]", contact.name) and "name_odd" not in contact.reasons:
        contact.reasons.append("name_odd")  # p. ej. "GLADYS MORASALOMON /4509030732": decide la persona

    soft: list[str] = []
    # Una correccion aproximada (memory_fix) es una suposicion: nunca cuenta como exacta.
    fixed_by_memory = any(r.startswith("memory_fix") for r in contact.reasons)
    memory_exact = not fixed_by_memory and is_memory_exact(contact, memory)
    if not hard and not memory_exact:
        thresholds = (
            ("name", AGREE_NAME), ("email", AGREE_EMAIL), ("phone", AGREE_PHONE),
            ("state", AGREE_STATE), ("bol", AGREE_BOL), ("phone_prefix", SUPPORT_PHONE_PREFIX),
        ) if need_contact else (("state", AGREE_STATE), ("bol", AGREE_BOL))
        for fname, limit in thresholds:
            if contact.agreement.get(fname, 0.0) < limit:
                soft.append(f"lowagree:{fname}")
    informational = {"bol_not_in_data", "bol_not_in_dims", "destination_mexico", "destination_latam"}
    flags = [r for r in contact.reasons if r not in informational and not r.startswith("memory_exact")]
    if memory_exact:
        # La coincidencia exacta con memoria absorbe dudas del OCR, no las de BOL/estado.
        flags = [r for r in flags if r.startswith(("tracking_mismatch", "bol_snapped", "pages_disagree:state"))]
        contact.reasons.append("memory_exact")

    contact.reasons = _dedupe(
        [r for r in contact.reasons if r in informational or r.startswith(("memory_exact", "memory_fix"))]
        + hard + flags + soft
    )
    if hard:
        contact.status = "fail"
    elif flags or soft:
        contact.status = "doubt"
    else:
        contact.status = "ok"
    return contact


def _dedupe(items: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out


def build_contacts(
    pdfs: Iterable[PdfRead],
    memory: Memory | None = None,
    data_barcodes: set[str] | None = None,
    dims_carriers: dict[str, str] | None = None,
    need_contact: bool = True,
) -> list[Contact]:
    """Une paginas, aplica memoria, compara con el DIMS (BOL y carrier) y clasifica."""
    contacts = merge_pages(pdfs)
    for contact in contacts:
        if need_contact:
            apply_memory(contact, memory)
        if data_barcodes:
            snap_to_data(contact, data_barcodes)
        if dims_carriers is not None:
            set_carrier(contact, dims_carriers.get(contact.bol, ""))
        classify(contact, memory, need_contact=need_contact)
    return contacts


def reclassify(contact: Contact, memory: Memory | None = None, need_contact: bool = True) -> Contact:
    """Vuelve a calcular el estado de una fila editada a mano por el usuario."""
    contact.phone = normalize_phone(contact.phone) if contact.phone else ""
    contact.agreement = {k: 1.0 for k in contact.agreement}  # lo escribio una persona
    destination = {"destination_mexico": "TX", "destination_latam": "FL"}
    contact.reasons = [
        r for r in contact.reasons
        if r.startswith(("memory_fix", "bol_not_in_data", "carrier_"))
        # La nota de destino solo sigue valiendo si el estado no se cambio a mano.
        or (r in destination and normalize_state(contact.state) == destination[r])
    ]
    return classify(contact, memory, check_zip=False, need_contact=need_contact)  # el estado lo reviso una persona
