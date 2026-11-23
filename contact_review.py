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
    MAYBE_AGREEMENT,
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
    carrier_sources: dict[str, str] = field(default_factory=dict)  # pdf / file -> BEX|VRP
    name: str = ""
    email: str = ""
    phone: str = ""
    state: str = ""
    zip_code: str = ""
    consignee: str = ""  # Consignee Number de DATE!T (solo si el Ship To es Laredo o Miami)
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


def resolve_carrier(pdf: str = "", file: str = "") -> tuple[str, str]:
    """Carrier final y motivo de duda ("" si todo coincide).

    Se identifica por el contenido del PDF (Ship From, como la pestana 1) y se confirma con el
    nombre del archivo. Si no coinciden manda lo que dice el propio BOL; en ese caso se marca para revisar.
    """
    found = {k: v for k, v in {"pdf": pdf, "file": file}.items() if v in ("BEX", "VRP")}
    if not found:
        return "", "carrier_unknown"
    if len(set(found.values())) == 1:
        return next(iter(found.values())), ""
    return found.get("pdf") or found["file"], "carrier_mismatch"


def set_carrier(contact: "Contact") -> None:
    """Calcula contact.carrier con todas las fuentes y deja (o quita) la bandera de duda."""
    sources = contact.carrier_sources
    carrier, reason = resolve_carrier(sources.get("pdf", ""), sources.get("file", ""))
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


def _merge_bol(fields: list[FieldRead]) -> FieldRead:
    """Une el BOL de varias paginas usando la confiabilidad que ya calculo cada pagina (`agreement`).

    Si todas leen lo mismo, dos paginas independientes se refuerzan; si difieren, gana la mas fiable y el
    acuerdo baja (la fila queda para revisar).
    """
    present = [f for f in fields if f.value]
    if not present:
        return FieldRead(reads=[r for f in fields for r in f.reads])
    score: dict[str, float] = defaultdict(float)
    for item in present:
        score[item.value] += max(item.agreement, 0.01)
    best = max(score, key=lambda v: (score[v], max(f.conf for f in present if f.value == v)))
    members = [f for f in present if f.value == best]
    if len(score) == 1:
        miss = 1.0
        for item in members:
            miss *= 1.0 - min(item.agreement, 1.0)
        agreement = 1.0 - miss
    else:
        agreement = score[best] / sum(score.values())
        agreement = min(agreement, max(f.agreement for f in members))
    # Un digito dividido en una pagina sigue siendo duda, salvo que otra pagina lo lea sin ninguna duda.
    alts = [f.alt for f in members if f.alt and f.alt != best]
    sure_elsewhere = any(f.agreement >= AGREE_BOL and not f.alt for f in members)
    alt = alts[0] if alts and not sure_elsewhere else ""
    if alt:
        agreement = min(agreement, MAYBE_AGREEMENT)
    reads = [r for f in fields for r in f.reads]
    return FieldRead(
        value=best, conf=sum(f.conf for f in members) / len(members), agreement=agreement, reads=reads,
        weights=[w for f in fields for w in (f.weights or [1.0] * len(f.reads))], support=agreement, alt=alt,
    )


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


def _hamming(a: str, b: str) -> int:
    return sum(x != y for x, y in zip(a, b)) if len(a) == len(b) else 99


def continues_bol(prev: PageRead, page: PageRead) -> bool:
    """La pagina es la continuacion (2 OF 2) de la anterior y sus numeros de BOL no se contradicen de verdad.

    Se exige que el total de paginas coincida y que el BOL, si difiere, sea casi igual (<= 2 digitos) y que al
    menos una de las dos lecturas sea dudosa; dos lecturas fiables y distintas son BOL distintos.
    """
    if not (prev.page_of and page.page_of) or page.page_no != prev.page_no + 1:
        return False
    if page.page_of[1] != prev.page_of[1] or page.page_of[0] != prev.page_of[0] + 1:
        return False
    a, b = prev.bol, page.bol
    if not b.value or not a.value or a.value == b.value:
        return True
    return _hamming(a.value, b.value) <= 2 and (a.agreement < 1.0 or b.agreement < 1.0)


def merge_pages(pdfs: Iterable[PdfRead]) -> list[Contact]:
    """Una contacto por BOL. Las paginas 1 OF 2 / 2 OF 2 se leen como dos lecturas."""
    groups: dict[str, list[tuple[str, PageRead]]] = defaultdict(list)
    for pdf in pdfs:
        prev_key, prev_page = "", None
        for page in pdf.pages:
            if page.kind != "bol":
                prev_key, prev_page = "", None
                continue
            key = page.bol.value or f"?{pdf.filename}#{page.page_no}"
            # "2 OF 2" justo despues de "1 OF 2" es el mismo BOL: si el numero se leyo distinto en una de las dos
            # paginas por un digito mal visto, no se vuelve un BOL mas.
            if prev_page is not None and continues_bol(prev_page, page):
                key = prev_key
            groups[key].append((pdf.filename, page))
            prev_key, prev_page = key, page

    contacts: list[Contact] = []
    for key, items in groups.items():
        pages = [p for _, p in items]
        filename = items[0][0]
        name = _merge([p.name for p in pages], _norm_name)
        email = _merge([p.email for p in pages], str.lower)
        email, picked = prefer_name_coherent_email(email, name.value)
        phone = _merge_phone([p.phone for p in pages])
        state = _merge([p.state for p in pages])
        bol = _merge_bol([p.bol for p in pages])

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
            consignee=_majority([p.consignee for p in pages]),
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
        if bol.alt:
            contact.reasons.append(f"bol_alt:{bol.alt}")
        if any(p.email_near_edge for p in pages):
            contact.reasons.append("email_near_edge")
        if all(p.empty_block for p in pages):
            contact.reasons.append("empty_block")
        # La fila Tracking No de una pagina de continuacion ("2 OF 2") empieza en otra caja: no se compara.
        tracking = {
            p.tracking_value for p in pages if p.tracking_value and not (p.page_of and p.page_of[0] > 1)
        }
        for other in sorted(tracking - {contact.bol}):
            # Un numero casi igual (<= 2 digitos) es una mala lectura de uno de los dos: se revisa. Uno que no se
            # parece en nada es otro numero (no una lectura del BOL): solo se avisa, sin pasar a amarillo.
            kind = "tracking_mismatch" if _hamming(other, contact.bol) <= 2 else "tracking_other"
            contact.reasons.append(f"{kind}:{other}")
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
    unsure = contact.agreement.get("bol", 0.0) < AGREE_BOL or any(r.startswith("tracking_mismatch") for r in contact.reasons)
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
    # Un Ship To Attention vacio en el PDF es normal: queda en verde y el Excel lleva nombre, correo y
    # telefono en blanco. Si ya trae algun dato (memoria o captura a mano), deja de ser un bloque vacio.
    empty = "empty_block" in contact.reasons and not (contact.name or contact.email or contact.phone)
    if not empty:
        contact.reasons = [r for r in contact.reasons if r != "empty_block"]
    if not contact.bol or contact.key.startswith("?"):
        hard.append("missing:bol")
    for fname in (("name", "email", "phone") if need_contact and not empty else ()):
        value = getattr(contact, fname)
        if not value:
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
    if not hard:
        # Una coincidencia exacta con la memoria absorbe las dudas del contacto, nunca las del BOL ni del estado.
        thresholds = (
            ("name", AGREE_NAME), ("email", AGREE_EMAIL), ("phone", AGREE_PHONE),
            ("state", AGREE_STATE), ("bol", AGREE_BOL), ("phone_prefix", SUPPORT_PHONE_PREFIX),
        ) if need_contact and not empty and not memory_exact else (("state", AGREE_STATE), ("bol", AGREE_BOL))
        for fname, limit in thresholds:
            if contact.agreement.get(fname, 0.0) < limit:
                soft.append(f"lowagree:{fname}")
    informational = {"bol_not_in_data", "bol_not_in_dims", "destination_mexico", "destination_latam", "empty_block"}
    flags = [r for r in contact.reasons if r not in informational and not r.startswith(("memory_exact", "tracking_other"))]
    if memory_exact:
        # La coincidencia exacta con memoria absorbe dudas del OCR, no las de BOL/estado.
        flags = [
            r for r in flags
            if r.startswith(("tracking_mismatch", "bol_snapped", "bol_alt:", "bol_similar:", "pages_disagree:state"))
        ]
        contact.reasons.append("memory_exact")

    contact.reasons = _dedupe(
        [r for r in contact.reasons if r in informational or r.startswith(("memory_exact", "memory_fix", "tracking_other"))]
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


def flag_similar_bols(contacts: list[Contact]) -> None:
    """Marca un BOL poco fiable que difiere en un solo digito de otro BOL de la lista (posible mala lectura).

    Los BOL consecutivos si pueden diferir en un digito, por eso solo se marca el dudoso y nunca se corrige solo:
    la tarjeta ofrece "Usar ese BOL" y, al confirmar, si ya estaba en la lista queda uno solo.
    """
    known = {c.bol for c in contacts if c.bol}
    for contact in contacts:
        if not contact.bol or contact.agreement.get("bol", 1.0) >= AGREE_BOL:
            continue
        near = sorted(
            b for b in known
            if b != contact.bol and _hamming(b, contact.bol) == 1 and f"bol_alt:{b}" not in contact.reasons
        )
        contact.reasons.extend(f"bol_similar:{b}" for b in near[:2])


def build_contacts(
    pdfs: Iterable[PdfRead],
    memory: Memory | None = None,
    data_barcodes: set[str] | None = None,
    need_contact: bool = True,
) -> list[Contact]:
    """Une paginas, aplica memoria, ajusta el BOL contra el DIMS (si hay) y clasifica."""
    contacts = merge_pages(pdfs)
    for contact in contacts:
        if need_contact:
            apply_memory(contact, memory)
        if data_barcodes:
            snap_to_data(contact, data_barcodes)
    flag_similar_bols(contacts)
    for contact in contacts:
        classify(contact, memory, need_contact=need_contact)
    return contacts


def merge_duplicate_bol(contacts: dict[str, Contact], edited: Contact) -> list[str]:
    """Deja un solo registro por BOL cuando se corrige un BOL a mano y ese numero ya estaba en la lista.

    Si el otro registro ya estaba confirmado se queda ese; si no, se queda el que se acaba de corregir. El que sobra
    se quita de `contacts` (sus paginas se suman al que se queda cuando son del mismo archivo). Devuelve los BOL
    cuyo duplicado se quito (vacio si no habia ninguno).
    """
    if not edited.bol:
        return []
    twins = [c for c in contacts.values() if c is not edited and c.bol == edited.bol]
    if not twins:
        return []
    keep = next((c for c in twins if c.confirmed), edited)
    for contact in (edited, *twins):
        if contact is keep:
            continue
        if contact.source_file == keep.source_file:
            keep.pages = sorted(set(keep.pages) | set(contact.pages))
        contacts.pop(contact.key, None)
    return [edited.bol]


def reclassify(contact: Contact, memory: Memory | None = None, need_contact: bool = True) -> Contact:
    """Vuelve a calcular el estado de una fila editada a mano por el usuario."""
    contact.phone = normalize_phone(contact.phone) if contact.phone else ""
    contact.agreement = {k: 1.0 for k in contact.agreement}  # lo escribio una persona
    destination = {"destination_mexico": "TX", "destination_latam": "FL"}
    contact.reasons = [
        r for r in contact.reasons
        if r.startswith(("memory_fix", "bol_not_in_data", "carrier_"))
        or r == "empty_block"  # classify lo quita si ya se escribio algun dato
        # La nota de destino solo sigue valiendo si el estado no se cambio a mano.
        or (r in destination and normalize_state(contact.state) == destination[r])
    ]
    return classify(contact, memory, check_zip=False, need_contact=need_contact)  # el estado lo reviso una persona
