"""Lectura de contactos (Ship To Attention) desde HPE Bill of Lading escaneados.

Cada pagina se procesa asi:
1. Se renderiza la parte superior de la pagina y se ubican las etiquetas
   ("Bill of Lading #", "Ship To", "Ship To Attention", "Tracking No"). Todo
   el resto se recorta *relativo a esas etiquetas*, no a coordenadas fijas, asi
   que tolera escaneos corridos o con distinta escala (VRP y BEX).
2. El bloque de contacto se lee varias veces con distinto preprocesado y los
   campos se votan. Correo y telefono se localizan caracter por caracter (en BEX
   vienen como "NOMBRE/correo/telefono" y se parten entre lineas) y esos
   fragmentos se vuelven a leer con listas de caracteres permitidos, que es lo
   que mas reduce los errores.
3. Se validan formatos y se cruzan fuentes independientes (BOL # del
   encabezado contra la fila Tracking No, pagina 1 contra pagina 2 del mismo BOL).

La clasificacion final (seguro / dudoso / no leido) vive en `contact_review.py`.
"""
from __future__ import annotations

import difflib
import io
import os
import re
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from statistics import mean
from typing import Callable, Iterable

import fitz  # PyMuPDF
import numpy as np
import pytesseract
from PIL import Image, ImageDraw, ImageFilter, ImageOps

import destinations
from service_levels import load_table, state_from_name, state_from_zip

# Cada llamada a tesseract es un proceso aparte: sin esto los hilos se pisan.
os.environ.setdefault("OMP_THREAD_LIMIT", "1")

RENDER_DPI = 200
ANCHOR_SCALE = 0.75  # los anclajes se buscan a 150 dpi (mas rapido)
TOP_FRACTION = 0.46  # parte superior de la pagina que contiene todo lo necesario

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)+")
PHONE_RE = re.compile(r"(?<![\w@.])\+?\d[\d\-\+\(\)\s\.]{5,}\d")
KNOWN_TLDS = {
    "com", "net", "org", "edu", "gov", "mil", "int", "us", "ca", "mx", "io", "co",
    "biz", "info", "ai", "tech", "cloud", "store", "online", "pr", "uk", "de", "fr",
    "es", "it", "br", "ar", "cl", "pe", "cr", "pa", "gt", "sv", "hn", "ni", "do",
    "jp", "cn", "in", "au", "nz", "app", "dev", "xyz", "pro", "email", "solutions",
}
EMAIL_CHARS = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789@._-+"
PHONE_CHARS = "0123456789+-"
_TABLE_HEADER_WORDS = {
    "hpe", "product", "description", "no", "of", "box-ids", "box", "ids", "gross", "weight",
    "unit", "measure", "volume", "dangerous", "goods", "pack-id", "pack", "id", "hm", "tracking",
}
# Confusiones tipicas del OCR para leer el codigo de estado.
_CONFUSABLE = {"0": "O", "1": "I", "5": "S", "8": "B", "2": "Z", "6": "G"}

Box = tuple[float, float, float, float]


# ----------------------------------------------------------------- modelos


@dataclass(frozen=True)
class OcrWord:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    conf: float
    line_id: tuple[int, int, int]


@dataclass
class FieldRead:
    value: str = ""
    conf: float = 0.0
    agreement: float = 0.0
    reads: list[tuple[str, float | None]] = field(default_factory=list)
    weights: list[float] = field(default_factory=list)
    support: float = 0.0  # peso de las lecturas que coinciden con el valor exacto / peso total

    @property
    def present(self) -> bool:
        return bool(self.value)


@dataclass
class PageRead:
    source_file: str
    page_no: int
    kind: str  # "bol" | "other"
    bol: FieldRead = field(default_factory=FieldRead)
    name: FieldRead = field(default_factory=FieldRead)
    email: FieldRead = field(default_factory=FieldRead)
    phone: FieldRead = field(default_factory=FieldRead)
    state: FieldRead = field(default_factory=FieldRead)
    zip_code: str = ""
    tracking_value: str = ""  # lo que dice la fila Tracking No (si existe)
    carrier: str = ""  # "BEX" / "VRP" segun el bloque Ship From del propio BOL ("" si no se vio)
    consignee: str = ""  # Consignee Number (DATE!T) si el Ship To es Laredo / Miami, si no ""
    empty_block: bool = False
    email_near_edge: bool = False
    reasons: list[str] = field(default_factory=list)
    capture: Image.Image | None = None
    capture_jpeg: bytes = b""  # la captura ya comprimida (ahorra memoria con muchos PDF)
    rotation: int = 0
    skew: float = 0.0
    thumb: Image.Image | None = None  # solo para paginas omitidas (no son BOL)


# -------------------------------------------------------------- utilidades


def _fuzzy(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()


def _otsu(arr: np.ndarray) -> int:
    hist = np.bincount(arr.ravel(), minlength=256).astype(float)
    total = arr.size
    sum_all = float(np.dot(np.arange(256), hist))
    weight_b = 0.0
    sum_b = 0.0
    best = -1.0
    threshold = 127
    for level in range(256):
        weight_b += hist[level]
        if weight_b == 0:
            continue
        weight_f = total - weight_b
        if weight_f == 0:
            break
        sum_b += level * hist[level]
        between = weight_b * weight_f * (sum_b / weight_b - (sum_all - sum_b) / weight_f) ** 2
        if between > best:
            best, threshold = between, level
    return threshold


def _ink_mask(gray: Image.Image) -> np.ndarray:
    arr = np.asarray(gray, dtype=np.uint8)
    if arr.size == 0:
        return np.zeros((0, 0), dtype=bool)
    background = float(np.percentile(arr, 90))
    return arr < background - 70


def _ink_fraction(gray: Image.Image) -> float:
    mask = _ink_mask(gray)
    return float(mask.mean()) if mask.size else 0.0


def _cut_at_rule(gray: Image.Image) -> Image.Image:
    """Corta el recorte donde empieza una linea horizontal (regla de tabla)."""
    mask = _ink_mask(gray)
    if mask.size == 0:
        return gray
    # Cerrar huecos de 1-2 px y medir la racha continua mas larga de cada fila:
    # una regla es una racha casi del ancho completo; una linea de texto, aunque
    # llegue de lado a lado, siempre tiene huecos entre letras.
    closed = mask.copy()
    closed[:, 1:] |= mask[:, :-1]
    closed[:, :-1] |= mask[:, 1:]
    limit = 0.6 * mask.shape[1]
    for y in range(7, mask.shape[0]):
        row = closed[y]
        if row.sum() < limit:
            continue
        edges = np.flatnonzero(np.diff(np.concatenate(([0], row.astype(np.int8), [0]))))
        if edges.size and (edges[1::2] - edges[::2]).max() >= limit:
            return gray.crop((0, 0, gray.width, y - 1))
    return gray


def _variants(gray: Image.Image, count: int = 3) -> list[tuple[Image.Image, int, int]]:
    """Devuelve (imagen, escala, padding) con distintos preprocesados."""
    out: list[tuple[Image.Image, int, int]] = []
    pad = 14
    base3 = ImageOps.autocontrast(
        gray.resize((max(1, gray.width * 3), max(1, gray.height * 3)), Image.LANCZOS), cutoff=1
    )
    out.append((ImageOps.expand(base3, border=pad, fill=255), 3, pad))
    if count >= 2:
        arr = np.asarray(base3, dtype=np.uint8)
        binary = Image.fromarray(((arr > _otsu(arr)) * 255).astype(np.uint8))
        out.append((ImageOps.expand(binary, border=pad, fill=255), 3, pad))
    if count >= 3:
        base4 = gray.resize((max(1, gray.width * 4), max(1, gray.height * 4)), Image.LANCZOS)
        base4 = ImageOps.autocontrast(base4.filter(ImageFilter.UnsharpMask(2, 140, 3)), cutoff=1)
        out.append((ImageOps.expand(base4, border=pad, fill=255), 4, pad))
    return out


def _ocr_words(img: Image.Image, config: str) -> list[OcrWord]:
    data = pytesseract.image_to_data(
        img, lang="eng", config=config, output_type=pytesseract.Output.DICT
    )
    words: list[OcrWord] = []
    for i, raw in enumerate(data["text"]):
        text = str(raw).strip()
        if not text:
            continue
        try:
            conf = float(data["conf"][i])
        except (TypeError, ValueError):
            conf = -1.0
        if conf < 0:
            continue
        left, top = float(data["left"][i]), float(data["top"][i])
        words.append(
            OcrWord(
                text,
                left,
                top,
                left + float(data["width"][i]),
                top + float(data["height"][i]),
                conf,
                (int(data["block_num"][i]), int(data["par_num"][i]), int(data["line_num"][i])),
            )
        )
    return words


def _lines(words: Iterable[OcrWord]) -> list[list[OcrWord]]:
    grouped: dict[tuple[int, int, int], list[OcrWord]] = defaultdict(list)
    for word in words:
        grouped[word.line_id].append(word)
    lines = [sorted(items, key=lambda w: w.x0) for items in grouped.values()]
    lines.sort(key=lambda line: min(w.y0 for w in line))
    return lines


def _mean_conf(words: Iterable[OcrWord]) -> float:
    values = [w.conf for w in words]
    return float(mean(values)) if values else 0.0


def normalize_phone(raw: str) -> str:
    """Deja solo digitos y '+', que es lo que acepta el sistema."""
    return re.sub(r"[^\d+]", "", raw or "")


def _norm_name(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip().casefold()


def valid_email(value: str) -> bool:
    if not EMAIL_RE.fullmatch(value or ""):
        return False
    local, _, domain = value.partition("@")
    tld = domain.rsplit(".", 1)[-1].lower()
    return len(local) >= 2 and ".." not in value and (tld in KNOWN_TLDS or (len(tld) == 2 and tld.isalpha()))


def valid_phone(value: str) -> bool:
    # Sin formato fijo: cualquier digito y signo '+' sirve; solo se pide un largo razonable.
    digits = re.sub(r"\D", "", value or "")
    return 7 <= len(digits) <= 20 and re.fullmatch(r"[\d+]+", value or "") is not None


def valid_name(value: str) -> bool:
    return len(re.findall(r"[A-Za-zÀ-ÿ]", value or "")) >= 2 and len(value) <= 80


# ---------------------------------------------------- separar nombre/correo/telefono


def _clean_name(text: str) -> str:
    text = re.sub(r"\s+", " ", text)
    return re.sub(r"^[\s/|,;:\-_.]+|[\s/|,;:\-_.]+$", "", text)


@dataclass
class Split:
    name: str = ""
    email: str = ""
    phone: str = ""
    name_span: tuple[int, int] | None = None
    email_span: tuple[int, int] | None = None
    phone_span: tuple[int, int] | None = None


def split_contact(flat: str) -> Split:
    """Separa nombre, correo y telefono de un texto de varias lineas.

    VRP: tres lineas (nombre / correo / telefono).
    BEX: "NOMBRE/correo@dominio.com/telefono" partido por el ancho de columna.
    """
    out = Split()
    email_match = EMAIL_RE.search(flat)
    cursor = 0
    if email_match:
        out.email = email_match.group(0).strip(".")
        out.email_span = email_match.span()
        cursor = email_match.end()
        out.name_span = (0, email_match.start())
        out.name = _clean_name(flat[: email_match.start()])
    phone_match = PHONE_RE.search(flat, cursor)
    if phone_match:
        out.phone = phone_match.group(0).strip()
        out.phone_span = phone_match.span()
        if not email_match:
            out.name_span = (0, phone_match.start())
            out.name = _clean_name(flat[: phone_match.start()])
    elif not email_match:
        out.name = _clean_name(flat)
        out.name_span = (0, len(flat))
    return out


def _line_right_edges(lines: list[list[OcrWord]]) -> list[float]:
    return [max(w.x1 for w in line) for line in lines]


def _wrapped_flags(edges: list[float], tolerance: float = 0.93) -> list[bool]:
    """Lineas que llegan al borde derecho del bloque: probablemente cortadas por el ancho."""
    if len(edges) < 2:
        return [False] * len(edges)
    limit = tolerance * max(edges)
    return [edge >= limit for edge in edges]


def _join_texts(texts: list[str], wrapped: list[bool]) -> str:
    out = texts[0] if texts else ""
    for index in range(1, len(texts)):
        out += ("" if wrapped[index - 1] else "\n") + texts[index]
    return out


def parse_words(words: list[OcrWord]) -> tuple[Split, float, float, float]:
    """Parseo a partir de palabras OCR. Devuelve (split, conf_nombre, conf_correo, conf_tel)."""
    lines = _lines(words)
    texts = [re.sub(r"\s*@\s*", "@", " ".join(w.text for w in line)) for line in lines]
    texts = [re.sub(r"(@[\w\-]+)\s*\.\s*(?=\w)", r"\1.", t) for t in texts]
    split = split_contact("\n".join(texts))
    if not valid_email(split.email):
        # BEX corta la cadena NOMBRE/correo/telefono por ancho, incluso en medio del correo.
        wrapped = _wrapped_flags(_line_right_edges(lines)) if lines else []
        if any(wrapped[:-1]):
            alt = split_contact(_join_texts(texts, wrapped))
            if valid_email(alt.email):
                split = alt
    name_words, email_words, phone_words = [], [], []
    phone_digits = re.sub(r"\D", "", split.phone)
    for word in (w for line in lines for w in line):
        token = word.text.strip("/|,;:")
        if not token:
            continue
        digits = re.sub(r"\D", "", token)
        if "@" in token or (split.email and token in split.email):
            email_words.append(word)
        elif phone_digits and digits and digits in phone_digits:
            phone_words.append(word)
        elif split.name and token in split.name:
            name_words.append(word)
    return split, _mean_conf(name_words), _mean_conf(email_words), _mean_conf(phone_words)


@dataclass(frozen=True)
class Char:
    ch: str
    x0: float
    y0: float
    x1: float
    y1: float
    line: int


def read_chars(img: Image.Image, scale: float, pad: float) -> tuple[str, list[Char | None]]:
    """Texto reconstruido caracter a caracter con la caja de cada caracter.

    Devuelve (texto, mapa) donde mapa[i] es el Char del indice i del texto, o
    None para los espacios/saltos de linea sinteticos. Las coordenadas quedan
    en el sistema del recorte original (sin escala ni padding).
    """
    raw = pytesseract.image_to_boxes(img, lang="eng", config="--psm 6")
    height = img.height
    chars: list[tuple[str, float, float, float, float]] = []
    for row in raw.splitlines():
        parts = row.rsplit(" ", 5)
        if len(parts) != 6:
            continue
        ch = parts[0]
        left, bottom, right, top = (float(v) for v in parts[1:5])
        chars.append((ch, left, height - top, right, height - bottom))
    text: list[str] = []
    mapping: list[Char | None] = []
    line_no = 0
    prev: tuple[str, float, float, float, float] | None = None
    for ch, x0, y0, x1, y1 in chars:
        if prev is not None:
            h = max(prev[4] - prev[2], y1 - y0, 1.0)
            new_line = (y0 + y1) / 2 - (prev[2] + prev[4]) / 2 > 0.7 * h or x0 < prev[1] - 2 * h
            if new_line:
                line_no += 1
                text.append("\n")
                mapping.append(None)
            elif x0 - prev[3] > 0.45 * h:
                text.append(" ")
                mapping.append(None)
        text.append(ch)
        mapping.append(Char(ch, (x0 - pad) / scale, (y0 - pad) / scale, (x1 - pad) / scale, (y1 - pad) / scale, line_no))
        prev = (ch, x0, y0, x1, y1)
    return "".join(text), mapping


def rejoin_wrapped(text: str, mapping: list[Char | None]) -> tuple[str, list[Char | None]]:
    """Quita los saltos de linea donde el texto fue cortado por ancho (no por espacio)."""
    edges: dict[int, float] = {}
    for item in mapping:
        if item is not None and item.ch.strip():
            edges[item.line] = max(edges.get(item.line, 0.0), item.x1)
    if len(edges) < 2:
        return text, mapping
    limit = 0.93 * max(edges.values())
    out_text: list[str] = []
    out_map: list[Char | None] = []
    for index, (ch, item) in enumerate(zip(text, mapping)):
        if ch == "\n" and item is None:
            prev = next((m for m in reversed(mapping[:index]) if m is not None), None)
            if prev is not None and edges.get(prev.line, 0.0) >= limit:
                continue
        out_text.append(ch)
        out_map.append(item)
    return "".join(out_text), out_map


def span_boxes(span: tuple[int, int] | None, mapping: list[Char | None]) -> list[Box]:
    """Cajas (una por linea) de un tramo del texto de `read_chars`."""
    if not span:
        return []
    by_line: dict[int, list[Char]] = defaultdict(list)
    for item in mapping[span[0]:span[1]]:
        if item is not None and item.ch.strip():
            by_line[item.line].append(item)
    return [
        (min(c.x0 for c in items), min(c.y0 for c in items), max(c.x1 for c in items), max(c.y1 for c in items))
        for _, items in sorted(by_line.items())
    ]


# ---------------------------------------------------------------- estado / zip

STATE_RE = re.compile(r"(?:^|[\s,])([A-Z0-9]{2})\s*,?\s*(\d{5})(?:[\s\-]?\d{4})?(?!\d)")
CA_POSTAL_RE = re.compile(r"(?:^|[\s,])([A-Z]{2})\s+[A-Z]\d[A-Z]\s?\d[A-Z]\d")


def _fix_state_token(token: str) -> tuple[str, bool]:
    """Devuelve (codigo valido, fue_corregido) o ('', False)."""
    table = load_table()
    token = token.upper()
    if token in table:
        return token, False
    fixed = "".join(_CONFUSABLE.get(ch, ch) for ch in token)
    if fixed in table:
        return fixed, True
    return "", False


ZIP_RE = re.compile(r"(?<!\d)(\d{5})(?:[\s\-]?\d{4})?(?!\d)")


def parse_ship_to(words: list[OcrWord]) -> tuple[str, str, str, Box | None]:
    """Devuelve (estado, zip, como, bbox_de_la_linea).

    `como`: "code" (codigo escrito), "fixed" (codigo corregido por confusion de OCR),
    "zip" (deducido del codigo postal) o "name" (nombre completo, p. ej. Puerto Rico).
    """
    lines = _lines(words)
    for line in reversed(lines):
        text = " ".join(w.text for w in line)
        for regex in (STATE_RE, CA_POSTAL_RE):
            match = regex.search(text.upper() if regex is CA_POSTAL_RE else text)
            if not match:
                continue
            code, fixed = _fix_state_token(match.group(1))
            if code:
                zip_code = match.group(2) if regex is STATE_RE else ""
                return code, zip_code, "fixed" if fixed else "code", _line_box(line)
    # Sin codigo de estado escrito: se usa el ZIP o el nombre completo (ultimas 3 lineas).
    for line in reversed(lines[-3:]):
        text = " ".join(w.text for w in line)
        match = ZIP_RE.search(text)
        if match:
            code = state_from_zip(match.group(1))
            if code:
                return code, match.group(1), "zip", _line_box(line)
    for line in reversed(lines[-3:]):
        code = state_from_name(" ".join(w.text for w in line))
        if code:
            zip_match = ZIP_RE.search(" ".join(w.text for w in line))
            return code, zip_match.group(1) if zip_match else "", "name", _line_box(line)
    return "", "", "", None


def _line_box(line: list[OcrWord]) -> Box:
    return (
        min(w.x0 for w in line), min(w.y0 for w in line),
        max(w.x1 for w in line), max(w.y1 for w in line),
    )


# ------------------------------------------------------------------- votos


def _avg_conf(items: Iterable[tuple[str, float | None]]) -> float:
    values = [c for _, c in items if c is not None]
    return float(mean(values)) if values else 0.0


def phone_key(value: str) -> str:
    """Clave para comparar telefonos: ultimos 10 digitos.

    El '+' casi no se distingue en el escaneo y el prefijo de pais ("1-+1") es lo que
    mas se confunde; lo que identifica al numero son los digitos finales.
    """
    return re.sub(r"\D", "", value or "")[-10:]


def prefers_plus(value: str) -> bool:
    return "+" in (value or "")


def vote(
    reads: list[tuple[str, float | None]],
    normalizer: Callable[[str], str] = lambda s: s,
    weights: list[float] | None = None,
    prefer: Callable[[str], bool] | None = None,
) -> FieldRead:
    """Elige el valor con mas peso. agreement = peso a favor / peso total.

    Las lecturas de recorte ajustado con lista de caracteres permitidos pesan mas
    que las lecturas libres del bloque completo.
    """
    if not reads:
        return FieldRead()
    weights = weights or [1.0] * len(reads)
    groups: dict[str, list[tuple[str, float | None, float]]] = defaultdict(list)
    for (value, conf), weight in zip(reads, weights):
        groups[normalizer(value)].append((value, conf, weight))
    non_empty = {k: v for k, v in groups.items() if k}
    if not non_empty:
        return FieldRead(reads=list(reads), weights=list(weights))
    best_key = max(
        non_empty,
        key=lambda k: (sum(w for _, _, w in non_empty[k]), _avg_conf((v, c) for v, c, _ in non_empty[k])),
    )
    members = non_empty[best_key]
    best_value = max(members, key=lambda item: (bool(prefer(item[0])) if prefer else False, item[1] or 0.0))[0]
    return FieldRead(
        value=best_value,
        conf=_avg_conf((v, c) for v, c, _ in members),
        agreement=sum(w for _, _, w in members) / sum(weights),
        reads=list(reads),
        weights=list(weights),
        support=sum(w for _, _, w in members) / sum(weights),
    )


def vote_phone(reads: list[tuple[str, float | None]], weights: list[float] | None = None) -> FieldRead:
    """Voto de telefonos en dos niveles.

    1. Se agrupa por los ultimos 10 digitos (lo que identifica el numero): ese es el `agreement`.
    2. Dentro del grupo ganador se elige el valor completo con mas peso de digitos iguales
       (el prefijo "1-+1" es lo que mas se confunde), prefiriendo la variante que trae '+'.
       El peso de ese valor exacto queda en `support`.
    """
    weights = list(weights) if weights else [1.0] * len(reads)
    base = vote(reads, phone_key, weights)
    if not base.value:
        return base
    key = phone_key(base.value)
    members = [(v, c, w) for (v, c), w in zip(reads, weights) if v and phone_key(v) == key]
    by_digits: dict[str, float] = defaultdict(float)
    for value, _, weight in members:
        by_digits[re.sub(r"\D", "", value)] += weight
    best_digits = max(by_digits, key=lambda d: (by_digits[d], len(d)))
    candidates = [m for m in members if re.sub(r"\D", "", m[0]) == best_digits]
    base.value = max(candidates, key=lambda m: ("+" in m[0], m[1] or 0.0))[0]
    base.conf = _avg_conf((v, c) for v, c, _ in candidates)
    base.support = by_digits[best_digits] / (sum(weights) or 1.0)
    return base


def _email_name_coherence(email: str, name: str) -> float:
    """Que tanto se parece la parte local del correo a un patron tipico del nombre."""
    local = email.split("@")[0].lower()
    tokens = [t for t in re.split(r"[^a-z]+", name.lower()) if t]
    if not local or not tokens:
        return 0.0
    first, last = tokens[0], tokens[-1]
    patterns = {
        first, first + last, first + "." + last, first + "_" + last, first + "-" + last,
        first[0] + last, first[0] + "." + last, last + first, last + "." + first,
        last + first[0], first + last[0], "".join(tokens),
    }
    return max(difflib.SequenceMatcher(None, local, p).ratio() for p in patterns)


def prefer_name_coherent_email(email: FieldRead, name: str) -> tuple[FieldRead, bool]:
    """Desempata correos casi iguales a favor del que cuadra con el nombre.

    Resuelve confusiones como adley / adiey (l-i) sin inventar nada: solo elige
    entre candidatos que el OCR ya leyo. Devuelve (campo, hubo_cambio).
    """
    if not email.reads or not name:
        return email, False
    total = sum(email.weights) or 1.0
    shares: dict[str, float] = defaultdict(float)
    confs: dict[str, list[float]] = defaultdict(list)
    for (value, conf), weight in zip(email.reads, email.weights):
        if value:
            shares[value.lower()] += weight / total
            if conf is not None:
                confs[value.lower()].append(conf)

    if not shares:
        return email, False

    def score(candidate: str) -> float:
        return shares[candidate] + (0.35 if _email_name_coherence(candidate, name) >= 0.92 else 0.0)

    best = max(shares, key=score)
    if best == email.value.lower():
        return email, False
    return FieldRead(
        value=best,
        conf=float(mean(confs[best])) if confs[best] else 0.0,
        agreement=shares[best],
        reads=email.reads,
        weights=email.weights,
    ), True


# ----------------------------------------------------------- lectura de pagina


def detect_carrier(words: list[OcrWord], height: float) -> str:
    """BEX o VRP segun el Ship From impreso en el BOL (misma regla que la pestana 1).

    La pestana 1 trata "IEC TECHNOLOGIES" como BEX y todo lo demas como VRP. Aqui el BOL
    escaneado dice "IEC TECHNOLOGIES S. DE R.L." (BEX) o "c/o Foxconn" (VRP); si no se ve
    ninguno de los dos se devuelve "" para no inventar el carrier.
    """
    tokens = [re.sub(r"[^a-z0-9]", "", w.text.lower()) for w in words if w.y0 < 0.7 * height]
    iec = any(tok in ("iec", "lec", "1ec") for tok in tokens)
    tech = any(_fuzzy(tok, "technologies") >= 0.8 for tok in tokens if len(tok) >= 8)
    bex = (iec and tech) or any(_fuzzy(tok, "inventec") >= 0.85 for tok in tokens if len(tok) >= 6)
    vrp = any(_fuzzy(tok, "foxconn") >= 0.8 for tok in tokens if len(tok) >= 6)
    if bex and not vrp:
        return "BEX"
    if vrp and not bex:
        return "VRP"
    return ""


def render_page(page: fitz.Page, dpi: int = RENDER_DPI) -> Image.Image:
    """Pagina completa: hace falta entera porque algunos escaneos vienen girados."""
    pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


@dataclass
class _Anchors:
    lading_x0: float
    lading_x1: float
    lading_y0: float
    lading_y1: float
    right_x0: float
    att_x0: float
    att_y0: float
    att_y1: float
    ship_x0: float
    ship_y0: float
    ship_y1: float
    table_y: float
    track_x0: float
    track_x1: float
    track_y1: float
    has_tracking: bool
    bol_word: OcrWord | None
    track_word: OcrWord | None


def _find_anchors(words: list[OcrWord], width: float, height: float) -> _Anchors | None:
    lading = [w for w in words if _fuzzy(w.text.strip(":#"), "lading") >= 0.8 and w.x0 > 0.5 * width]
    attention = [w for w in words if _fuzzy(w.text.strip(":"), "attention") >= 0.78 and w.x0 > 0.25 * width]
    if not lading or not attention:
        return None
    lad = min(lading, key=lambda w: w.y0)
    att = min(attention, key=lambda w: w.y0)
    band = 0.012 * height + 6

    same_line_left = [w for w in words if abs(w.y0 - lad.y0) <= band and 0.5 * width < w.x0 <= lad.x0]
    right_x0 = min((w.x0 for w in same_line_left), default=lad.x0)

    att_left = [w for w in words if abs(w.y0 - att.y0) <= band and att.x0 - 140 < w.x0 <= att.x0]
    att_x0 = min((w.x0 for w in att_left), default=att.x0)

    shipto = [
        w for w in words
        if _fuzzy(w.text.strip(":"), "to") >= 0.99 and w.x0 < 0.25 * width and abs(w.y0 - att.y0) <= band * 1.6
    ]
    if shipto:
        to_word = min(shipto, key=lambda w: abs(w.y0 - att.y0))
        left = [w for w in words if abs(w.y0 - to_word.y0) <= band and w.x0 <= to_word.x0]
        ship_x0 = min(w.x0 for w in left)
        ship_y0, ship_y1 = to_word.y0, to_word.y1
    else:
        ship_x0, ship_y0, ship_y1 = 0.03 * width, att.y0, att.y1

    # "TRADING" (nombre de empresa) se parece a "Tracking": exigimos ademas que la
    # misma fila tenga otra palabra propia del encabezado de la tabla.
    strong_header = {"hpe", "product", "description", "hm", "gross", "weight", "pack-id", "box-ids", "volume", "dangerous", "goods", "measure"}
    tracking = [
        w for w in words
        if _fuzzy(w.text, "tracking") >= 0.86 and w.x0 < 0.25 * width and w.y0 > att.y1 + 10
        and any(
            o.x0 > w.x1 and abs(o.y0 - w.y0) <= 0.02 * height and o.text.strip(".:").lower() in strong_header
            for o in words
        )
    ]
    if tracking:
        trk = min(tracking, key=lambda w: w.y0)
        track_x0, track_x1, track_y1, has_tracking = trk.x0, trk.x1, trk.y1, True
        # El encabezado de la tabla puede tener varias lineas ("HPE Product / Description"):
        # el techo real es la palabra mas alta de ese encabezado, no solo "Tracking".
        header_words = [
            w for w in words
            if trk.y0 - 0.06 * height <= w.y0 <= trk.y0 + 4
            and w.text.strip(".:").lower() in _TABLE_HEADER_WORDS
        ]
        table_y = min([trk.y0] + [w.y0 for w in header_words])
    else:
        table_y = min(height - 4, att.y1 + 0.2 * height)
        track_x0, track_x1, track_y1, has_tracking = ship_x0, ship_x0 + 90, table_y + 20, False

    # El numero del BOL y el de la fila Tracking No, tal como los vio el OCR de anclajes.
    bol_word = next(
        (w for w in sorted(words, key=lambda w: w.x0)
         if abs(w.y0 - lad.y0) <= band and w.x0 > lad.x1 - 4 and re.fullmatch(r"\d{6,9}", w.text)),
        None,
    )
    track_word = None
    if has_tracking:
        track_word = next(
            (w for w in sorted(words, key=lambda w: w.y0)
             if track_y1 < w.y0 < track_y1 + 60 and track_x0 - 12 < w.x0 < track_x1 + 200
             and re.fullmatch(r"\d{6,9}", w.text)),
            None,
        )
    return _Anchors(
        lad.x0, lad.x1, lad.y0, lad.y1, right_x0, att_x0, att.y0, att.y1,
        ship_x0, ship_y0, ship_y1, table_y, track_x0, track_x1, track_y1, has_tracking,
        bol_word, track_word,
    )


def _crop(gray: Image.Image, box: Box) -> Image.Image:
    x0, y0, x1, y1 = (int(round(v)) for v in box)
    x0, y0 = max(0, x0), max(0, y0)
    x1, y1 = min(gray.width, max(x1, x0 + 2)), min(gray.height, max(y1, y0 + 2))
    return gray.crop((x0, y0, x1, y1))


def _grow(box: Box, margin: float) -> Box:
    return (box[0] - margin, box[1] - margin, box[2] + margin, box[3] + margin)


def _read_line(gray: Image.Image, whitelist: str, count: int = 2) -> list[tuple[str, float]]:
    """Relee un fragmento de una linea con lista de caracteres permitidos."""
    reads: list[tuple[str, float]] = []
    config = f"--psm 7 -c tessedit_char_whitelist={whitelist} -c load_system_dawg=0 -c load_freq_dawg=0"
    for img, _, _ in _variants(gray, count):
        words = _ocr_words(img, config)
        reads.append(("".join(w.text for w in words), _mean_conf(words)))
    return reads


def _digits(text: str, min_len: int, max_len: int) -> str:
    digits = re.sub(r"\D", "", text)
    return digits if min_len <= len(digits) <= max_len else ""



def estimate_skew(gray: Image.Image) -> float:
    """Angulo (grados) que hay que girar la imagen para enderezar sus lineas de texto.

    Busca el giro que hace mas "afilado" el perfil de filas (las lineas de texto y
    las reglas horizontales quedan alineadas). Rango +-3 grados.
    """
    scale = 700 / max(1, gray.width)
    small = gray.resize((700, max(1, int(gray.height * scale))), Image.BILINEAR)
    arr = np.asarray(small, dtype=np.uint8)
    binary = Image.fromarray(((arr < _otsu(arr)) * 255).astype(np.uint8))

    def sharpness(angle: float) -> float:
        rotated = binary.rotate(angle, resample=Image.NEAREST, fillcolor=0)
        profile = np.asarray(rotated, dtype=np.float32).sum(axis=1)
        return float(np.square(np.diff(profile)).sum())

    coarse = max(np.arange(-3.0, 3.01, 0.5), key=sharpness)
    fine = max(np.arange(coarse - 0.45, coarse + 0.46, 0.1), key=sharpness)
    return float(fine)


def deskew(full: Image.Image, min_angle: float = 0.25) -> tuple[Image.Image, float]:
    top = full.convert("L").crop((0, 0, full.width, int(full.height * TOP_FRACTION)))
    angle = estimate_skew(top)
    if abs(angle) < min_angle:
        return full, 0.0
    return full.rotate(angle, resample=Image.BICUBIC, fillcolor=(255, 255, 255)), angle


def read_page(full: Image.Image, source_file: str, page_no: int, contact: bool = True) -> PageRead:
    """Lee una pagina completa (200 dpi). Si no ve las etiquetas, prueba girada 180 grados."""
    for angle in (0, 180):
        page_img = full if angle == 0 else full.rotate(180)
        page_img, skew = deskew(page_img)
        top = page_img.crop((0, 0, page_img.width, int(page_img.height * TOP_FRACTION)))
        result = _read_top(top, source_file, page_no, contact)
        if result.kind == "bol":
            result.rotation = angle
            result.skew = skew
            result.capture_jpeg = capture_bytes(result.capture)
            result.capture = None
            return result
    thumb = full.copy()
    thumb.thumbnail((360, 480))
    result.thumb = thumb
    return result


def _read_top(img: Image.Image, source_file: str, page_no: int, contact: bool = True) -> PageRead:
    """Lee la parte superior de una pagina ya orientada."""
    gray_full = img.convert("L")
    small = gray_full.resize(
        (int(gray_full.width * ANCHOR_SCALE), int(gray_full.height * ANCHOR_SCALE)), Image.LANCZOS
    )
    words = _ocr_words(small, "--psm 6")
    anchors = _find_anchors(words, small.width, small.height)
    result = PageRead(source_file, page_no, "other")
    if anchors is None:
        return result

    result.kind = "bol"
    result.carrier = detect_carrier(words, small.height)
    k = 1 / ANCHOR_SCALE
    a = anchors
    W, _ = gray_full.size

    # --- cajas en coordenadas de pagina completa (200 dpi) ---------------
    att_box = (a.att_x0 * k - 6, a.att_y1 * k + 2, a.right_x0 * k - 6, a.table_y * k - 6)
    ship_box = (a.ship_x0 * k - 6, a.ship_y1 * k + 2, a.att_x0 * k - 12, a.table_y * k - 6)
    if a.bol_word:
        bol_box = _grow((a.bol_word.x0 * k, a.bol_word.y0 * k, a.bol_word.x1 * k, a.bol_word.y1 * k), 8)
    else:
        bol_box = (a.lading_x1 * k + 30, a.lading_y0 * k - 8, min(W - 4, a.lading_x1 * k + 0.30 * W), a.lading_y1 * k + 8)
    if a.track_word:
        trk_box: Box | None = _grow((a.track_word.x0 * k, a.track_word.y0 * k, a.track_word.x1 * k, a.track_word.y1 * k), 8)
    else:
        trk_box = None

    # --- BOL # (3 fuentes: OCR de anclajes, recorte del encabezado, fila Tracking No)
    bol_reads: list[tuple[str, float | None]] = []
    if a.bol_word:
        bol_reads.append((a.bol_word.text, a.bol_word.conf))
    for text, conf in _read_line(_crop(gray_full, bol_box), "0123456789", count=2):
        bol_reads.append((_digits(text, 6, 9), conf))
    result.bol = vote(bol_reads)
    if trk_box:
        track_reads = [(_digits(t, 6, 9), c) for t, c in _read_line(_crop(gray_full, trk_box), "0123456789", count=2)]
        if a.track_word:
            track_reads.append((a.track_word.text, a.track_word.conf))
        result.tracking_value = vote(track_reads).value

    # --- Ship To Attention ----------------------------------------------
    att_gray = _cut_at_rule(_crop(gray_full, att_box))
    email_box: Box | None = None
    phone_boxes: list[Box] = []
    name_boxes: list[Box] = []
    if not contact:
        pass  # Inbonds: solo hace falta BOL, estado y carrier
    elif _ink_fraction(att_gray) < 0.0015:
        result.empty_block = True
    else:
        name_votes: list[tuple[str, float | None, float]] = []
        email_votes: list[tuple[str, float | None, float]] = []
        phone_votes: list[tuple[str, float | None, float]] = []
        variants = _variants(att_gray, 3)
        for img_v, _, _ in variants:
            split, name_conf, email_conf, phone_conf = parse_words(_ocr_words(img_v, "--psm 6"))
            name_votes.append((split.name, name_conf, 1.0))
            email_votes.append((split.email, email_conf, 1.0))
            phone_votes.append((normalize_phone(split.phone), phone_conf, 1.0))

        # Lectura caracter por caracter: ubica correo y telefono con precision.
        img_a, scale_a, pad_a = variants[0]
        text, mapping = read_chars(img_a, scale_a, pad_a)
        split = split_contact(text)
        if not valid_email(split.email):
            text_alt, mapping_alt = rejoin_wrapped(text, mapping)
            split_alt = split_contact(text_alt)
            if valid_email(split_alt.email):
                text, mapping, split = text_alt, mapping_alt, split_alt
        name_votes.append((split.name, None, 1.0))
        email_votes.append((split.email, None, 1.0))
        phone_votes.append((normalize_phone(split.phone), None, 1.0))
        name_boxes = span_boxes(split.name_span, mapping)
        email_boxes = span_boxes(split.email_span, mapping)
        phone_boxes = span_boxes(split.phone_span, mapping)
        email_box = email_boxes[0] if email_boxes else None

        # Relectura de recorte ajustado con caracteres permitidos (pesa doble).
        if email_box:
            result.email_near_edge = email_box[2] >= att_gray.width - 8
            tight = (email_box[0] - 2, email_box[1] - 5, email_box[2] + 2, email_box[3] + 5)
            for text_read, conf in _read_line(_crop(att_gray, tight), EMAIL_CHARS, count=3):
                match = EMAIL_RE.search(text_read)
                email_votes.append((match.group(0).strip(".") if match else "", conf, 2.0))
        if phone_boxes:
            fragments = [
                _read_line(_crop(att_gray, (b[0] - 2, b[1] - 5, b[2] + 2, b[3] + 5)), PHONE_CHARS, count=3)
                for b in phone_boxes
            ]
            for variant_index in range(3):
                parts = [frag[variant_index] for frag in fragments if variant_index < len(frag)]
                joined = "".join(p[0] for p in parts)
                phone_votes.append(
                    (normalize_phone(joined), float(mean(p[1] for p in parts)) if parts else None, 2.0)
                )

        def tally(items, normalizer=lambda s: s, prefer=None) -> FieldRead:
            return vote([(v, c) for v, c, _ in items], normalizer, [w for _, _, w in items], prefer)

        result.name = tally(name_votes, _norm_name)
        result.email = tally(email_votes, str.lower)
        result.phone = vote_phone([(v, c) for v, c, _ in phone_votes], [w for _, _, w in phone_votes])
        result.email, picked = prefer_name_coherent_email(result.email, result.name.value)
        if picked:
            result.reasons.append("email_by_name")

    # --- Ship To (estado) --------------------------------------------------
    ship_gray = _cut_at_rule(_crop(gray_full, ship_box))
    state_votes: list[tuple[str, float | None]] = []
    state_how: dict[str, set[str]] = defaultdict(set)
    zip_votes: list[str] = []
    state_box: Box | None = None
    dest_votes: list[str] = []
    consignee_votes: list[str] = []
    for img_v, scale, pad in _variants(ship_gray, 2):
        ship_words = _ocr_words(img_v, "--psm 6")
        ship_lines = [" ".join(w.text for w in line) for line in _lines(ship_words)]
        dest_votes.append(destinations.destination_rule(ship_lines))
        consignee_votes.append(destinations.consignee_number(ship_lines))
        state, zip_code, how, bbox = parse_ship_to(ship_words)
        state_votes.append((state, None))
        if state:
            state_how[state].add(how)
            zip_votes.append(zip_code)
            if state_box is None and bbox:
                state_box = ((bbox[0] - pad) / scale, (bbox[1] - pad) / scale, (bbox[2] - pad) / scale, (bbox[3] - pad) / scale)
    result.state = vote(state_votes)
    if zip_votes:
        result.zip_code = max(set(zip_votes), key=zip_votes.count)
    # Un texto tan distintivo como "Conly" casi nunca sale de un error de OCR: basta con una lectura.
    found = [c for c in consignee_votes if c]
    if found:
        result.consignee = max(set(found), key=found.count)
    rules = [r for r in dest_votes if r]
    if rules:
        # Mexico / Latinoamerica: igual que la pestana 1, se entrega en Laredo (TX) o Miami (FL).
        rule = max(set(rules), key=rules.count)
        state = destinations.DESTINATION_STATE[rule]
        sure = len(rules) == len(dest_votes) and len(set(rules)) == 1
        result.state = FieldRead(
            value=state, agreement=1.0 if sure else 0.5, reads=[(state, None)], weights=[1.0], support=1.0
        )
        result.zip_code = ""
        result.reasons.append(f"destination_{rule}")
        if not sure:
            result.reasons.append("destination_unsure")
        state_box = None
    # La marca de "estado deducido" solo si ninguna lectura vio el codigo escrito.
    hows: set[str] = set() if rules else state_how.get(result.state.value, set())
    if hows and "code" not in hows:
        for how, flag in (("fixed", "state_fixed"), ("zip", "state_from_zip"), ("name", "state_from_name")):
            if how in hows:
                result.reasons.append(flag)
                break

    result.capture = _build_capture(
        img, a, k, att_box, ship_box, bol_box, email_box, phone_boxes, name_boxes, state_box, ship_only=not contact
    )
    return result


# ------------------------------------------------------------------ captura


def _build_capture(
    img: Image.Image,
    a: _Anchors,
    k: float,
    att_box: Box,
    ship_box: Box,
    bol_box: Box,
    email_box: Box | None,
    phone_boxes: list[Box],
    name_boxes: list[Box],
    state_box: Box | None,
    ship_only: bool = False,
) -> Image.Image:
    """Recorte tipo 'captura' con lo que se leyo resaltado por colores."""
    rgb = img.copy()
    draw = ImageDraw.Draw(rgb)
    ax, ay = att_box[0], att_box[1]
    colors = {"name": (30, 90, 220), "email": (20, 150, 60), "phone": (230, 120, 0), "state": (150, 40, 170)}

    def rect(box: Box, ox: float, oy: float, color: tuple[int, int, int]) -> None:
        x0, y0, x1, y1 = box
        draw.rectangle((ox + x0 - 3, oy + y0 - 2, ox + x1 + 3, oy + y1 + 2), outline=color, width=3)

    for box in name_boxes:
        rect(box, ax, ay, colors["name"])
    if email_box:
        rect(email_box, ax, ay, colors["email"])
    for box in phone_boxes:
        rect(box, ax, ay, colors["phone"])
    if state_box:
        rect(state_box, ship_box[0], ship_box[1], colors["state"])
    draw.rectangle(bol_box, outline=(200, 30, 30), width=3)

    top = max(0, int(a.ship_y0 * k - 10))
    bottom = min(rgb.height, int(a.table_y * k - 2))
    left = max(0, int(a.ship_x0 * k - 12))
    # Inbonds: solo el Ship To (se corta antes de la columna del Ship To Attention)
    right = min(rgb.width, int((a.att_x0 - 6 if ship_only else a.right_x0) * k - 2))
    main = rgb.crop((left, top, right, max(bottom, top + 20)))

    head_top = max(0, int(a.lading_y0 * k - 10))
    head_bottom = min(rgb.height, int(a.lading_y1 * k + 12))
    head_left = max(0, int(a.right_x0 * k - 6))
    head = rgb.crop((head_left, head_top, min(rgb.width, int(bol_box[2]) + 6), head_bottom))

    width = max(main.width, head.width)
    canvas = Image.new("RGB", (width, main.height + head.height + 8), "white")
    canvas.paste(head, (0, 0))
    canvas.paste(main, (0, head.height + 8))
    if canvas.width > 900:
        ratio = 900 / canvas.width
        canvas = canvas.resize((900, int(canvas.height * ratio)), Image.LANCZOS)
    return canvas


def capture_bytes(image: Image.Image | None, quality: int = 82) -> bytes:
    if image is None:
        return b""
    buffer = io.BytesIO()
    image.save(buffer, format="JPEG", quality=quality)
    return buffer.getvalue()


# --------------------------------------------------------------- PDF completo


@dataclass
class PdfRead:
    filename: str
    pages: list[PageRead]
    total_pages: int

    @property
    def skipped(self) -> list[int]:
        return [p.page_no for p in self.pages if p.kind != "bol"]


def read_pdf(
    pdf_bytes: bytes,
    filename: str,
    *,
    workers: int = 2,
    progress: Callable[[int, int], None] | None = None,
    contact: bool = True,
) -> PdfRead:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total = len(doc)
    pages: list[PageRead] = []
    chunk = max(2, workers * 2)
    done = 0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for start in range(0, total, chunk):
            # PyMuPDF no es thread-safe: se renderiza en serie y solo el OCR va en hilos.
            batch = [(i + 1, render_page(doc[i])) for i in range(start, min(total, start + chunk))]
            futures = [pool.submit(read_page, img, filename, no, contact) for no, img in batch]
            for future in futures:
                pages.append(future.result())
                done += 1
                if progress:
                    progress(done, total)
    doc.close()
    return PdfRead(filename, pages, total)
