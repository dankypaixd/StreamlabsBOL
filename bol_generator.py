from __future__ import annotations

from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path
import re
from typing import Iterable, Sequence
from zoneinfo import ZoneInfo

import fitz  # PyMuPDF
from PIL import Image
import pytesseract


DEFAULT_TIMEZONE = "America/Costa_Rica"

MEXICO_ADDRESS = (
    "Corporativo Galvan S.C",
    "11905 Conly Road",
    "Laredo TX 78045",
    "USA",
)

LATAM_ADDRESS = (
    "US_MIAMI DSV Inc. - MIA",
    "12430 NW 25th Street, Suite 100",
    "Miami, FL 33182",
    "US",
)

# Puerto Rico is intentionally excluded because the user requested that it stay unchanged.
LATAM_COUNTRY_NAMES = {
    "argentina",
    "bolivia",
    "brasil",
    "brazil",
    "chile",
    "colombia",
    "costa rica",
    "cuba",
    "dominican republic",
    "republica dominicana",
    "ecuador",
    "el salvador",
    "guatemala",
    "haiti",
    "honduras",
    "nicaragua",
    "panama",
    "paraguay",
    "peru",
    "uruguay",
    "venezuela",
}

LATAM_COUNTRY_CODES = {
    "AR", "BO", "BR", "CL", "CO", "CR", "CU", "DO", "EC", "SV", "GT",
    "HT", "HN", "NI", "PA", "PY", "PE", "UY", "VE",
}


@dataclass(frozen=True)
class Word:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str


@dataclass(frozen=True)
class BolRecord:
    air_waybill: str
    packing_id: str
    heart_order: str
    ship_to_lines: tuple[str, ...]
    source_page: int
    source_format: str
    original_ship_to_lines: tuple[str, ...] = ()
    destination_rule: str = "Original"

    @property
    def display_ship_to(self) -> str:
        return " | ".join(self.ship_to_lines)


class BolExtractionError(ValueError):
    pass


def _clean_text(value: str) -> str:
    return (
        value.replace("\u00a0", " ")
        .replace("\u00ad", "")
        .replace("\ufeff", "")
        .replace("–", "-")
        .strip()
    )


def _normalize(value: str) -> str:
    value = _clean_text(value).lower()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _native_words(page: fitz.Page) -> list[Word]:
    words: list[Word] = []
    for item in page.get_text("words", sort=True):
        x0, y0, x1, y1, text, *_ = item
        text = _clean_text(str(text))
        if text:
            words.append(Word(float(x0), float(y0), float(x1), float(y1), text))
    return words


def _ocr_words(page: fitz.Page, dpi: int = 220) -> list[Word]:
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)

    data = pytesseract.image_to_data(
        image,
        lang="eng",
        config="--psm 6",
        output_type=pytesseract.Output.DICT,
    )

    scale_x = page.rect.width / image.width
    scale_y = page.rect.height / image.height
    words: list[Word] = []
    for index, raw_text in enumerate(data["text"]):
        text = _clean_text(str(raw_text))
        try:
            confidence = float(data["conf"][index])
        except (TypeError, ValueError):
            confidence = -1
        if not text or confidence < 25:
            continue

        left = float(data["left"][index]) * scale_x
        top = float(data["top"][index]) * scale_y
        width = float(data["width"][index]) * scale_x
        height = float(data["height"][index]) * scale_y
        words.append(Word(left, top, left + width, top + height, text))
    return words


def _group_lines(words: Iterable[Word], tolerance: float = 3.0) -> list[list[Word]]:
    groups: list[tuple[float, list[Word]]] = []
    for word in sorted(words, key=lambda item: (item.y0, item.x0)):
        for position, (reference_y, line_words) in enumerate(groups):
            if abs(word.y0 - reference_y) <= tolerance:
                line_words.append(word)
                new_y = sum(item.y0 for item in line_words) / len(line_words)
                groups[position] = (new_y, line_words)
                break
        else:
            groups.append((word.y0, [word]))

    return [sorted(line_words, key=lambda item: item.x0) for _, line_words in groups]


def _line_text(line: Sequence[Word]) -> str:
    return re.sub(r"\s+", " ", " ".join(word.text for word in line)).strip()


def _find_label_value_same_line(
    words: list[Word],
    label_tokens: tuple[str, ...],
    *,
    x_min: float = 0,
    x_max: float = float("inf"),
    y_min: float = 0,
    y_max: float = float("inf"),
    gap: float = 7,
) -> str | None:
    candidates = [
        word for word in words
        if x_min <= word.x0 <= x_max and y_min <= word.y0 <= y_max
    ]
    expected = tuple(_normalize(token) for token in label_tokens)
    for line in _group_lines(candidates, tolerance=3.4):
        normalized = [_normalize(word.text.strip("#:.")) for word in line]
        for start in range(0, len(normalized) - len(expected) + 1):
            if tuple(normalized[start:start + len(expected)]) != expected:
                continue
            label_end = line[start + len(expected) - 1].x1
            values = [word.text for word in line if word.x0 > label_end + gap]
            if values:
                return " ".join(values)
    return None


def _detect_layout(words: list[Word], page: fitz.Page) -> str:
    page_text = _normalize(" ".join(word.text for word in words))
    width = page.rect.width
    if "iec technologies" in page_text or ("order item no" in page_text and width > 700):
        return "BEX"
    return "VRP"


def _extract_ship_to_vrp(words: list[Word], width: float, height: float) -> tuple[str, ...]:
    label_y: float | None = None
    for line in _group_lines(
        [word for word in words if word.x0 < width * 0.22 and word.y0 < height * 0.42]
    ):
        tokens = [_normalize(word.text.strip(":")) for word in line]
        if "ship" in tokens and "to" in tokens:
            label_y = min(word.y0 for word in line)
            break

    if label_y is None:
        return ()

    destination_words = [
        word for word in words
        if word.x0 < width * 0.34 and label_y + 7 <= word.y0 < height * 0.39
    ]

    lines: list[str] = []
    for line_words in _group_lines(destination_words, tolerance=3.2):
        line = _line_text(line_words)
        if line:
            lines.append(line)
    return tuple(lines)


def _extract_heart_order_vrp(words: list[Word], width: float, height: float) -> str | None:
    candidates: list[tuple[float, str]] = []
    for word in words:
        digits = re.sub(r"\D", "", word.text)
        if not (8 <= len(digits) <= 12) or digits.startswith("0000"):
            continue
        if width * 0.07 <= word.x0 <= width * 0.22 and height * 0.43 <= word.y0 <= height * 0.90:
            candidates.append((word.y0, digits))
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def _packing_id_from_raw(raw: str | None) -> str:
    if not raw:
        return ""
    candidates = re.findall(r"[A-Za-z0-9][A-Za-z0-9_-]{4,}", raw)
    return candidates[0] if candidates else ""


def _extract_record_vrp(words: list[Word], page: fitz.Page, source_page: int) -> BolRecord | None:
    width = page.rect.width
    height = page.rect.height

    air_raw = _find_label_value_same_line(
        words, ("air", "waybill"),
        x_min=width * 0.53, x_max=width, y_max=height * 0.35, gap=10,
    )
    air_waybill = re.sub(r"\D", "", air_raw or "")

    packing_raw = _find_label_value_same_line(
        words, ("packing", "id"),
        x_min=width * 0.53, x_max=width, y_max=height * 0.35, gap=10,
    )
    packing_id = _packing_id_from_raw(packing_raw)

    ship_to_lines = _extract_ship_to_vrp(words, width, height)
    heart_order = _extract_heart_order_vrp(words, width, height) or ""

    if not air_waybill:
        return None
    return BolRecord(
        air_waybill=air_waybill,
        packing_id=packing_id,
        heart_order=heart_order,
        ship_to_lines=ship_to_lines,
        source_page=source_page,
        source_format="VRP",
    )


def _extract_ship_to_bex(words: list[Word], width: float, height: float) -> tuple[str, ...]:
    label_y: float | None = None
    for line in _group_lines([word for word in words if word.x0 < width * 0.22 and word.y0 < height * 0.42]):
        tokens = [_normalize(word.text.strip(":")) for word in line]
        if "ship" in tokens and "to" in tokens:
            label_y = min(word.y0 for word in line)
            break
    if label_y is None:
        return ()

    # The BEX format places ORDER ITEM NO. on the same line as SHIP TO.
    # Starting below the label line intentionally ignores that number.
    table_header_y = height * 0.40
    destination_words = [
        word for word in words
        if word.x0 < width * 0.25 and label_y + 6 <= word.y0 < table_header_y
    ]
    lines: list[str] = []
    for line_words in _group_lines(destination_words, tolerance=3.2):
        line = _line_text(line_words)
        if line:
            lines.append(line)
    return tuple(lines)


def _extract_heart_order_bex(words: list[Word], width: float, height: float) -> str | None:
    # Heart Order appears in the first data row under the HEART ORDER NO. header.
    candidates: list[tuple[float, str]] = []
    for word in words:
        digits = re.sub(r"\D", "", word.text)
        if not (8 <= len(digits) <= 12) or digits.startswith("0000"):
            continue
        if width * 0.025 <= word.x0 <= width * 0.14 and height * 0.42 <= word.y0 <= height * 0.70:
            candidates.append((word.y0, digits))
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def _extract_record_bex(words: list[Word], page: fitz.Page, source_page: int) -> BolRecord | None:
    width = page.rect.width
    height = page.rect.height

    air_raw = _find_label_value_same_line(
        words, ("air", "waybill"),
        x_min=width * 0.58, x_max=width, y_max=height * 0.28, gap=20,
    )
    air_waybill = re.sub(r"\D", "", air_raw or "")

    packing_raw = _find_label_value_same_line(
        words, ("packing", "id"),
        x_min=width * 0.58, x_max=width, y_max=height * 0.30, gap=20,
    )
    packing_id = _packing_id_from_raw(packing_raw)

    ship_to_lines = _extract_ship_to_bex(words, width, height)
    heart_order = _extract_heart_order_bex(words, width, height) or ""

    if not air_waybill:
        return None
    return BolRecord(
        air_waybill=air_waybill,
        packing_id=packing_id,
        heart_order=heart_order,
        ship_to_lines=ship_to_lines,
        source_page=source_page,
        source_format="BEX",
    )


def _extract_record(words: list[Word], page: fitz.Page, source_page: int) -> BolRecord | None:
    layout = _detect_layout(words, page)
    if layout == "BEX":
        return _extract_record_bex(words, page, source_page)
    return _extract_record_vrp(words, page, source_page)


def _country_token(lines: Sequence[str]) -> str:
    if not lines:
        return ""
    last = _normalize(lines[-1])
    if last:
        return last
    return _normalize(" ".join(lines))


def apply_destination_rule(record: BolRecord) -> BolRecord:
    original = record.original_ship_to_lines or record.ship_to_lines
    joined = _normalize(" ".join(original))
    last_raw = original[-1].strip().upper() if original else ""
    last = _country_token(original)

    is_puerto_rico = last_raw == "PR" or last in {"pr", "puerto rico"} or "puerto rico" in joined
    if is_puerto_rico:
        return replace(
            record,
            ship_to_lines=tuple(original),
            original_ship_to_lines=tuple(original),
            destination_rule="Puerto Rico - original",
        )

    is_mexico = last_raw == "MX" or last in {"mx", "mexico"} or re.search(r"\bmexico\b", joined) is not None
    if is_mexico:
        return replace(
            record,
            ship_to_lines=MEXICO_ADDRESS,
            original_ship_to_lines=tuple(original),
            destination_rule="Mexico -> Laredo",
        )

    is_latam = last_raw in LATAM_COUNTRY_CODES or last in LATAM_COUNTRY_NAMES
    if not is_latam:
        is_latam = any(re.search(rf"\b{re.escape(country)}\b", joined) for country in LATAM_COUNTRY_NAMES)
    if is_latam:
        return replace(
            record,
            ship_to_lines=LATAM_ADDRESS,
            original_ship_to_lines=tuple(original),
            destination_rule="Latinoamerica -> Miami",
        )

    return replace(
        record,
        ship_to_lines=tuple(original),
        original_ship_to_lines=tuple(original),
        destination_rule="Original",
    )


def extract_unique_records(ci_pdf: bytes) -> list[BolRecord]:
    document = fitz.open(stream=ci_pdf, filetype="pdf")
    records: dict[str, BolRecord] = {}

    for page_index, page in enumerate(document):
        native_word_list = _native_words(page)
        candidate = _extract_record(native_word_list, page, page_index + 1)

        existing = records.get(candidate.air_waybill) if candidate is not None else None
        existing_complete = (
            existing is not None
            and existing.packing_id
            and existing.heart_order
            and existing.ship_to_lines
        )
        if existing_complete:
            continue

        page_is_image_only = len(native_word_list) < 20
        candidate_incomplete = candidate is not None and not (
            candidate.packing_id and candidate.heart_order and candidate.ship_to_lines
        )
        if page_is_image_only or candidate_incomplete:
            ocr_candidate = _extract_record(_ocr_words(page), page, page_index + 1)
            if ocr_candidate is not None:
                candidate = ocr_candidate

        if candidate is None:
            continue

        existing = records.get(candidate.air_waybill)
        if existing is None:
            records[candidate.air_waybill] = candidate
        else:
            records[candidate.air_waybill] = BolRecord(
                air_waybill=existing.air_waybill,
                packing_id=existing.packing_id or candidate.packing_id,
                heart_order=existing.heart_order or candidate.heart_order,
                ship_to_lines=existing.ship_to_lines or candidate.ship_to_lines,
                source_page=existing.source_page,
                source_format=existing.source_format,
            )

    document.close()

    if not records:
        raise BolExtractionError("No se encontro ningun Air Waybill en el PDF.")

    incomplete = [
        record.air_waybill
        for record in records.values()
        if not (record.packing_id and record.heart_order and record.ship_to_lines)
    ]
    if incomplete:
        raise BolExtractionError(
            "No se pudieron leer todos los campos de estos Air Waybill: " + ", ".join(incomplete)
        )

    transformed = [apply_destination_rule(record) for record in records.values()]

    def numeric_sort_key(record: BolRecord):
        value = record.air_waybill.strip()
        if value.isdigit():
            return (0, int(value), value)
        return (1, value, value)

    return sorted(transformed, key=numeric_sort_key)


def _fit_lines(lines: Sequence[str], max_width: float, fontsize: float) -> list[str]:
    fitted: list[str] = []
    for source_line in lines:
        words = source_line.split()
        if not words:
            continue
        current = words[0]
        for word in words[1:]:
            trial = f"{current} {word}"
            if fitz.get_text_length(trial, fontname="helv", fontsize=fontsize) <= max_width:
                current = trial
            else:
                fitted.append(current)
                current = word
        fitted.append(current)
    return fitted


def _format_generation_time(now: datetime | None = None) -> str:
    if now is None:
        now = datetime.now(ZoneInfo(DEFAULT_TIMEZONE))
    elif now.tzinfo is None:
        now = now.replace(tzinfo=ZoneInfo(DEFAULT_TIMEZONE))
    return now.strftime("%b %d, %Y %H:%M:%S")


def _stamp_record(template: fitz.Document, record: BolRecord, date_time: str) -> fitz.Document:
    result = fitz.open()
    result.insert_pdf(template)
    page = result[0]

    font = "helv"
    color = (0, 0, 0)

    page.insert_text((501, 76), record.air_waybill, fontname=font, fontsize=9, color=color)
    page.insert_text((486, 99), date_time, fontname=font, fontsize=8.3, color=color)
    page.insert_text((507, 140), record.heart_order, fontname=font, fontsize=9, color=color)
    page.insert_text((18, 244), record.air_waybill, fontname=font, fontsize=8.5, color=color)
    page.insert_text((237, 244), record.packing_id, fontname=font, fontsize=8.5, color=color)

    # Strictly keep Ship To in the left column and wrap long lines downward.
    ship_fontsize = 8.2
    ship_max_width = 170.0
    max_lines = 7
    ship_lines = _fit_lines(record.ship_to_lines, ship_max_width, ship_fontsize)
    while len(ship_lines) > max_lines and ship_fontsize > 6.2:
        ship_fontsize = round(ship_fontsize - 0.4, 1)
        ship_lines = _fit_lines(record.ship_to_lines, ship_max_width, ship_fontsize)

    ship_lines = ship_lines[:max_lines]
    if ship_lines:
        start_y = 154.0
        bottom_y = 198.0
        line_spacing = min(
            ship_fontsize * 1.18,
            (bottom_y - start_y) / max(1, len(ship_lines) - 1),
        )
        for index, line in enumerate(ship_lines):
            page.insert_text(
                (22, start_y + index * line_spacing),
                line,
                fontname=font,
                fontsize=ship_fontsize,
                color=color,
            )

    return result


def generate_bol_pdf(
    records: Sequence[BolRecord],
    template_path: str | Path,
    now: datetime | None = None,
) -> bytes:
    if not records:
        raise ValueError("Selecciona al menos un BOL.")

    template = fitz.open(str(template_path))
    date_time = _format_generation_time(now)
    output = fitz.open()

    for record in records:
        stamped = _stamp_record(template, record, date_time)
        output.insert_pdf(stamped)
        stamped.close()

    payload = output.tobytes(garbage=4, deflate=True)
    output.close()
    template.close()
    return payload
