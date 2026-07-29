from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import re
from typing import Iterable
from zoneinfo import ZoneInfo

import fitz  # PyMuPDF
from PIL import Image
import pytesseract


DEFAULT_TIMEZONE = "America/Costa_Rica"


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
    order_item: str
    ship_to_lines: tuple[str, ...]
    source_page: int


def _clean_text(value: str) -> str:
    return (
        value.replace("\u00a0", " ")
        .replace("\u00ad", "")
        .replace("\ufeff", "")
        .strip()
    )


def _native_words(page: fitz.Page) -> list[Word]:
    words: list[Word] = []
    for item in page.get_text("words", sort=True):
        x0, y0, x1, y1, text, *_ = item
        text = _clean_text(str(text))
        if text:
            words.append(Word(float(x0), float(y0), float(x1), float(y1), text))
    return words


def _ocr_words(page: fitz.Page, dpi: int = 200) -> list[Word]:
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
        text = _clean_text(raw_text)
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


def _find_value_right_of_label(
    words: list[Word],
    page_width: float,
    page_height: float,
    label_tokens: tuple[str, ...],
) -> str | None:
    candidates = [
        word
        for word in words
        if word.x0 >= page_width * 0.53 and word.y0 <= page_height * 0.32
    ]

    for line in _group_lines(candidates):
        normalized = [word.text.lower().strip("#:.") for word in line]
        for start in range(0, len(normalized) - len(label_tokens) + 1):
            if tuple(normalized[start : start + len(label_tokens)]) != label_tokens:
                continue
            label_end = line[start + len(label_tokens) - 1].x1
            values = [word.text for word in line if word.x0 > label_end + 12]
            if values:
                return " ".join(values)
    return None


def _extract_ship_to(words: list[Word], width: float, height: float) -> tuple[str, ...]:
    label_y: float | None = None
    for line in _group_lines(
        [word for word in words if word.x0 < width * 0.22 and word.y0 < height * 0.42]
    ):
        tokens = [word.text.upper().strip(":") for word in line]
        if "SHIP" in tokens and "TO" in tokens:
            label_y = min(word.y0 for word in line)
            break

    if label_y is None:
        return ()

    destination_words = [
        word
        for word in words
        if word.x0 < width * 0.34
        and label_y + 7 <= word.y0 < height * 0.39
    ]

    lines: list[str] = []
    for line_words in _group_lines(destination_words, tolerance=3.2):
        line = " ".join(word.text for word in line_words)
        line = re.sub(r"\s+", " ", line).strip()
        if line:
            lines.append(line)
    return tuple(lines)


def _extract_order_item(words: list[Word], width: float, height: float) -> str | None:
    candidates: list[tuple[float, str]] = []
    for word in words:
        digits = re.sub(r"\D", "", word.text)
        if not (8 <= len(digits) <= 12):
            continue
        if digits.startswith("0000"):
            continue
        if width * 0.07 <= word.x0 <= width * 0.22 and height * 0.43 <= word.y0 <= height * 0.90:
            candidates.append((word.y0, digits))

    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def _extract_record(words: list[Word], page: fitz.Page, source_page: int) -> BolRecord | None:
    width = page.rect.width
    height = page.rect.height

    air_raw = _find_value_right_of_label(words, width, height, ("air", "waybill"))
    air_waybill = re.sub(r"\D", "", air_raw or "")

    packing_raw = _find_value_right_of_label(words, width, height, ("packing", "id"))
    packing_id = (packing_raw or "").split()[0].strip("#:.,")

    ship_to_lines = _extract_ship_to(words, width, height)
    order_item = _extract_order_item(words, width, height) or ""

    if not air_waybill:
        return None
    if not (packing_id and order_item and ship_to_lines):
        return BolRecord(
            air_waybill=air_waybill,
            packing_id=packing_id,
            order_item=order_item,
            ship_to_lines=ship_to_lines,
            source_page=source_page,
        )

    return BolRecord(
        air_waybill=air_waybill,
        packing_id=packing_id,
        order_item=order_item,
        ship_to_lines=ship_to_lines,
        source_page=source_page,
    )


def extract_unique_records(ci_pdf: bytes) -> list[BolRecord]:
    document = fitz.open(stream=ci_pdf, filetype="pdf")
    records: dict[str, BolRecord] = {}

    for page_index, page in enumerate(document):
        native_word_list = _native_words(page)
        native = _extract_record(native_word_list, page, page_index + 1)
        candidate = native

        # Repeated pages for an Air Waybill already completed do not need OCR.
        existing_native = records.get(candidate.air_waybill) if candidate is not None else None
        existing_is_complete = (
            existing_native is not None
            and existing_native.packing_id
            and existing_native.order_item
            and existing_native.ship_to_lines
        )
        if existing_is_complete:
            continue

        # OCR is only needed for image-only pages or when the first detected CI page is incomplete.
        page_is_image_only = len(native_word_list) < 20
        detected_ci_is_incomplete = (
            candidate is not None
            and (not candidate.packing_id or not candidate.order_item or not candidate.ship_to_lines)
        )
        if page_is_image_only or detected_ci_is_incomplete:
            ocr_candidate = _extract_record(_ocr_words(page), page, page_index + 1)
            if ocr_candidate is not None:
                candidate = ocr_candidate

        if candidate is None:
            continue

        existing = records.get(candidate.air_waybill)
        if existing is None:
            records[candidate.air_waybill] = candidate
            continue

        # A later page may complete a field missing on the first page.
        records[candidate.air_waybill] = BolRecord(
            air_waybill=existing.air_waybill,
            packing_id=existing.packing_id or candidate.packing_id,
            order_item=existing.order_item or candidate.order_item,
            ship_to_lines=existing.ship_to_lines or candidate.ship_to_lines,
            source_page=existing.source_page,
        )

    incomplete = [
        record.air_waybill
        for record in records.values()
        if not (record.packing_id and record.order_item and record.ship_to_lines)
    ]
    if incomplete:
        joined = ", ".join(incomplete)
        raise ValueError(f"No se pudieron leer todos los campos de estos Air Waybill: {joined}")

    if not records:
        raise ValueError("No se encontró ningún Air Waybill en el PDF.")

    return list(records.values())


def _fit_lines(lines: tuple[str, ...], max_width: float, fontsize: float) -> list[str]:
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
    page.insert_text((507, 140), record.order_item, fontname=font, fontsize=9, color=color)
    page.insert_text((18, 244), record.air_waybill, fontname=font, fontsize=8.5, color=color)
    page.insert_text((237, 244), record.packing_id, fontname=font, fontsize=8.5, color=color)

    ship_fontsize = 8.2
    ship_lines = _fit_lines(record.ship_to_lines, max_width=322, fontsize=ship_fontsize)
    if len(ship_lines) > 5:
        ship_fontsize = 7.4
        ship_lines = _fit_lines(record.ship_to_lines, max_width=322, fontsize=ship_fontsize)
    ship_lines = ship_lines[:6]
    if ship_lines:
        available_height = 43.0
        line_spacing = min(ship_fontsize * 1.35, available_height / max(1, len(ship_lines) - 1))
        start_y = 154.0
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
    ci_pdf: bytes,
    template_path: str | Path,
    now: datetime | None = None,
) -> tuple[bytes, list[BolRecord]]:
    records = extract_unique_records(ci_pdf)
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
    return payload, records
