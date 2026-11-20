"""PDF de BOL sintetico (datos inventados) para probar la lectura sin usar documentos reales."""
from __future__ import annotations

import fitz


def make_bol_pdf(pages: list[dict]) -> bytes:
    """Cada pagina: bol, name, email, phone, city_line, empty (bool). Distribucion tipo VRP."""
    doc = fitz.open()
    for spec in pages:
        page = doc.new_page(width=612, height=792)
        page.insert_text((14, 40), "Bill of Lading", fontname="hebo", fontsize=14)
        page.insert_text((412, 66), "Bill of Lading #:", fontname="hebo", fontsize=9)
        page.insert_text((490, 66), spec["bol"], fontname="helv", fontsize=9)
        page.insert_text((412, 80), "Page:", fontname="hebo", fontsize=9)
        page.insert_text((490, 80), "1 OF 2", fontname="helv", fontsize=9)
        if spec.get("ship_from"):
            page.insert_text((14, 92), "Ship From:", fontname="hebo", fontsize=10)
            for i, line in enumerate(spec["ship_from"]):
                page.insert_text((20, 106 + i * 12), line, fontname="helv", fontsize=8)
        page.insert_text((14, 143), "Ship To:", fontname="hebo", fontsize=10)
        for i, line in enumerate(["ACME TEST LLC", "100 MAIN STREET", spec["city_line"], "US"]):
            page.insert_text((20, 158 + i * 13), line, fontname="helv", fontsize=8)
        page.insert_text((204, 143), "Ship To Attention:", fontname="hebo", fontsize=10)
        if not spec.get("empty"):
            for i, line in enumerate([spec["name"], spec["email"], spec["phone"]]):
                page.insert_text((210, 158 + i * 13), line, fontname="helv", fontsize=8)
        page.draw_line((14, 205), (598, 205))
        page.insert_text((14, 222), "Tracking No", fontname="hebo", fontsize=8)
        page.insert_text((150, 222), "HM", fontname="hebo", fontsize=8)
        page.insert_text((170, 222), "HPE Product", fontname="hebo", fontsize=8)
        page.insert_text((170, 232), "Description", fontname="hebo", fontsize=8)
        page.insert_text((14, 248), spec["bol"], fontname="helv", fontsize=8)
        page.draw_line((14, 255), (598, 255))
    data = doc.tobytes()
    doc.close()
    return data


def make_ci_pdf(awb: str, ship_to: list[str]) -> bytes:
    """Commercial Invoice sintetica tipo VRP (texto nativo): Air Waybill arriba a la derecha y Ship To."""
    doc = fitz.open()
    page = doc.new_page(width=612, height=792)
    page.insert_text((14, 40), "Commercial Invoice", fontname="hebo", fontsize=14)
    page.insert_text((340, 70), "Air Waybill", fontname="hebo", fontsize=9)
    page.insert_text((420, 70), awb, fontname="helv", fontsize=9)
    page.insert_text((340, 84), "Packing ID", fontname="hebo", fontsize=9)
    page.insert_text((420, 84), "PK12345", fontname="helv", fontsize=9)
    page.insert_text((14, 120), "Ship To", fontname="hebo", fontsize=10)
    for i, line in enumerate(ship_to):
        page.insert_text((14, 136 + i * 11), line, fontname="helv", fontsize=8)
    data = doc.tobytes()
    doc.close()
    return data


SAMPLE = [
    {"bol": "3049901", "name": "Jane Testwell", "email": "jane.testwell@example.com",
     "phone": "1-+15551230147", "city_line": "DALLAS, TX 75201"},
    {"bol": "3049902", "name": "Omar Fictional", "email": "omar.f@sample.org",
     "phone": "1-+15559876543", "city_line": "MIAMI, FL 33126"},
    {"bol": "3049903", "name": "", "email": "", "phone": "", "city_line": "HAZLETON, 18202", "empty": True},
]
