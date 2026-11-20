"""PDF de BOL sintetico (datos inventados) para probar la lectura sin usar documentos reales."""
from __future__ import annotations

import fitz


def make_bol_pdf(pages: list[dict]) -> bytes:
    """Cada pagina: bol, name, email, phone, city_line, empty (bool), country (por defecto "US"), address (por defecto "100 MAIN STREET"). Distribucion tipo VRP."""
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
        for i, line in enumerate(["ACME TEST LLC", spec.get("address", "100 MAIN STREET"), spec["city_line"], spec.get("country", "US")]):
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


def make_ci_pdf(awb: str, ship_to: list[str], image_only: bool = False) -> bytes:
    """Commercial Invoice sintetica tipo VRP: Ship From y Ship To a la izquierda, Sold To / Ship Via a la derecha."""
    doc = fitz.open()
    page = doc.new_page(width=612, height=792)
    page.insert_text((14, 40), "Commercial Invoice", fontname="hebo", fontsize=14)
    page.insert_text((340, 70), "Air Waybill", fontname="hebo", fontsize=9)
    page.insert_text((420, 70), awb, fontname="helv", fontsize=9)
    page.insert_text((340, 84), "Packing ID", fontname="hebo", fontsize=9)
    page.insert_text((420, 84), "PK12345", fontname="helv", fontsize=9)
    page.insert_text((14, 100), "SHIP FROM", fontname="hebo", fontsize=10)
    page.insert_text((20, 114), "Hewlett Packard Enterprise C/O eCMMS", fontname="helv", fontsize=8)
    page.insert_text((340, 100), "SOLD TO", fontname="hebo", fontsize=10)
    page.insert_text((344, 114), "Hewlett Packard Enterprise", fontname="helv", fontsize=8)
    page.insert_text((14, 146), "SHIP TO", fontname="hebo", fontsize=10)
    for i, line in enumerate(ship_to):
        page.insert_text((20, 160 + i * 11), line, fontname="helv", fontsize=8)
    page.insert_text((340, 170), "SHIP VIA", fontname="hebo", fontsize=10)
    page.insert_text((344, 184), "Pegasus Logistics Group", fontname="helv", fontsize=8)
    page.insert_text((60, 250), "HEART ORDER NO", fontname="hebo", fontsize=8)
    if image_only:  # como un escaneo: sin capa de texto
        pix = page.get_pixmap(dpi=200)
        scan = fitz.open()
        scan_page = scan.new_page(width=612, height=792)
        scan_page.insert_image(scan_page.rect, pixmap=pix)
        data = scan.tobytes()
        scan.close()
        doc.close()
        return data
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
