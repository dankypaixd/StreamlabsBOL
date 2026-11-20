import io
import shutil

import pytest
from openpyxl import load_workbook

import date_export as de
import inbonds_tab as ib
from tests.synthetic import SAMPLE, make_bol_pdf, make_ci_pdf


def test_state_from_invoice_lines():
    assert ib.state_from_lines(["ACME", "100 MAIN ST", "DALLAS, TX 75201", "US"]) == ("TX", "75201", "code")
    assert ib.state_from_lines(["ACME", "HAZLETON, 18202", "US"]) == ("PA", "18202", "zip")  # sin codigo de estado
    assert ib.state_from_lines(["ACME", "SAN JUAN", "Puerto Rico"])[0] == "PR"
    assert ib.state_from_lines(["ACME", "MONTERREY NL", "Mexico"]) == ("", "", "")
    # Canada: provincia + codigo postal (antes no se reconocia y el BOL quedaba sin estado)
    assert ib.state_from_lines(["TD SYNNEX CANADA ULC", "6901 Creditview Road", "Mississauga, ON  L5N 8E9", "CA"])[0] == "ON"
    assert ib.sl.service_label("ON") == "4D"


def test_pdf_kind_tells_invoice_from_scanned_bol():
    assert ib.pdf_kind(make_ci_pdf("3049901", ["ACME", "DALLAS, TX 75201", "US"])) == "ci"
    assert ib.pdf_kind(make_bol_pdf(SAMPLE[:1])) == "bol"  # el BOL sintetico no trae "air waybill"


def test_commercial_invoice_gives_bol_state_and_service():
    contacts, kinds, errors = ib.read_files(
        [("VRP_ci.pdf", make_ci_pdf("3049901", ["ACME TEST LLC", "100 MAIN STREET", "DALLAS, TX 75201", "US"]))]
    )
    assert kinds == {"VRP_ci.pdf": "ci"} and errors == {}
    (contact,) = contacts
    assert (contact.bol, contact.state, contact.carrier, contact.status) == ("3049901", "TX", "VRP", "ok")
    assert contact.name == contact.email == contact.phone == ""  # no lee contacto


@pytest.mark.skipif(shutil.which("tesseract") is None, reason="Tesseract no instalado")
def test_scanned_bols_are_read_without_contact():
    contacts, kinds, _ = ib.read_files([("BEX_x.pdf", make_bol_pdf(SAMPLE))])
    assert kinds == {"BEX_x.pdf": "bol"}
    by_bol = {c.bol: c for c in contacts}
    assert by_bol["3049901"].state == "TX" and by_bol["3049901"].carrier == "BEX"
    assert by_bol["3049901"].status == "ok" and by_bol["3049901"].name == ""
    assert by_bol["3049903"].state == "PA" and by_bol["3049903"].status == "doubt"  # estado sacado del ZIP


def test_inbond_export_has_inbd_in_priority_code():
    rows = [de.ShipmentRow("3049901", state="TX", carrier="VRP"), de.ShipmentRow("98162916", state="FL", carrier="BEX")]
    blob, _ = de.build_workbook(rows, "11/19/26 1356", None, include_data=False, priority=de.PRIORITY_INBOND)
    sheet = load_workbook(io.BytesIO(blob))["DATE"]
    assert [sheet["FK2"].value, sheet["FK3"].value] == ["INBD", "INBD"]
    assert (sheet["E2"].value, sheet["E3"].value) == ("144463", "149570")
    assert sheet["AK2"].value == "3D" and sheet["AB2"].value is None  # sin contacto


CANADA = ["TD SYNNEX CANADA ULC TDCATORRCV@tdsynnex.com", "6901 Creditview Road", "", "Mississauga, ON  L5N 8E9", "CA"]


def test_invoice_to_canada_gets_its_province():
    (contact,), _, _ = ib.read_files([("VRP_ci.pdf", make_ci_pdf("3049251", CANADA))])
    assert (contact.bol, contact.state, contact.status) == ("3049251", "ON", "ok")


def test_invoice_capture_is_only_the_ship_to_block():
    import fitz
    from PIL import Image

    content = make_ci_pdf("3049251", CANADA)
    page = fitz.open(stream=content, filetype="pdf")[0]
    scale = 150 / 72
    words = [(w[4], w[0] * scale, w[1] * scale, w[2] * scale, w[3] * scale) for w in page.get_text("words")]
    x0, y0, x1, y1 = ib.ship_to_box(words, 612 * scale, 792 * 0.55 * scale)
    assert y0 > 112 * scale and y1 < 250 * scale  # debajo del Ship From y arriba de HEART ORDER NO
    assert x1 < 340 * scale  # sin la columna Sold To / Ship Via
    image = Image.open(io.BytesIO(ib.ci_ship_to_capture(content, 1)))
    assert image.height < 0.4 * 792 * 0.55 * scale and image.width < 0.6 * 612 * scale


@pytest.mark.skipif(shutil.which("tesseract") is None, reason="Tesseract no instalado")
def test_scanned_invoice_is_read_and_cropped_with_ocr():
    from PIL import Image

    scan = make_ci_pdf("3049251", CANADA, image_only=True)
    assert ib.pdf_kind(scan) == "bol"  # sin texto no se puede saber a primera vista...
    (contact,), kinds, _ = ib.read_files([("VRP_scan.pdf", scan)])  # ...pero el lector cae a Commercial Invoice
    assert kinds == {"VRP_scan.pdf": "ci"} and (contact.bol, contact.state) == ("3049251", "ON")
    image = Image.open(io.BytesIO(ib.ci_ship_to_capture(scan, 1)))
    assert image.width < 0.6 * 612 * 150 / 72


def test_invoice_to_mexico_or_latam_uses_the_delivery_state():
    mexico = ["CORP SA DE CV", "CALLE 5 No 10", "MONTERREY, NL 66260", "MX"]  # NL no es Newfoundland aqui
    brazil = ["EMPRESA LTDA", "RUA X 10", "SAO PAULO SP 01310-100", "BR"]
    panama = ["ACME TEST LLC", "100 MAIN STREET", "Panama City, FL 32401", "US"]
    contacts, _, errors = ib.read_files([
        ("VRP_mx.pdf", make_ci_pdf("3049301", mexico)),
        ("VRP_br.pdf", make_ci_pdf("3049302", brazil)),
        ("VRP_us.pdf", make_ci_pdf("3049303", panama)),
    ])
    assert errors == {}
    by_bol = {c.bol: c for c in contacts}
    assert (by_bol["3049301"].state, by_bol["3049301"].status) == ("TX", "ok")
    assert "destination_mexico" in by_bol["3049301"].reasons
    assert (by_bol["3049302"].state, by_bol["3049302"].status) == ("FL", "ok")
    assert "destination_latam" in by_bol["3049302"].reasons
    assert by_bol["3049303"].state == "FL" and not any(r.startswith("destination_") for r in by_bol["3049303"].reasons)


def _inbonds_app():
    import streamlit as st

    import contacts_tab
    import inbonds_tab

    if "language" not in st.session_state:
        st.session_state.language = "es"

    def t(key, **values):
        lang = st.session_state.language
        text = {**contacts_tab.TEXT[lang], **inbonds_tab.TEXT[lang]}[key]
        return text.format(**values) if values else text

    inbonds_tab.render(t)


@pytest.mark.skipif(shutil.which("tesseract") is None, reason="Tesseract no instalado")
def test_inbonds_has_arrived_and_left_plant_messages_and_push_warning():
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_function(_inbonds_app, default_timeout=60)
    at.run()
    assert not at.exception
    pod = next(w for w in at.text_input if w.label.startswith("POD DATE TIME"))
    pod.set_value("11/20/26 0800").run()
    assert [c.value for c in at.code] == ["ARRIVED AT 11/20/26 0800"]
    ready = next(w for w in at.text_input if w.label.startswith("READY DATE"))
    ready.set_value("11/20/26 0800").run()
    assert [c.value for c in at.code] == [
        "LEFT PLANT 11/20/26 0800 ARRIVED AT 11/20/26 0800", "LEFT PLANT 11/20/26 0800", "ARRIVED AT 11/20/26 0800",
    ]
    assert not any("push" in w.value for w in at.warning)
    ready.set_value("11/20/26 1730").run()
    assert any("17:30" in w.value and "push" in w.value for w in at.warning)


def test_inbonds_export_has_consignee_number_for_laredo_and_miami():
    cases = {
        "3049401": ["CORP SA DE CV", "CALLE 5 No 10", "MONTERREY, NL 66260", "MX"],
        "3049402": ["EMPRESA LTDA", "RUA X 10", "SAO PAULO SP 01310-100", "BR"],
        "3049403": ["Corporativo Galvan S.C", "11905 Conly Road", "Laredo TX 78045", "USA"],
        "3049404": ["ACME TEST LLC", "100 MAIN STREET", "Houston, TX 77001", "US"],
    }
    contacts, _, errors = ib.read_files([(f"VRP_{b}.pdf", make_ci_pdf(b, lines)) for b, lines in cases.items()])
    assert errors == {}
    by_bol = {c.bol: c for c in contacts}
    assert [by_bol[b].consignee for b in cases] == ["147229", "210475", "147229", ""]
    rows = [de.ShipmentRow(b, state=by_bol[b].state, carrier="VRP", consignee=by_bol[b].consignee) for b in cases]
    blob, _ = de.build_workbook(rows, "11/19/26 1356", None, include_data=False, priority=de.PRIORITY_INBOND)
    sheet = load_workbook(io.BytesIO(blob))["DATE"]
    assert [sheet[f"T{n}"].value for n in (2, 3, 4, 5)] == ["147229", "210475", "147229", None]
