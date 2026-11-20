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
