"""Regla de destino de la pestana 1 (Mexico -> Laredo TX, Latinoamerica -> Miami FL) en las pestanas 3 y 4."""
import shutil

import pytest

import destinations as d
from tests.synthetic import SAMPLE, make_bol_pdf


@pytest.mark.parametrize(
    "lines, rule, state",
    [
        (["ACME SA", "CALLE 5 No 10", "MONTERREY, NL 66260", "MX"], "mexico", "TX"),  # NL = Newfoundland si se leyera mal
        (["ACME SA", "AV REFORMA 100", "CIUDAD DE MEXICO 06600", "MEXICO"], "mexico", "TX"),
        (["EMPRESA LTDA", "RUA X 10", "SAO PAULO SP 01310-100", "BR"], "latam", "FL"),
        (["EMPRESA", "CALLE 1", "BOGOTA 110111", "CO"], "latam", "FL"),
        (["ACME", "1 MAIN ST", "Panama City, FL 32401", "US"], "", ""),  # ciudad con nombre de pais
        (["ACME", "1 MAIN ST", "Panama City, FL 32401"], "", ""),
        (["ACME", "1 MAIN ST", "Mexico, MO 65265", "US"], "", ""),
        (["ACME", "1 MAIN ST", "Mexico, MO 65265"], "", ""),
        (["ACME", "1 CALLE", "SAN JUAN, PR 00901", "PR"], "", ""),  # Puerto Rico no cambia
        (["ACME", "1 CALLE", "SAN JUAN 00901", "PUERTO RICO"], "", ""),
        (["ACME", "1 MAIN ST", "Mississauga, ON L5N 8E9", "CA"], "", ""),  # Canada no cambia
        (["ACME", "1 MAIN ST", "Houston, TX 77001", "US"], "", ""),
        ([], "", ""),
    ],
)
def test_destination_rule(lines, rule, state):
    assert d.destination_state(lines) == (rule, state)


def test_states_come_from_tab_1_addresses():
    import bol_generator as bg

    assert d.DESTINATION_STATE == {"mexico": "TX", "latam": "FL"}
    assert "Laredo TX" in " ".join(bg.MEXICO_ADDRESS) and "FL" in " ".join(bg.LATAM_ADDRESS)


@pytest.mark.skipif(shutil.which("tesseract") is None, reason="Tesseract no instalado")
def test_bol_scan_applies_destination_rule():
    import contact_extractor as ce
    import contact_review as cr

    pages = [
        {**SAMPLE[0], "bol": "3049911", "city_line": "MONTERREY, NL 66260", "country": "MX"},
        {**SAMPLE[1], "bol": "3049912", "city_line": "SAO PAULO SP 01310-100", "country": "BR"},
        {**SAMPLE[2], "bol": "3049913", "city_line": "Panama City, FL 32401", "country": "US"},
    ]
    pdf = ce.read_pdf(make_bol_pdf(pages), "VRP_dest.pdf", workers=1)
    contacts = {c.bol: c for c in cr.build_contacts([pdf])}

    mexico = contacts["3049911"]
    assert mexico.state == "TX" and "destination_mexico" in mexico.reasons
    brazil = contacts["3049912"]
    assert brazil.state == "FL" and "destination_latam" in brazil.reasons
    usa = contacts["3049913"]
    assert usa.state == "FL" and not any(r.startswith("destination_") for r in usa.reasons)
    # El aviso de destino es informativo: no pone la fila en amarillo por si solo.
    assert not {"state_fixed", "state_from_zip", "state_from_name", "destination_unsure"} & set(mexico.reasons)
    assert "lowagree:state" not in mexico.reasons


def test_reclassify_drops_destination_note_when_state_is_edited():
    import contact_review as cr

    contact = cr.Contact(key="1", bol="1", source_file="x.pdf", pages=[1], state="TX",
                         agreement={"state": 1.0, "bol": 1.0}, reasons=["destination_mexico"])
    cr.reclassify(contact, None, need_contact=False)
    assert "destination_mexico" in contact.reasons
    contact.state = "CA"
    cr.reclassify(contact, None, need_contact=False)
    assert "destination_mexico" not in contact.reasons


@pytest.mark.parametrize(
    "lines, number",
    [
        (["ACME SA", "CALLE 5 No 10", "MONTERREY, NL 66260", "MX"], "147229"),  # Mexico: se convierte a Laredo
        (["EMPRESA LTDA", "RUA X 10", "SAO PAULO SP 01310-100", "BR"], "210475"),  # Latam: se convierte a Miami
        (["Corporativo Galvan S.C", "11905 Conly Road", "Laredo TX 78045", "USA"], "147229"),  # ya es la de Laredo
        (["ACME", "11905 CONLY RD", "LAREDO, TX 78045", "US"], "147229"),
        (["US_MIAMI DSV Inc. - MIA", "12430 NW 25th Street, Suite 100", "Miami, FL 33182", "US"], "210475"),  # ya es la de Miami
        (["ACME", "12430 NW 25TH ST", "MIAMI FL 33182", "US"], "210475"),
        (["ACME", "100 MAIN STREET", "Houston, TX 77001", "US"], ""),  # el resto queda vacio
        (["ACME", "1 MAIN ST", "Panama City, FL 32401", "US"], ""),
        (["ACME", "1 MAIN ST", "Mexico, MO 65265", "US"], ""),
        (["ACME", "1 CALLE", "SAN JUAN, PR 00901", "PR"], ""),
        (["ACME", "1 MAIN ST", "Mississauga, ON L5N 8E9", "CA"], ""),
        ([], ""),
    ],
)
def test_consignee_number(lines, number):
    assert d.consignee_number(lines) == number


@pytest.mark.skipif(shutil.which("tesseract") is None, reason="Tesseract no instalado")
def test_bol_scan_fills_consignee_number():
    import contact_extractor as ce
    import contact_review as cr

    pages = [
        {**SAMPLE[0], "bol": "3049921", "city_line": "MONTERREY, NL 66260", "country": "MX"},
        {**SAMPLE[1], "bol": "3049922", "city_line": "SAO PAULO SP 01310-100", "country": "BR"},
        {**SAMPLE[0], "bol": "3049923", "address": "11905 CONLY ROAD", "city_line": "Laredo TX 78045", "country": "USA"},
        {**SAMPLE[1], "bol": "3049924", "address": "12430 NW 25TH STREET", "city_line": "Miami, FL 33182", "country": "US"},
        {**SAMPLE[0], "bol": "3049925", "city_line": "Houston, TX 77001", "country": "US"},
    ]
    pdf = ce.read_pdf(make_bol_pdf(pages), "VRP_cn.pdf", workers=1)
    contacts = {c.bol: c for c in cr.build_contacts([pdf])}
    assert [contacts[b].consignee for b in ("3049921", "3049922", "3049923", "3049924", "3049925")] == [
        "147229", "210475", "147229", "210475", "",
    ]
    assert contacts["3049923"].state == "TX" and contacts["3049924"].state == "FL"
