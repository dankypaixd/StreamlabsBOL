import shutil

import pytest

from tests.synthetic import SAMPLE, make_bol_pdf

pytestmark = pytest.mark.skipif(shutil.which("tesseract") is None, reason="Tesseract no instalado")


def test_reads_synthetic_bols_end_to_end():
    import contact_extractor as ce
    import contact_review as cr

    pdf = ce.read_pdf(make_bol_pdf(SAMPLE), "VRP_test.pdf", workers=1)
    assert pdf.skipped == []
    contacts = {c.bol: c for c in cr.build_contacts([pdf])}

    jane = contacts["3049901"]
    assert (jane.name, jane.email, jane.state) == ("Jane Testwell", "jane.testwell@example.com", "TX")
    assert jane.phone.endswith("15551230147")
    assert jane.status == "ok"

    omar = contacts["3049902"]
    assert omar.email == "omar.f@sample.org" and omar.state == "FL"

    # Bloque vacio: ya no es rojo, pero con el estado deducido del ZIP (no estaba escrito) nunca queda en verde
    empty = contacts["3049903"]
    assert empty.status == "doubt" and "empty_block" in empty.reasons
    assert (empty.name, empty.email, empty.phone) == ("", "", "")
    assert empty.state == "PA" and "state_from_zip" in empty.reasons


def test_carrier_is_read_from_the_ship_from_block():
    import contact_extractor as ce

    bex = [{**SAMPLE[0], "ship_from": ["IEC TECHNOLOGIES S. DE R.L. DE C.V.", "3000 HANOVER ST"]}]
    vrp = [{**SAMPLE[1], "ship_from": ["Hewlett Packard Enterprise c/o Foxconn", "Houston TX"]}]
    nada = [SAMPLE[2]]
    assert [p.carrier for p in ce.read_pdf(make_bol_pdf(bex), "x.pdf", workers=1).pages] == ["BEX"]
    assert [p.carrier for p in ce.read_pdf(make_bol_pdf(vrp), "x.pdf", workers=1).pages] == ["VRP"]
    assert [p.carrier for p in ce.read_pdf(make_bol_pdf(nada), "x.pdf", workers=1).pages] == [""]


def test_light_mode_skips_the_contact_block():
    import contact_extractor as ce

    page = ce.read_pdf(make_bol_pdf(SAMPLE[:1]), "x.pdf", workers=1, contact=False).pages[0]
    assert page.bol.value == "3049901" and page.state.value == "TX"
    assert page.name.value == page.email.value == page.phone.value == ""


def test_email_cut_by_the_line_width_is_put_back_together():
    """BEX parte la cadena nombre/correo/telefono por ancho: el .COM sale como ".CO" + "M" y un dominio, a media palabra."""
    import contact_extractor as ce

    pdf = make_bol_pdf([
        {"bol": "98100001", "name": "", "email": "", "phone": "", "city_line": "MADISON, WI 53703",
         "lines": ["JANE TESTWELL /", "R0123/JTESTWELL@SAMPLEMMC.CO", "M/1-+15551230147"]},
        {"bol": "98100002", "name": "", "email": "", "phone": "", "city_line": "COLUMBUS, OH 43004",
         "lines": ["OMAR FICTIONAL", "PO#AB-LL-006-1/omar.fictional@sam", "plegroup.com/1-+15559876543"]},
    ])
    first, second = ce.read_pdf(pdf, "BEX_x.pdf", workers=1).pages
    assert first.email.value == "JTESTWELL@SAMPLEMMC.COM" and first.phone.value == "1+15551230147"
    assert second.email.value == "omar.fictional@samplegroup.com"  # sin un punto inventado entre "sam" y "plegroup"
    assert second.phone.value == "1+15559876543"
