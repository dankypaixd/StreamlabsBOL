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

    # Bloque vacio y estado deducido del ZIP (no estaba escrito): nunca queda en verde
    empty = contacts["3049903"]
    assert empty.status == "fail" and "empty_block" in empty.reasons
    assert empty.state == "PA" and "state_from_zip" in empty.reasons
