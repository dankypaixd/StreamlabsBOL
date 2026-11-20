import contact_review as cr
from contact_extractor import FieldRead, PageRead, PdfRead


def field(value, agreement=1.0):
    """Campo cuyas lecturas dan ese acuerdo (la mezcla de paginas lo recalcula desde las lecturas)."""
    if agreement >= 1.0:
        reads = [(value, 90.0), (value, 90.0)]
    else:
        reads = [(value, 95.0), (value + "~", 90.0)]  # 1 de 2 lecturas coincide
    return FieldRead(value=value, agreement=agreement, reads=reads, weights=[1.0, 1.0], support=agreement)


def page(no, bol="3049901", name="Jane Testwell", email="jane@example.com", phone="1+15551230147",
         state="FL", zip_code="33126", empty=False, **kwargs):
    p = PageRead("VRP_x.pdf", no, "bol")
    p.bol, p.name, p.email, p.phone, p.state = field(bol), field(name), field(email), field(phone), field(state)
    p.zip_code, p.empty_block = zip_code, empty
    for k, v in kwargs.items():
        setattr(p, k, v)
    return p


def build(pages, memory=None, barcodes=None):
    return cr.build_contacts([PdfRead("VRP_x.pdf", pages, len(pages))], memory, barcodes)


def test_green_doubt_red():
    clean = page(1)
    doubtful = page(2, bol="3049902", email="xy@example.com")
    doubtful.email = field("xy@example.com", 0.5)
    no_phone = page(3, bol="3049903", phone="")
    no_phone.phone = FieldRead()
    empty = page(4, bol="3049904", name="", email="", phone="", empty=True)
    for blank in ("name", "email", "phone"):
        setattr(empty, blank, FieldRead())
    got = {c.bol: c for c in build([clean, doubtful, no_phone, empty])}
    assert got["3049901"].status == "ok"
    assert got["3049902"].status == "doubt" and "lowagree:email" in got["3049902"].reasons
    assert got["3049903"].status == "fail" and "missing:phone" in got["3049903"].reasons
    assert got["3049904"].status == "fail" and "empty_block" in got["3049904"].reasons


def test_zip_state_mismatch_and_derived_state_are_never_green():
    wrong = build([page(1, state="CA", zip_code="33126")])[0]  # 331xx es FL
    assert wrong.status == "doubt" and "state_zip_mismatch" in wrong.reasons
    derived = build([page(1, reasons=["state_from_zip"])])[0]
    assert derived.status == "doubt"


def test_memory_fix_is_doubt_and_exact_match_is_green(tmp_path):
    import contact_memory as cm

    memory = cm.ContactMemory(cm.LocalStore(tmp_path / "m.bin"))
    memory.learn([{"name": "Jane Testwell", "email": "jane.testwell@example.com", "phone": "1+15551230147", "state": "FL", "zip_code": "33126"}])

    fixed = build([page(1, email="jane.testweil@example.com")], memory)[0]  # l -> i
    assert fixed.email == "jane.testwell@example.com" and fixed.status == "doubt"
    assert fixed.fixes["email"] == "jane.testweil@example.com"

    weak = page(1, email="jane.testwell@example.com")
    weak.email = field("jane.testwell@example.com", 0.55)  # el OCR dudo, pero es identico a lo confirmado
    weak.name = field("Jane Testwell")
    exact = build([weak], memory)[0]
    assert exact.status == "ok" and "memory_exact" in exact.reasons


def test_two_pages_of_same_bol_are_merged_and_disagreements_flagged():
    contacts = build([page(1), page(2, name="Jane Testwel")])
    assert len(contacts) == 1 and contacts[0].pages == [1, 2]
    assert "pages_disagree:name" in contacts[0].reasons


def test_bol_snaps_to_data_only_when_read_was_doubtful():
    sure = build([page(1, bol="3049111")], barcodes={"3049211"})[0]
    assert sure.bol == "3049111" and "bol_not_in_data" in sure.reasons
    shaky = page(1, bol="3049111")
    shaky.bol = field("3049111", 0.5)
    snapped = build([shaky], barcodes={"3049211"})[0]
    assert snapped.bol == "3049211" and "bol_snapped" in snapped.reasons


def test_ambiguous_i_l_glyph_is_never_green_until_known(tmp_path):
    import contact_memory as cm

    suspicious = build([page(1, email="greg.a.clem@Imco.com")])[0]
    assert suspicious.status == "doubt" and "ambiguous_glyph" in suspicious.reasons

    memory = cm.ContactMemory(cm.LocalStore(tmp_path / "m.bin"))
    memory.learn([{"name": "Jane Testwell", "email": "greg.a.clem@lmco.com", "phone": "1+15551230147", "state": "FL", "zip_code": "33126"}])
    fixed = build([page(1, email="greg.a.clem@Imco.com")], memory)[0]
    assert fixed.email == "greg.a.clem@lmco.com" and fixed.status == "doubt"  # corregido con lo aprendido, a revisar


def test_name_with_digits_or_slash_needs_review():
    odd = build([page(1, name="GLADYS MORASALOMON /4509030732")])[0]
    assert odd.status == "doubt" and "name_odd" in odd.reasons


def test_carrier_is_confirmed_between_pdf_dims_and_file_name():
    assert cr.resolve_carrier("BEX", "BEX", "BEX") == ("BEX", "")
    assert cr.resolve_carrier("", "VRP", "VRP") == ("VRP", "")  # basta con que no se contradigan
    assert cr.resolve_carrier("BEX", "VRP", "VRP") == ("VRP", "carrier_mismatch")  # manda la mayoria
    assert cr.resolve_carrier("BEX", "VRP", "") == ("BEX", "carrier_mismatch")  # empate: manda el PDF
    assert cr.resolve_carrier("", "", "") == ("", "carrier_unknown")


def test_pdf_layout_beats_file_name_and_mismatch_is_flagged_yellow():
    bex_page = page(1, carrier="BEX")  # el BOL dice IEC TECHNOLOGIES, el archivo dice VRP_x.pdf
    got = build([bex_page])[0]
    assert got.carrier == "BEX" and "carrier_mismatch" in got.reasons and got.status == "doubt"

    # con el DIMS de acuerdo con el PDF ya no hay duda
    dims = {"3049901": "BEX"}
    got = cr.build_contacts([PdfRead("x.pdf", [page(1, carrier="BEX")], 1)], None, None, dims)[0]
    assert got.carrier == "BEX" and "carrier_mismatch" not in got.reasons and got.status == "ok"


def test_inbond_classification_does_not_need_contact_fields():
    blank = page(1, name="", email="", phone="")
    for fname in ("name", "email", "phone"):
        setattr(blank, fname, FieldRead())
    got = cr.build_contacts([PdfRead("VRP_x.pdf", [blank], 1)], None, None, None, need_contact=False)[0]
    assert got.status == "ok" and got.state == "FL"
