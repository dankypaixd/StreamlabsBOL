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
    # Un Ship To Attention vacio es normal: verde, con nombre, correo y telefono en blanco
    assert got["3049904"].status == "ok" and "empty_block" in got["3049904"].reasons
    assert (got["3049904"].name, got["3049904"].email, got["3049904"].phone) == ("", "", "")
    assert not [r for r in got["3049904"].reasons if r.startswith(("missing", "lowagree"))]


def _empty_page(no=4, **kwargs):
    blank = page(no, bol="3049904", name="", email="", phone="", empty=True, **kwargs)
    for fname in ("name", "email", "phone"):
        setattr(blank, fname, FieldRead())
    return blank


def test_empty_block_is_green_but_the_state_is_still_required_and_checked():
    no_state = _empty_page(state="")
    no_state.state = FieldRead()
    got = build([no_state])[0]
    assert got.status == "fail" and "missing:state" in got.reasons  # el service level sale del estado
    guessed = build([_empty_page(reasons=["state_from_zip"])])[0]
    assert guessed.status == "doubt" and "state_from_zip" in guessed.reasons  # un estado deducido sigue en amarillo


def test_empty_block_survives_a_manual_review_only_while_it_stays_empty():
    contact = build([_empty_page()])[0]
    assert contact.status == "ok"
    again = cr.reclassify(contact, None)  # se abre y se guarda sin escribir nada
    assert again.status == "ok" and "empty_block" in again.reasons

    again.name = "Jane Testwell"  # ya hay un dato: deja de ser un bloque vacio y falta el resto
    partial = cr.reclassify(again, None)
    assert partial.status == "fail" and "empty_block" not in partial.reasons
    assert "missing:email" in partial.reasons and "missing:phone" in partial.reasons


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


def test_carrier_is_confirmed_between_pdf_and_file_name():
    assert cr.resolve_carrier("BEX", "BEX") == ("BEX", "")
    assert cr.resolve_carrier("", "VRP") == ("VRP", "")  # basta con que no se contradigan
    assert cr.resolve_carrier("BEX", "VRP") == ("BEX", "carrier_mismatch")  # en desacuerdo manda el PDF
    assert cr.resolve_carrier("", "") == ("", "carrier_unknown")


def test_pdf_layout_beats_file_name_and_mismatch_is_flagged_yellow():
    bex_page = page(1, carrier="BEX")  # el BOL dice IEC TECHNOLOGIES, el archivo dice VRP_x.pdf
    got = build([bex_page])[0]
    assert got.carrier == "BEX" and "carrier_mismatch" in got.reasons and got.status == "doubt"

    # y si el archivo dice lo mismo que el BOL no hay duda
    got = cr.build_contacts([PdfRead("BEX_x.pdf", [page(1, carrier="BEX")], 1)])[0]
    assert got.carrier == "BEX" and "carrier_mismatch" not in got.reasons and got.status == "ok"


def test_inbond_classification_does_not_need_contact_fields():
    blank = page(1, name="", email="", phone="")
    for fname in ("name", "email", "phone"):
        setattr(blank, fname, FieldRead())
    got = cr.build_contacts([PdfRead("VRP_x.pdf", [blank], 1)], None, None, need_contact=False)[0]
    assert got.status == "ok" and got.state == "FL"


def _job(*bols, confirmed=()):
    contacts = {}
    for index, bol in enumerate(bols):
        contact = cr.Contact(key=f"k{index}", bol=bol, source_file="VRP_x.pdf", pages=[index + 1])
        contact.confirmed = contact.key in confirmed
        contacts[contact.key] = contact
    return contacts


def test_fixing_a_bol_to_one_already_in_the_list_leaves_a_single_record():
    # "98463300" era una mala lectura de "98163300": al corregirlo ya no queda como un BOL mas
    contacts = _job("98163300", "98463300", "98163301")
    edited = contacts["k1"]
    edited.bol = "98163300"
    assert cr.merge_duplicate_bol(contacts, edited) == ["98163300"]
    assert sorted(c.bol for c in contacts.values()) == ["98163300", "98163301"]
    assert contacts["k1"].pages == [1, 2]  # se queda el corregido y suma las paginas del que sobraba
    assert "k0" not in contacts


def test_the_already_confirmed_record_is_the_one_that_stays():
    contacts = _job("98163300", "98463300", confirmed=("k0",))
    edited = contacts["k1"]
    edited.bol = "98163300"
    cr.merge_duplicate_bol(contacts, edited)
    assert list(contacts) == ["k0"] and contacts["k0"].pages == [1, 2]


def test_no_duplicates_means_nothing_changes():
    contacts = _job("98163300", "98163301")
    edited = contacts["k1"]
    assert cr.merge_duplicate_bol(contacts, edited) == [] and len(contacts) == 2
    edited.bol = ""  # sin BOL no se junta con nadie
    assert cr.merge_duplicate_bol(contacts, edited) == [] and len(contacts) == 2
