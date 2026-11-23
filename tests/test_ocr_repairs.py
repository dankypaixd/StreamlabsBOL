"""Reparaciones de texto del OCR: correos y nombres escritos en varias lineas, BOL leidos con digitos dudosos.

Todo con datos sinteticos (nada de BOL reales).
"""
import contact_extractor as ce
from contact_extractor import FieldRead, OcrWord


def mend(text):
    return ce.mend_email_breaks(text)[0]


def test_mend_joins_a_com_that_fell_to_the_next_line():
    assert mend("Jane Testwell\njane.testwell@example\n.com\n1+15551230147") == "Jane Testwell\njane.testwell@example.com\n1+15551230147"
    assert mend("jane@example\ncom") == "jane@example.com"  # sin el punto: lo agrega


def test_mend_joins_a_break_right_after_the_at_sign_and_an_at_sign_that_starts_a_line():
    assert mend("jane.testwell@\nexample.com") == "jane.testwell@example.com"
    assert mend("jane.testwell\n@example.com") == "jane.testwell@example.com"


def test_mend_repairs_a_dot_read_as_space_or_colon_and_ignores_scan_specks():
    assert mend("jane@example : com") == "jane@example.com"
    assert mend("jane@example com") == "jane@example.com"
    assert mend("jane@example |\n.com") == "jane@example.com"  # mancha del escaneo entre las dos lineas


def test_mend_leaves_a_valid_email_and_the_next_line_alone():
    text = "Jane Testwell\njane@example.com\n1+15551230147"
    assert mend(text) == text
    # sin correo no hay nada que unir
    assert mend("Jane Testwell\n1+15551230147") == "Jane Testwell\n1+15551230147"


def test_mend_does_not_swallow_the_phone_when_the_domain_has_no_ending():
    text = "jane@example\n1+15551230147"
    assert mend(text) == text


def test_mend_keeps_the_character_boxes_aligned_with_the_new_text():
    chars = [ce.Char(c, float(i), 0.0, float(i) + 1, 10.0, 0) for i, c in enumerate("jane@example\n.com")]
    text, mapping = ce.mend_email_breaks("jane@example\n.com", chars)
    assert text == "jane@example.com"
    assert len(mapping) == len(text)
    assert [m.ch for m in mapping if m] == list(text)


def test_tidy_name_joins_lines_and_drops_symbol_specks():
    assert ce.tidy_name("Jane\nTestwell") == "Jane Testwell"
    assert ce.tidy_name("Jane | Testwell -") == "Jane Testwell"
    assert ce.tidy_name("  Maria   Del  Carmen ") == "Maria Del Carmen"
    assert ce.tidy_name("") == ""


def _word(text, conf, line=0, x=0.0):
    return OcrWord(text, x, line * 20.0, x + 10, line * 20.0 + 10, conf, (0, 0, line))


def test_drop_noise_words_removes_low_confidence_specks_next_to_good_words():
    words = [_word("Jane", 92, x=0), _word("a", 20, x=20), _word("Testwell", 90, x=40), _word("|", 80, x=90)]
    assert [w.text for w in ce.drop_noise_words(words)] == ["Jane", "Testwell"]


def test_drop_noise_words_keeps_a_short_doubtful_word_when_it_is_all_there_is():
    alone = [_word("Al", 30)]
    assert [w.text for w in ce.drop_noise_words(alone)] == ["Al"]
    # una palabra corta pero bien leida se queda
    assert [w.text for w in ce.drop_noise_words([_word("Jane", 92), _word("Li", 85, x=20)])] == ["Jane", "Li"]


def test_glyph_fix_uses_the_name_to_resolve_i_l_confusions():
    email = FieldRead(value="dvaidivia@example.com", agreement=1.0, reads=[("dvaidivia@example.com", 90.0)], weights=[1.0], support=1.0)
    fixed, changed = ce.fix_email_glyphs_by_name(email, "Daniel Valdivia")
    assert changed and fixed.value == "dvaldivia@example.com"
    assert fixed.agreement <= 0.7  # queda en amarillo para que alguien lo mire


def test_glyph_fix_leaves_alone_an_email_that_already_matches_or_has_no_clear_better_option():
    good = FieldRead(value="dvaldivia@example.com", agreement=1.0, reads=[("dvaldivia@example.com", 90.0)], weights=[1.0], support=1.0)
    assert ce.fix_email_glyphs_by_name(good, "Daniel Valdivia") == (good, False)
    unrelated = FieldRead(value="xilo@example.com", agreement=1.0, reads=[("xilo@example.com", 90.0)], weights=[1.0], support=1.0)
    assert ce.fix_email_glyphs_by_name(unrelated, "Daniel Valdivia")[1] is False
    assert ce.fix_email_glyphs_by_name(good, "")[1] is False


def test_parse_page_of_reads_page_n_of_m_near_the_bol_header():
    words = [_word("Page:", 90, x=700), _word("2", 90, x=760), _word("OF", 90, x=780), _word("3", 90, x=810)]
    assert ce.parse_page_of(words, 0.0, 1000.0, 1000.0) == (2, 3)
    # lejos del encabezado o fuera de rango: no cuenta
    assert ce.parse_page_of(words, 500.0, 1000.0, 1000.0) is None
    bad = [_word("Page:", 90, x=700), _word("5", 90, x=760), _word("OF", 90, x=780), _word("3", 90, x=810)]
    assert ce.parse_page_of(bad, 0.0, 1000.0, 1000.0) is None


def test_number_consensus_votes_whole_numbers_and_reports_the_share():
    reads = [("98163300", 90.0)] * 7 + [("98463300", 80.0)] * 2 + [("", 0.0)]
    vote = ce.number_consensus(reads)
    assert vote.value == "98163300"
    assert abs(vote.agreement - 0.7) < 1e-9
    assert ce.number_consensus([("12", 90.0)]).value == ""  # muy corto para ser un BOL


def _read(value, share, total=10, other="55512345"):
    n = round(share * total)
    reads = [(value, 90.0)] * n + [(other, 80.0)] * (total - n)
    return ce.number_consensus(reads)


def test_combine_bol_trusts_two_sources_that_agree_or_one_that_is_clearly_read():
    header, tracking = _read("98163300", 0.9), _read("98163300", 0.9)
    bol, track = ce.combine_bol(header, tracking)
    assert (bol.value, bol.agreement, track) == ("98163300", 1.0, "98163300")
    # el encabezado se lee claro y la fila de abajo no: el encabezado manda
    bol, _ = ce.combine_bol(_read("98163300", 0.9), _read("98163300", 0.4))
    assert bol.value == "98163300" and bol.agreement == 1.0


def test_combine_bol_sends_a_number_confirmed_only_by_the_tracking_row_to_review():
    weak_header = ce.number_consensus([("55512345", 70.0)] * 4 + [("98163300", 70.0)] * 4 + [("", 0.0)] * 2)
    bol, _ = ce.combine_bol(weak_header, _read("98163300", 0.9))
    assert bol.value == "98163300" and bol.agreement <= 0.9


def test_combine_bol_without_tracking_row_needs_a_clear_header():
    clear, _ = ce.combine_bol(_read("98163300", 0.9), None)
    assert clear.agreement == 1.0
    unclear, _ = ce.combine_bol(_read("98163300", 0.5), None)
    assert unclear.agreement < 0.99


def test_two_sources_that_agree_on_a_split_digit_still_go_to_review_with_the_other_reading():
    # el encabezado y la fila de abajo leen "96433794", pero la mitad de las lecturas dicen "98433794":
    # un 8 que parece 6 se lee igual en los dos lugares, asi que coincidir no basta
    def split(winner, rival, wins, rivals):
        return ce.number_consensus([(winner, 80.0)] * wins + [(rival, 80.0)] * rivals + [("", 0.0)] * (19 - wins - rivals))

    bol, _ = ce.combine_bol(split("96433794", "98433794", 10, 8), split("96433794", "98433794", 6, 2))
    assert bol.value == "96433794" and bol.agreement <= 0.9 and bol.alt == "98433794"


def test_a_single_stray_reading_of_a_similar_number_does_not_cast_doubt():
    header = ce.number_consensus([("98163300", 90.0)] * 15 + [("98463300", 60.0)] * 2 + [("", 0.0)] * 2)
    tracking = ce.number_consensus([("98163300", 90.0)] * 15 + [("98463300", 60.0)] * 2 + [("", 0.0)] * 2)
    bol, _ = ce.combine_bol(header, tracking)
    assert bol.agreement == 1.0 and bol.alt == ""


def test_a_rival_that_differs_in_more_than_one_digit_is_not_a_split_digit():
    assert ce.bol_rival(_read("98163300", 0.6, other="98999300"), None, "98163300") == ""
    assert ce.bol_rival(_read("98163300", 0.6, other="98463300"), None, "98163300") == "98463300"
