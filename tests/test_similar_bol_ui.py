"""El boton "Usar X" de las tarjetas con un BOL dudoso que se parece a otro de la lista."""


def _app():
    import streamlit as st

    import contact_review as cr
    import contacts_tab as ct

    contact = cr.Contact(key="98463300", bol="98463300", source_file="VRP_x.pdf", pages=[1])
    contact.reasons = ["bol_similar:98163300", "bol_alt:98163300", "bol_alt:98463301"]  # repetido: un solo boton
    st.text_input("BOL", value=contact.bol, key="ct_bol_98463300")
    ct.similar_bol_buttons(lambda key, **v: f"{key}:{v.get('bol', '')}", contact, prefix="ct", key_k="98463300", k="doubt_98463300")


def test_each_similar_bol_gets_a_button_that_fills_the_bol_field_without_confirming():
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_function(_app, default_timeout=30)
    at.run()
    assert not at.exception
    labels = sorted(b.label for b in at.button)
    assert labels == ["ct_use_bol:98163300", "ct_use_bol:98463301"]

    at.button(key="doubt_98463300_use_98163300").click().run()
    assert not at.exception
    assert at.text_input(key="ct_bol_98463300").value == "98163300"  # el usuario todavia tiene que confirmar


def test_reason_text_mentions_the_similar_bol():
    import contacts_tab as ct

    assert "98163300" in ct._reason_text(lambda key, **v: f"{key}:{v}", "bol_similar:98163300")


def test_reason_text_for_a_split_digit_names_the_other_reading():
    import contacts_tab as ct

    assert "98433794" in ct._reason_text(lambda key, **v: f"{key}:{v}", "bol_alt:98433794")
    assert "98433794" in ct.TEXT["es"]["r_bol_alt"].format(bol="98433794") and "{bol}" in ct.TEXT["en"]["r_bol_alt"]
