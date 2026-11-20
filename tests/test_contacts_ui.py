import shutil
import time

import pytest

from tests.synthetic import SAMPLE, make_bol_pdf

pytestmark = pytest.mark.skipif(shutil.which("tesseract") is None, reason="Tesseract no instalado")

PDF = make_bol_pdf(SAMPLE)


def _app():
    import streamlit as st

    import contacts_tab
    from tests.test_contacts_ui import PDF

    class Upload:
        def __init__(self, name, content):
            self.name, self._content = name, content

        def getvalue(self):
            return self._content

    def fake_uploader(label, type=None, accept_multiple_files=False, key=None, help=None):
        return [Upload("VRP_test.pdf", PDF)] if key == "ct_pdfs" else None

    st.file_uploader = fake_uploader
    if "language" not in st.session_state:
        st.session_state.language = "es"

    def t(key, **values):
        text = contacts_tab.TEXT[st.session_state.language][key]
        return text.format(**values) if values else text

    contacts_tab.render(t)


def test_read_confirm_and_export(monkeypatch, tmp_path):
    from streamlit.testing.v1 import AppTest

    import contacts_tab

    monkeypatch.chdir(tmp_path)  # la memoria local se guarda aqui
    contacts_tab.get_memory.clear()
    at = AppTest.from_function(_app, default_timeout=120)
    at.run()
    assert not at.exception

    at.text_input(key="ct_ready_nodata").set_value("11/19/26 1356").run()
    at.button(key="ct_read_btn").click().run()
    assert not at.exception
    metrics = {m.label: m.value for m in at.metric}
    assert metrics["🟢 Sin error"] == "2" and metrics["🔴 Incompletos"] == "1"

    # El incompleto (bloque vacio) se confirma a mano: el estado sale del ZIP y el service level se calcula
    at.button(key="fail_3049903_ok").click().run()
    assert not at.exception
    assert {m.label: m.value for m in at.metric}["✔ Confirmados"] == "1"

    at.button(key="ct_confirm_all").click().run()
    assert {m.label: m.value for m in at.metric}["✔ Confirmados"] == "3"
    saved = tmp_path / ".bol_memory" / "contacts.bin"
    for _ in range(50):  # el guardado corre en un hilo aparte
        if saved.exists():
            break
        time.sleep(0.1)
    assert saved.exists()  # aprendio en segundo plano
