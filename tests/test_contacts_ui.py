import shutil
import time

import pytest

from tests.synthetic import SAMPLE, make_bol_pdf

pytestmark = pytest.mark.skipif(shutil.which("tesseract") is None, reason="Tesseract no instalado")

PDF = make_bol_pdf(SAMPLE)


def _app():
    import streamlit as st

    import contacts_tab
    import delivery_ui
    from tests.test_contacts_ui import PDF

    class Upload:
        def __init__(self, name, content):
            self.name, self._content = name, content

        def getvalue(self):
            return self._content

    def fake_uploader(label, type=None, accept_multiple_files=False, key=None, help=None):
        if key == "ct_pdfs":
            return [Upload("VRP_test.pdf", PDF)]
        if key == "ct_data" and st.session_state.get("_dims"):
            return Upload("dims.xlsx", st.session_state["_dims"])
        return None

    st.file_uploader = fake_uploader
    if "language" not in st.session_state:
        st.session_state.language = "es"

    def t(key, **values):
        text = {**contacts_tab.TEXT[st.session_state.language], **delivery_ui.TEXT[st.session_state.language]}[key]
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


def _dims_bytes():
    import io

    from openpyxl import Workbook

    book = Workbook()
    ws = book.active
    ws.append(["CARRIER", "TRAILER", "Barcode", "Length", "Width", "Height", "TRACKER", "Actual Weight", "BOX", "HAZMAT"])
    for code, hazmat in (("3049901", None), ("3049902", "Y"), ("3049999", None)):
        ws.append(["VRP ", 53458, int(code), 48, 40, 40, "j123", 100, 1, hazmat])
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def test_dims_comparison_carrier_messages_and_file_name(monkeypatch, tmp_path):
    from streamlit.testing.v1 import AppTest

    import contacts_tab

    monkeypatch.chdir(tmp_path)
    contacts_tab.get_memory.clear()
    at = AppTest.from_function(_app, default_timeout=120)
    at.session_state["_dims"] = _dims_bytes()
    at.run()
    assert not at.exception
    ready = next(w for w in at.text_input if w.label.startswith("READY DATE"))
    ready.set_value("11/19/26 1356").run()
    pod = next(w for w in at.text_input if w.label.startswith("POD DATE TIME"))
    pod.set_value("11/19/26 20:54").run()
    assert "LEFT PLANT 11/19/26 1356 ARRIVED AT 11/19/26 20:54" in [c.value for c in at.code]

    at.button(key="ct_read_btn").click().run()
    assert not at.exception
    warnings = " | ".join(w.value for w in at.warning)
    assert "3049999" in warnings and "3049903" in warnings  # el DIMS y los PDF no son iguales
    name = next(w for w in at.text_input if w.label.startswith("Nombre del archivo"))
    assert name.placeholder == "11/19/26 1356"  # sin nombre se guarda con el READY DATE
    assert any("11-19-26 1356.xlsx" in c.value for c in at.caption)


def _download_buttons(at):
    return [el for el in at.main if getattr(el, "type", "") == "download_button" or el.type == "download_button"]


def test_ready_is_optional_and_delivery_can_be_chosen_for_all(monkeypatch, tmp_path):
    from datetime import date, time as dtime

    from streamlit.testing.v1 import AppTest

    import contacts_tab
    import date_export as de

    calls = []
    original = de.build_workbook

    def spy(shipments, ready, dims=None, **kwargs):
        calls.append({"ready": ready, "rows": list(shipments), **kwargs})
        return original(shipments, ready, dims, **kwargs)

    monkeypatch.setattr(de, "build_workbook", spy)
    monkeypatch.chdir(tmp_path)
    contacts_tab.get_memory.clear()
    at = AppTest.from_function(_app, default_timeout=120)
    at.run()
    at.button(key="ct_read_btn").click().run()  # sin READY DATE
    assert not at.exception
    assert any("Sin READY DATE" in i.value for i in at.info)  # avisa, pero deja continuar
    assert len(_download_buttons(at)) >= 1
    assert calls and calls[-1]["ready"] == "" and calls[-1].get("default_date", "") == ""

    # Una sola seleccion de fecha y hora para todos
    at.date_input(key="ct_deliv_date").set_value(date(2026, 12, 1)).run()
    at.time_input(key="ct_deliv_time").set_value(dtime(9, 30)).run()
    assert not at.exception
    assert calls[-1]["default_date"] == "12/01/26" and calls[-1]["default_time"] == "0930"
    assert calls[-1]["ready"] == ""

    # Con un READY mal escrito no se exporta en blanco por error
    next(w for w in at.text_input if w.label.startswith("READY DATE")).set_value("mañana").run()
    assert any("Corrige el READY DATE" in i.value for i in at.info)
    assert _download_buttons(at) == []
