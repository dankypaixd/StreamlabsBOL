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
        if key == "ct_pdfs":
            return [Upload("VRP_test.pdf", PDF)]
        if key == "ct_data" and st.session_state.get("_dims"):
            return Upload("dims.xlsx", st.session_state["_dims"])
        return None

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
    return [el for el in at.main if el.type == "download_button"]


def test_ready_is_optional_and_messages_follow_what_is_typed(monkeypatch, tmp_path):
    from streamlit.testing.v1 import AppTest

    import contacts_tab

    monkeypatch.chdir(tmp_path)
    contacts_tab.get_memory.clear()
    at = AppTest.from_function(_app, default_timeout=120)
    at.run()
    at.button(key="ct_read_btn").click().run()  # sin READY DATE
    assert not at.exception
    assert len(_download_buttons(at)) >= 1  # se puede descargar con los campos en blanco
    assert not at.warning or not any("push" in w.value for w in at.warning)
    assert list(at.code) == []  # sin READY DATE ni POD no se ve ningun mensaje

    # POD solo: sale ARRIVED AT con lo que se escribio
    next(w for w in at.text_input if w.label.startswith("POD DATE TIME")).set_value("11/20/26 0800").run()
    assert [c.value for c in at.code] == ["ARRIVED AT 11/20/26 0800"]

    # READY antes de las 17:00: LEFT PLANT sin aviso; despues de las 17:00: aviso de push
    ready = next(w for w in at.text_input if w.label.startswith("READY DATE"))
    ready.set_value("11/20/26 1456").run()
    assert "LEFT PLANT 11/20/26 1456" in [c.value for c in at.code]
    assert not any("push" in w.value for w in at.warning)
    ready.set_value("11/20/26 1701").run()
    assert any("17:01" in w.value and "push" in w.value for w in at.warning)
    ready.set_value("11/20/26 1700").run()  # justo a las 17:00 todavia esta dentro del cut-off
    assert not any("push" in w.value for w in at.warning)

    # Un READY mal escrito no se exporta en blanco por error
    ready.set_value("mañana").run()
    assert any("Corrige el READY DATE" in i.value for i in at.info)
    assert _download_buttons(at) == []


def test_tables_show_ready_date_and_time_instead_of_delivery():
    import contact_review as cr
    import contacts_tab
    import inbonds_tab

    def t(key, **values):
        text = contacts_tab.TEXT["es"][key]
        return text.format(**values) if values else text

    assert contacts_tab.ready_parts("11/19/26 1356") == ("11/19/26", "1356")
    assert contacts_tab.ready_parts("") == ("", "")
    contact = cr.Contact(key="1", bol="3049901", source_file="x.pdf", pages=[1], carrier="VRP", state="TX")
    for table in (contacts_tab._table(t, [contact], "11/19/26 1356"), inbonds_tab._table(t, [contact], "11/19/26 1356")):
        assert list(table.columns)[-2:] == ["Fecha ready", "Hora ready"]
        assert table.iloc[0]["Fecha ready"] == "11/19/26" and table.iloc[0]["Hora ready"] == "1356"
        assert "Entrega" not in table.columns
    # sin READY DATE esas columnas no aparecen
    assert "Fecha ready" not in contacts_tab._table(t, [contact], "").columns
    assert contacts_tab._service_text(t, "TX", "11/19/26 1356") == "Service level: 3D · READY DATE: 11/19/26 · 1356"
    assert contacts_tab._service_text(t, "TX", "") == "Service level: 3D"
