import shutil
import time

import pytest

from tests.synthetic import SAMPLE, make_bol_pdf

pytestmark = pytest.mark.skipif(shutil.which("tesseract") is None, reason="Tesseract no instalado")

PDF = make_bol_pdf(SAMPLE)


def _app():
    import streamlit as st

    import contacts_tab
    import dims_compare
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
        texts = {**contacts_tab.TEXT[st.session_state.language], **dims_compare.TEXT[st.session_state.language]}
        return texts[key].format(**values) if values else texts[key]

    contacts_tab.render(t)


def test_read_confirm_and_export(monkeypatch, tmp_path):
    from streamlit.testing.v1 import AppTest

    import contacts_tab

    monkeypatch.chdir(tmp_path)  # la memoria local se guarda aqui
    contacts_tab.get_memory.clear()
    at = AppTest.from_function(_app, default_timeout=120)
    at.run()
    assert not at.exception

    at.text_input(key="ct_ready").set_value("11/19/26 1356").run()
    at.button(key="ct_read_btn").click().run()
    assert not at.exception
    metrics = {m.label: m.value for m in at.metric}
    # El bloque Ship To Attention vacio (BOL 3049903) ya no sale en rojo; solo queda en amarillo porque su
    # estado se dedujo del ZIP
    assert metrics["🟢 Sin error"] == "2" and metrics["🔴 Incompletos"] == "0" and metrics["🟡 Dudosos"] == "1"

    # El de bloque vacio se confirma a mano: el estado sale del ZIP y el service level se calcula
    at.button(key="doubt_3049903_ok").click().run()
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


def _dims_bytes(codes=("3049901", "3049902", "3049999")):
    import io

    from openpyxl import Workbook

    book = Workbook()
    ws = book.active
    ws.append(["CARRIER", "TRAILER", "Barcode", "Length", "Width", "Height", "TRACKER", "Actual Weight", "BOX", "HAZMAT"])
    for code in codes:
        ws.append(["VRP ", 53458, int(code), 48, 40, 40, "j123", 100, 1, None])
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def _download_buttons(at):
    return [el for el in at.main if el.type == "download_button"]


def _read_with_dims(monkeypatch, tmp_path, dims=None):
    from streamlit.testing.v1 import AppTest

    import contacts_tab

    monkeypatch.chdir(tmp_path)
    contacts_tab.get_memory.clear()
    at = AppTest.from_function(_app, default_timeout=120)
    if dims is not None:
        at.session_state["_dims"] = dims
    at.run()
    assert not at.exception
    ready = next(w for w in at.text_input if w.label.startswith("READY DATE"))
    ready.set_value("11/19/26 1356").run()
    at.button(key="ct_read_btn").click().run()
    assert not at.exception
    return at


def test_dims_is_optional_and_without_it_the_date_downloads(monkeypatch, tmp_path):
    at = _read_with_dims(monkeypatch, tmp_path)
    (button,) = _download_buttons(at)
    assert not button.proto.disabled and button.proto.label == "Descargar la hoja DATE"  # solo la hoja DATE
    assert not any("DIMS" in w.value for w in at.warning)  # sin DIMS no hay nada que comparar ni que avisar


def test_dims_that_matches_shows_ok_and_does_not_block(monkeypatch, tmp_path):
    same = _dims_bytes(("3049901", "3049902", "3049903"))
    at = _read_with_dims(monkeypatch, tmp_path, same)
    assert any("Los 3 BOL de los PDF son los mismos que los del DIMS" in s.value for s in at.success)
    assert not at.multiselect and not [c for c in at.checkbox if c.label == "Ignorar la comparación"]
    assert not _download_buttons(at)[0].proto.disabled


def test_dims_differences_block_the_download_until_accepted_or_ignored(monkeypatch, tmp_path):
    at = _read_with_dims(monkeypatch, tmp_path, _dims_bytes())  # el DIMS trae 3049999 y no trae 3049903
    warnings = " | ".join(w.value for w in at.warning)
    assert "PDF 3" in warnings and "DIMS 3" in warnings
    markdown = " | ".join(m.value for m in at.markdown)
    assert "3049903" in markdown and "3049999" in markdown  # cada diferencia, en su lista
    name = next(w for w in at.text_input if w.label.startswith("Nombre del archivo"))
    assert name.placeholder == "11/19/26 1356"  # sin nombre se guarda con el READY DATE
    assert any("11-19-26 1356.xlsx" in c.value for c in at.caption)

    # Sin resolver: la descarga queda deshabilitada y se dice por que
    assert _download_buttons(at)[0].proto.disabled
    assert any("acepta los BOL que no coinciden" in i.value for i in at.info)

    # Aceptar uno solo no alcanza
    picker = at.multiselect[0]
    assert picker.options == ["3049903 · solo en PDF", "3049999 · solo en DIMS"]  # dice de que lado esta cada uno
    picker.select("3049903 · solo en PDF").run()
    assert _download_buttons(at)[0].proto.disabled
    assert any("Faltan 1 BOL" in i.value for i in at.info)

    # Aceptar las dos diferencias habilita la descarga
    at.multiselect[0].select("3049999 · solo en DIMS").run()
    assert not _download_buttons(at)[0].proto.disabled
    assert any("Todas las diferencias están aceptadas" in c.value for c in at.caption)


def test_ignoring_the_comparison_enables_the_download(monkeypatch, tmp_path):
    at = _read_with_dims(monkeypatch, tmp_path, _dims_bytes())
    assert _download_buttons(at)[0].proto.disabled
    next(c for c in at.checkbox if c.label == "Ignorar la comparación").check().run()
    assert not _download_buttons(at)[0].proto.disabled
    assert any("Comparación ignorada" in c.value for c in at.caption)
    assert at.multiselect[0].proto.disabled  # ya no hace falta elegir BOL


def test_a_dims_without_barcode_is_reported_and_does_not_block(monkeypatch, tmp_path):
    import io

    from openpyxl import Workbook

    book = Workbook()
    book.active.append(["CARRIER", "Length"])
    book.active.append(["BEX", 1])
    buffer = io.BytesIO()
    book.save(buffer)
    at = _read_with_dims(monkeypatch, tmp_path, buffer.getvalue())
    assert any("no encontré una columna llamada Barcode" in e.value for e in at.error)
    assert not _download_buttons(at)[0].proto.disabled


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
