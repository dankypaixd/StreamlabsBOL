"""Fecha de entrega opcional: calculada con el READY DATE y el service level, editable y visible antes de descargar."""
import shutil

import pytest

pytestmark = pytest.mark.skipif(shutil.which("tesseract") is None, reason="Tesseract no instalado")

COL = "Fecha de entrega"


def _spy_app():
    import streamlit as st

    import date_export as de
    from tests.test_contacts_ui import _app

    real = de.build_workbook
    if not getattr(real, "_spy", False):
        def spy(*args, **kwargs):
            blob, results = real(*args, **kwargs)
            st.session_state["_rows"] = [(r.bol, r.values["AF"], r.values["AG"]) for r in results]
            return blob, results

        spy._spy = True
        de.build_workbook = spy
    _app()


def _read(monkeypatch, tmp_path, ready="11/19/26 1356"):
    from streamlit.testing.v1 import AppTest

    import contacts_tab

    monkeypatch.chdir(tmp_path)
    contacts_tab.get_memory.clear()
    at = AppTest.from_function(_spy_app, default_timeout=120)
    at.run()
    next(w for w in at.text_input if w.label.startswith("READY DATE")).set_value(ready).run()
    at.button(key="ct_read_btn").click().run()
    assert not at.exception
    return at


def _delivery_frame(at):
    frames = [d.value for d in at.dataframe if COL in d.value.columns]
    assert len(frames) == 1
    return frames[0]


def _download_disabled(at):
    (button,) = [el for el in at.main if el.type == "download_button"]
    return button.proto.disabled


def _edit(at, edits):
    """Cambia celdas del editor de la vista previa (AppTest no edita tablas: se escribe su estado).

    Como en el navegador, los cambios se acumulan en el mismo editor: la llave no cambia entre cambio y cambio.
    """
    key = at.session_state["ct_delivery_editor_key"]
    current = at.session_state[key]["edited_rows"] if key in at.session_state.keys() else {}
    rows = {int(i): dict(cells) for i, cells in current.items()}  # copia: el estado de un widget es de solo lectura
    for index, value in edits.items():
        rows.setdefault(index, {})[COL] = value
    at.session_state[key] = {"edited_rows": rows, "added_rows": [], "deleted_rows": []}
    at.run()
    assert not at.exception
    assert at.session_state["ct_delivery_editor_key"] == key  # no se perdio el editor


def test_delivery_is_off_by_default_and_nothing_about_it_shows(monkeypatch, tmp_path):
    at = _read(monkeypatch, tmp_path)
    assert not at.checkbox(key="ct_include_delivery").value
    assert not [d for d in at.dataframe if COL in d.value.columns]
    assert not [w for w in at.text_input if w.label.startswith("Fecha de entrega")]
    assert all(row[1:] == ("", "") for row in at.session_state["_rows"])  # AF y AG vacias


def test_delivery_is_computed_from_the_ready_date_and_the_service_level(monkeypatch, tmp_path):
    at = _read(monkeypatch, tmp_path)  # jueves 19, 3D: viernes 20, lunes 23 y martes 24
    at.checkbox(key="ct_include_delivery").check().run()
    assert not at.exception
    assert list(_delivery_frame(at)[COL]) == ["11/24/26"] * 3
    assert at.session_state["_rows"] == [(b, "11/24/26", "1600") for b in ("3049901", "3049902", "3049903")]
    assert not _download_disabled(at)

    # cambiar el READY DATE recalcula solo: lunes 23, 3D = jueves 26
    next(w for w in at.text_input if w.label.startswith("READY DATE")).set_value("11/23/26 0800").run()
    assert list(_delivery_frame(at)[COL]) == ["11/26/26"] * 3
    assert at.session_state["_rows"][0][1] == "11/26/26"


def test_a_date_for_all_the_bols_replaces_the_calculated_one(monkeypatch, tmp_path):
    at = _read(monkeypatch, tmp_path)
    at.checkbox(key="ct_include_delivery").check().run()
    at.text_input(key="ct_delivery_all").set_value("11/30/2026").run()
    assert list(_delivery_frame(at)[COL]) == ["11/30/26"] * 3
    assert {row[1] for row in at.session_state["_rows"]} == {"11/30/26"}

    at.text_input(key="ct_delivery_all").set_value("30/11/26").run()  # no es una fecha
    assert any("no es válida" in w.value for w in at.warning) and _download_disabled(at)


def test_a_date_changed_in_the_table_goes_to_the_excel_and_survives_a_new_ready_date(monkeypatch, tmp_path):
    at = _read(monkeypatch, tmp_path)
    at.checkbox(key="ct_include_delivery").check().run()
    _edit(at, {1: "12/01/2026"})
    assert at.session_state["ct_delivery_over"] == {"3049902": "12/01/26"}  # normalizada a MM/DD/YY
    assert at.session_state["_rows"] == [("3049901", "11/24/26", "1600"), ("3049902", "12/01/26", "1600"), ("3049903", "11/24/26", "1600")]
    assert any("1 fecha(s) de entrega cambiadas a mano: 3049902" in c.value for c in at.caption)
    assert not _download_disabled(at)

    next(w for w in at.text_input if w.label.startswith("READY DATE")).set_value("11/23/26 0800").run()
    assert list(_delivery_frame(at)[COL]) == ["11/26/26", "12/01/26", "11/26/26"]  # la de mano se queda

    at.button(key="ct_delivery_reset").click().run()  # "Volver a calcular todas"
    assert list(_delivery_frame(at)[COL]) == ["11/26/26"] * 3
    assert at.session_state["ct_delivery_over"] == {}


def test_a_cell_set_back_to_the_calculated_date_stops_being_manual_and_a_blank_cell_means_no_date(monkeypatch, tmp_path):
    at = _read(monkeypatch, tmp_path)
    at.checkbox(key="ct_include_delivery").check().run()
    _edit(at, {0: "11/24/26", 2: ""})
    assert at.session_state["ct_delivery_over"] == {"3049903": ""}  # la misma que la calculada no cuenta como cambio
    assert at.session_state["_rows"] == [("3049901", "11/24/26", "1600"), ("3049902", "11/24/26", "1600"), ("3049903", "", "")]


def test_an_invalid_date_in_the_table_blocks_the_download_until_fixed(monkeypatch, tmp_path):
    at = _read(monkeypatch, tmp_path)
    at.checkbox(key="ct_include_delivery").check().run()
    _edit(at, {0: "13/45/26"})
    assert any("no válida en: 3049901" in w.value for w in at.warning) and _download_disabled(at)
    _edit(at, {0: "11/27/26"})
    assert not any("no válida" in w.value for w in at.warning) and not _download_disabled(at)


def test_two_changes_in_a_row_both_stay(monkeypatch, tmp_path):
    """Regresion: la llave del editor no debe cambiar al guardar el primer cambio, o el segundo se perderia."""
    at = _read(monkeypatch, tmp_path)
    at.checkbox(key="ct_include_delivery").check().run()
    _edit(at, {0: "12/01/26"})
    _edit(at, {1: "12/02/26"})
    assert at.session_state["ct_delivery_over"] == {"3049901": "12/01/26", "3049902": "12/02/26"}
    assert [r[1] for r in at.session_state["_rows"]] == ["12/01/26", "12/02/26", "11/24/26"]
