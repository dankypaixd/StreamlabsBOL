import dataclasses

import date_export as de
import delivery_ui


def _base(rows, ready="11/19/26 1356", date="", time=""):
    return [de.build_date_row(r, ready, "", date, time) for r in rows]


ROWS = [de.ShipmentRow("3049901", state="TX", carrier="VRP"), de.ShipmentRow("3049902", state="FL", carrier="BEX")]


def test_untouched_table_changes_nothing():
    base = _base(ROWS)
    rows, invalid = delivery_ui.apply_row_edits(ROWS, base, [r.delivery for r in base], [r.time for r in base])
    assert invalid == [] and all(a is b for a, b in zip(ROWS, rows))


def test_edited_cells_become_row_overrides():
    base = _base(ROWS)
    rows, invalid = delivery_ui.apply_row_edits(ROWS, base, [base[0].delivery, "12/2/2026"], [base[0].time, "9:30"])
    assert invalid == [] and rows[0] is ROWS[0]
    assert (rows[1].delivery_date, rows[1].delivery_time) == ("12/02/26", "0930")


def test_cleared_cell_leaves_the_field_empty_and_invalid_is_ignored():
    base = _base(ROWS)
    rows, invalid = delivery_ui.apply_row_edits(ROWS, base, [None, "nunca"], [float("nan"), "99:99"])
    assert (rows[0].delivery_date, rows[0].delivery_time) == (de.BLANK, de.BLANK)
    assert rows[1] is ROWS[1] and invalid == ["3049902: nunca", "3049902: 99:99"]


def test_only_cells_different_from_the_global_choice_count():
    base = _base(ROWS, date="11/30/26", time="0900")  # lo elegido para todos ya esta en la tabla
    rows, _ = delivery_ui.apply_row_edits(ROWS, base, ["11/30/26", "12/01/26"], ["0900", "0900"])
    assert rows[0] is ROWS[0]
    assert dataclasses.asdict(rows[1])["delivery_date"] == "12/01/26" and rows[1].delivery_time == ""
