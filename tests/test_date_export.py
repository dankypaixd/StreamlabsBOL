import io

from openpyxl import load_workbook

import date_export as de


def _rows():
    return [
        de.ShipmentRow("3049901", "Jane Testwell", "jane@example.com", "1+15551230147", "FL", "VRP", False),
        de.ShipmentRow("98162916", "Omar F", "omar@sample.org", "1+15559876543", "PR", "BEX", True),
        de.ShipmentRow("3049903", carrier="VRP"),  # sin datos ni estado
    ]


def test_date_sheet_matches_template_columns():
    ready = de.format_ready("11/19/2026 13:56")
    assert ready == "11/19/26 1356"
    blob, results = de.build_workbook(_rows(), ready)
    sheet = load_workbook(io.BytesIO(blob))["DATE"]
    headers = de.load_headers()["DATE"]
    assert len(headers) == 198 and sheet.max_column == 198
    assert [c.value for c in sheet[1]] == headers
    row = {c: sheet[f"{c}2"].value for c in ("A", "B", "E", "Q", "R", "S", "AB", "AC", "AF", "AG", "AK", "CO", "DJ", "DK", "DU")}
    assert row == {
        "A": "DFW", "B": 3049901, "E": "144463", "Q": "11/19/26", "R": " 1356", "S": "1700",
        "AB": "Jane Testwell", "AC": "1+15551230147", "AF": "11/24/26", "AG": "1600",
        "AK": "3D", "CO": "HPDOM", "DJ": "RAX", "DK": "0", "DU": "jane@example.com",
    }
    assert sheet["E3"].value == "149570" and sheet["DK3"].value == "1" and sheet["AK3"].value == "6D"
    assert sheet["AK4"].value is None and "no_service" in results[2].warnings


def test_full_workbook_fills_service_and_delivery_in_data(tmp_path):
    data = de.DataSheet(
        headers=de.load_headers()["DATA"],
        rows=[["VRP", "VRP", 3049901, 48, 40, 40, 100, 1, None, None, "Nov 19 2026", None, None, "11/19/26 1356", None, None, None, None]],
        ready="11/19/26 1356",
    )
    blob, _ = de.build_workbook(_rows()[:1], data.ready, data)
    book = load_workbook(io.BytesIO(blob))
    assert book.sheetnames == ["DATE", "DATA"]
    assert (book["DATA"]["Q2"].value, book["DATA"]["R2"].value) == ("3D", "11/24/26")
    assert book["DATA"]["N2"].value == "11/19/26 1356"  # lo del usuario queda intacto
