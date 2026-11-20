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


def _dims_xlsx(headers, rows, sheet="VRP 53458"):
    from openpyxl import Workbook

    book = Workbook()
    ws = book.active
    ws.title = sheet
    ws.append(headers)
    for row in rows:
        ws.append(row)
    buffer = io.BytesIO()
    book.save(buffer)
    return buffer.getvalue()


def test_dims_reads_by_header_with_any_column_set():
    # con TRACKER y HAZMAT vacio (como el archivo recibido)
    full = _dims_xlsx(
        ["CARRIER", "TRAILER", "Barcode", "Length", "Width", "Height", "TRACKER", "Actual Weight", "BOX", "RACK", "HAZMAT", "Scanned Time"],
        [["VRP ", 53458, 3049264, 51, 36, 86, "j1514571", 1000, 1, None, None, "Nov 20 2026"],
         ["VRP ", 53458, 3049264, 51, 36, 86, "j1513918", 1000, 1, None, None, "Nov 20 2026"],
         ["VRP ", 53458, 3049280, 49.8, 42.3, 15.3, None, 100, 1, None, "Y", "Nov 20 2026"]],
    )
    dims = de.read_dims(full, "dims.xlsx")
    assert dims.barcodes == ["3049264", "3049280"] and dims.carriers == {"3049264": "VRP", "3049280": "VRP"}
    assert dims.hazmat == {"3049280": True} and dims.trackers["3049264"] == ["j1514571", "j1513918"]
    assert dims.rows[0]["weight"] == 1000 and dims.missing == []

    # sin HAZMAT ni TRACKER ni RACK y con las columnas en otro orden (como el template)
    short = _dims_xlsx(["Barcode", "CARRIER", "Length", "Width", "Height", "Actual Weight", "BOX"],
                       [[98162916, "BEX", 48, 40, 40, 107, 2]], sheet="DATA")
    dims = de.read_dims(short, "x.xlsx")
    assert dims.barcodes == ["98162916"] and dims.carriers == {"98162916": "BEX"}
    assert dims.hazmat == {} and sorted(dims.missing) == ["hazmat", "tracker"]

    csv = b"CARRIER,Barcode,HAZMAT\nBEX,98162916,\nBEX,98162932,YES\n"
    dims = de.read_dims(csv, "dims.csv")
    assert dims.hazmat == {"98162932": True}


def test_dims_without_barcode_is_rejected():
    import pytest

    with pytest.raises(ValueError, match="barcode_header"):
        de.read_dims(_dims_xlsx(["CARRIER", "Length"], [["BEX", 1]]), "x.xlsx")


def test_full_workbook_builds_data_from_dims_in_template_order():
    dims = de.read_dims(
        _dims_xlsx(
            ["CARRIER", "TRAILER", "Barcode", "Length", "Width", "Height", "TRACKER", "Actual Weight", "BOX", "RACK", "HAZMAT", "Scanned Time"],
            [["VRP ", 53458, 3049901, 48, 40, 40, "j1", 100, 1, None, None, "Nov 19 2026"]],
        ),
        "d.xlsx",
    )
    extras = {"pod": "11/19/26 20:54", **de.cutoff_messages("11/19/26 1356", "11/19/26 20:54")}
    blob, _ = de.build_workbook(_rows()[:1], "11/19/26 1356", dims, extras=extras)
    book = load_workbook(io.BytesIO(blob))
    assert book.sheetnames == ["DATE", "DATA"]
    data = book["DATA"]
    assert [data.cell(1, c).value for c in (1, 3, 7, 14, 17, 18)] == [
        "CARRIER", "Barcode", "Actual Weight", "READY DATE", "SERVICE LEVEL", "DELIVERY DATE"]
    assert data["G2"].value == 100  # el peso va en G aunque el DIMS traiga TRACKER antes
    assert data["S1"].value == "TRACKER" and data["S2"].value == "j1"
    assert (data["Q2"].value, data["R2"].value) == ("3D", "11/24/26")
    assert data["N2"].value == "11/19/26 1356" and data["N3"].value == "POD DATE TIME" and data["N4"].value == "11/19/26 20:54"
    assert data["O18"].value == "LEFT PLANT 11/19/26 1356 ARRIVED AT 11/19/26 20:54"


def test_cutoff_messages_like_the_template():
    assert de.cutoff_messages("11/19/26 1356", "11/19/26 20:54") == {
        "both": "LEFT PLANT 11/19/26 1356 ARRIVED AT 11/19/26 20:54",
        "left": "LEFT PLANT 11/19/26 1356", "arrived": "ARRIVED AT 11/19/26 20:54"}
    assert de.cutoff_messages("11/19/26 1356", "")["both"] == ""
    assert de.cutoff_messages("", "11/07/26 10:44", weekday=True)["arrived"] == "ARRIVED AT SATURDAY 11/07/26 10:44"


def test_file_name_defaults_to_ready_date():
    assert de.safe_filename("", "11/19/26 1356") == "11-19-26 1356.xlsx"
    assert de.safe_filename("Mi envio.xlsx", "11/19/26 1356") == "Mi envio.xlsx"
    assert de.safe_filename("a/b:c", "x") == "a-b-c.xlsx"
    assert de.safe_filename("", "") == "DATE_IMPORT.xlsx"


def test_inbond_priority_code_goes_in_fk():
    blob, _ = de.build_workbook(_rows()[:1], "11/19/26 1356", priority=de.PRIORITY_INBOND)
    sheet = load_workbook(io.BytesIO(blob))["DATE"]
    assert sheet["FK1"].value == "Priority Code" and sheet["FK2"].value == "INBD"
    blob, _ = de.build_workbook(_rows()[:1], "11/19/26 1356")
    assert load_workbook(io.BytesIO(blob))["DATE"]["FK2"].value is None


def test_ready_is_optional_and_leaves_dependent_fields_blank():
    row = de.ShipmentRow("3049901", state="TX", carrier="VRP")
    result = de.build_date_row(row, "")
    values = result.values
    assert (values["Q"], values["R"], values["S"], values["AF"]) == ("", "", "", "")
    assert values["AG"] == "1600" and values["AK"] == "3D" and values["E"] == "144463"
    assert "no_ready" in result.warnings and result.delivery == ""
