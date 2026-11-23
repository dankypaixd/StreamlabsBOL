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
        "AB": "Jane Testwell", "AC": "1+15551230147", "AF": None, "AG": None,  # la entrega no va al Excel
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


def test_dims_reads_only_the_barcodes_by_header_in_any_column_order():
    # con TRACKER y HAZMAT (como el archivo recibido): solo importa Barcode, una caja por renglon
    full = _dims_xlsx(
        ["CARRIER", "TRAILER", "Barcode", "Length", "Width", "Height", "TRACKER", "Actual Weight", "BOX", "RACK", "HAZMAT", "Scanned Time"],
        [["VRP ", 53458, 3049264, 51, 36, 86, "j1514571", 1000, 1, None, None, "Nov 20 2026"],
         ["VRP ", 53458, 3049264, 51, 36, 86, "j1513918", 1000, 1, None, None, "Nov 20 2026"],
         ["VRP ", 53458, 3049280.0, 49.8, 42.3, 15.3, None, 100, 1, None, "Y", "Nov 20 2026"],
         [None, None, None, 1, 1, 1, None, 1, 1, None, None, None]],  # sin barcode: no es una caja
    )
    dims = de.read_dims(full, "dims.xlsx")
    assert dims.barcodes == ["3049264", "3049280"]
    assert dims.boxes == {"3049264": 2, "3049280": 1} and dims.sheet_name == "VRP 53458"

    # sin las demas columnas y con Barcode primero (como el template)
    short = _dims_xlsx(["Barcode", "CARRIER", "Length"], [[98162916, "BEX", 48]], sheet="DATA")
    assert de.read_dims(short, "x.xlsx").barcodes == ["98162916"]

    csv = b"CARRIER,Barcode,HAZMAT\nBEX,98162916,\nBEX,98162932,YES\n"
    assert de.read_dims(csv, "dims.csv").barcodes == ["98162916", "98162932"]


def test_dims_without_barcode_is_rejected():
    import pytest

    with pytest.raises(ValueError, match="barcode_header"):
        de.read_dims(_dims_xlsx(["CARRIER", "Length"], [["BEX", 1]]), "x.xlsx")


def test_the_workbook_only_has_the_date_sheet_and_no_dims_data():
    assert not hasattr(de, "_write_data_sheet") and not hasattr(de, "DATA_LAYOUT")
    blob, _ = de.build_workbook(_rows()[:1], "11/19/26 1356")
    assert load_workbook(io.BytesIO(blob)).sheetnames == ["DATE"]


def test_date_sheet_has_no_formatting_so_it_imports_as_is():
    """Lo que obligaba a copiar la hoja y pegar "solo valores": paneles, anchos, cadenas en linea y estilos."""
    import re
    import zipfile

    ready = "11/19/26 1356"
    blob, _ = de.build_workbook(_rows(), ready, priority=de.PRIORITY_INBOND)
    archive = zipfile.ZipFile(io.BytesIO(blob))
    sheet_xml = archive.read("xl/worksheets/sheet1.xml").decode()
    assert "<cols>" not in sheet_xml  # sin anchos de columna
    assert "<pane" not in sheet_xml and "<selection" not in sheet_xml  # sin paneles inmovilizados
    assert "inlineStr" not in sheet_xml and "<f>" not in sheet_xml  # cadenas compartidas como en Excel, sin formulas
    assert not re.search(r'<c [^>]*\bs="[1-9]', sheet_xml)  # ninguna celda con estilo propio
    assert not re.search(r"<(mergeCell|conditionalFormatting|dataValidation|autoFilter|tableParts|legacyDrawing|drawing)", sheet_xml)
    assert "xl/sharedStrings.xml" in archive.namelist()
    styles = archive.read("xl/styles.xml").decode()
    assert 'cellXfs count="1"' in styles  # solo el estilo por defecto

    sheet = load_workbook(io.BytesIO(blob))["DATE"]
    assert sheet.freeze_panes is None and len(sheet.column_dimensions) == 0
    assert all(cell.style_id == 0 for row in sheet.iter_rows() for cell in row)


def test_date_blank_values_are_really_empty_cells_and_types_are_kept():
    import re
    import zipfile

    blob, _ = de.build_workbook(_rows(), "11/19/26 1356")
    xml = zipfile.ZipFile(io.BytesIO(blob)).read("xl/worksheets/sheet1.xml").decode()
    written = re.findall(r'<c r="([A-Z]+)(\d+)"', xml)
    third_row = {col for col, row in written if row == "4"}  # BOL sin datos ni estado
    assert third_row == {"A", "B", "E", "Q", "R", "S", "CO", "DJ", "DK"}  # nada de AB, AC, AK, DU, T, FK en ""
    sheet = load_workbook(io.BytesIO(blob))["DATE"]
    assert isinstance(sheet["B2"].value, int) and sheet["B2"].value == 3049901  # el BOL numerico es numero
    assert sheet["E2"].value == "144463" and sheet["R2"].value == " 1356"  # el resto son textos (R con su espacio)
    assert sheet["T2"].value is None and sheet["FK2"].value is None
    # sin tocar nada: ni numeros convertidos ni textos que parezcan formulas o ligas
    odd, _ = de.build_workbook([de.ShipmentRow("3049901", name="=A1", state="TX", carrier="VRP")], "")
    cell = load_workbook(io.BytesIO(odd))["DATE"]["AB2"]
    assert cell.value == "=A1" and cell.data_type == "s"


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
    assert (values["Q"], values["R"], values["S"]) == ("", "", "")
    assert values["AK"] == "3D" and values["E"] == "144463" and "AF" not in values and "AG" not in values
    assert "no_ready" in result.warnings and result.delivery == ""


def test_consignee_number_goes_in_column_t_only_when_given():
    headers = de.load_headers()["DATE"]
    assert headers[19] == "Consignee Number"  # columna T
    rows = [
        de.ShipmentRow("3049901", state="TX", carrier="VRP", consignee="147229"),
        de.ShipmentRow("3049902", state="FL", carrier="BEX", consignee="210475"),
        de.ShipmentRow("3049903", state="TX", carrier="VRP"),
    ]
    blob, _ = de.build_workbook(rows, "11/19/26 1356")
    sheet = load_workbook(io.BytesIO(blob))["DATE"]
    assert [sheet[f"T{n}"].value for n in (2, 3, 4)] == ["147229", "210475", None]


def test_delivery_is_never_written_but_ready_date_is_in_every_workbook():
    ready = "11/19/26 1356"
    rows = _rows()[:1]
    plain, results = de.build_workbook(rows, ready)
    inbond, _ = de.build_workbook(rows, ready, priority=de.PRIORITY_INBOND)
    assert results[0].delivery == "11/24/26"  # se calcula para mostrarla en pantalla
    for blob in (plain, inbond):
        sheet = load_workbook(io.BytesIO(blob))["DATE"]
        # READY DATE: Pickup Date / Time / Close Time en todos los libros
        assert (sheet["Q2"].value, sheet["R2"].value, sheet["S2"].value) == ("11/19/26", " 1356", "1700")
        # Scheduled Delivery Date / Time (AF, AG): vacias; el service level (AK) se conserva
        assert (sheet["AF1"].value, sheet["AG1"].value) == ("Scheduled Delivery Date", "Scheduled Delivery Time")
        assert (sheet["AF2"].value, sheet["AG2"].value, sheet["AK2"].value) == (None, None, "3D")


def test_attention_and_service_level_can_each_be_left_out_of_the_sheet():
    rows = _rows()[:1]
    for attention, service in ((True, True), (False, True), (True, False), (False, False)):
        blob, results = de.build_workbook(rows, "11/19/26 1356", include_attention=attention, include_service=service)
        sheet = load_workbook(io.BytesIO(blob))["DATE"]
        assert (sheet["AB2"].value, sheet["AC2"].value, sheet["DU2"].value) == (
            ("Jane Testwell", "1+15551230147", "jane@example.com") if attention else (None, None, None)
        )
        assert sheet["AK2"].value == ("3D" if service else None)
        # lo demas no cambia, y el service level se sigue calculando para mostrarlo
        assert (sheet["B2"].value, sheet["E2"].value, sheet["Q2"].value, sheet["CO2"].value) == (3049901, "144463", "11/19/26", "HPDOM")
        assert results[0].service == "3D" and ("no_service" not in results[0].warnings)
    # sin service level no se avisa que falta el estado
    blank = de.build_date_row(de.ShipmentRow("3049903", carrier="VRP"), "11/19/26 1356", include_service=False)
    assert "no_service" not in blank.warnings
    assert "no_service" in de.build_date_row(de.ShipmentRow("3049903", carrier="VRP"), "11/19/26 1356").warnings
