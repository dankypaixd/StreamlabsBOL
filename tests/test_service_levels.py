import service_levels as sl


def test_examples_from_user():
    # 11/19/26 0800 + 3D -> 11/24 ; despues de las 17:00 el dia 0 es el 20 -> 11/25
    assert sl.delivery_text("11/19/26 0800", "FL") == "11/24/26"
    assert sl.delivery_text("11/19/26 1356", "FL") == "11/24/26"
    assert sl.delivery_text("11/19/26 1800", "FL") == "11/25/26"
    assert sl.delivery_text("11/19/26 1700", "FL") == "11/24/26"
    assert sl.delivery_text("11/19/26 1701", "FL") == "11/25/26"


def test_labels_and_variants():
    assert sl.service_label("FL") == "3D"
    assert sl.service_label("tx2") == "3D"  # TX1/TX2/TX3 -> TX
    assert sl.service_label("PR") == "6D"
    assert sl.service_label("ZZ") == ""


def test_weekend_is_not_counted():
    # viernes 11/20/26 despues del cut-off -> dia 0 lunes 11/23 -> +3 = jueves 11/26
    assert sl.delivery_text("11/20/26 1800", "FL") == "11/26/26"
    # sabado -> dia 0 lunes
    assert sl.delivery_text("11/21/26 1000", "FL") == "11/26/26"


def test_zip_and_name_fallbacks():
    assert sl.state_from_zip("18202") == "PA"
    assert sl.state_from_zip("00979") == "PR"
    assert sl.state_from_name("San Juan 00979 Puerto Rico") == "PR"
    assert sl.state_from_name("Charleston West Virginia") == "WV"
