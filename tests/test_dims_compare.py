import dims_compare as dc


def test_same_bols_and_same_quantity_match():
    result = dc.compare_bols(["3049902", "3049901", "3049901"], ["3049901", "3049902"])  # un BOL repetido cuenta una vez
    assert result.matches and result.differences == []
    assert result.pdf == result.dims == ["3049901", "3049902"]


def test_differences_are_listed_by_side_in_numeric_order():
    result = dc.compare_bols(["3049903", "98162916", "3049901"], ["3049901", "3049999", "98162932"])
    assert not result.matches
    assert result.only_pdf == ["3049903", "98162916"] and result.only_dims == ["3049999", "98162932"]
    assert result.differences == ["3049903", "3049999", "98162916", "98162932"]


def test_different_quantity_is_a_difference_even_if_every_pdf_bol_is_in_the_dims():
    result = dc.compare_bols(["3049901"], ["3049901", "3049902"])
    assert not result.matches and result.only_pdf == [] and result.only_dims == ["3049902"]


def test_blank_bols_are_not_compared():
    assert dc.compare_bols(["", " ", "3049901"], ["3049901", ""]).matches


def test_pending_differences_after_accepting_or_ignoring():
    result = dc.compare_bols(["3049901", "3049903"], ["3049901", "3049999"])
    assert dc.pending_after(result, [], False) == ["3049903", "3049999"]
    assert dc.pending_after(result, ["3049903"], False) == ["3049999"]
    assert dc.pending_after(result, ["3049903", "3049999"], False) == []
    assert dc.pending_after(result, [], True) == []  # ignorar la comparacion lo resuelve todo
    assert dc.pending_after(dc.compare_bols(["3049901"], ["3049901"]), [], False) == []


def test_every_text_exists_in_spanish_and_english():
    assert dc.TEXT["es"].keys() == dc.TEXT["en"].keys()
