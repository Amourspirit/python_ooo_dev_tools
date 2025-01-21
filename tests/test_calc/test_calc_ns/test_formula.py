from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])


def test_text_formula(loader) -> None:
    from ooodev.calc import CalcDoc
    from ooodev.utils.data_type.range_obj import RangeObj

    # https://ask.libreoffice.org/t/how-to-get-the-results-of-a-formula-for-text-based-formulas/116938/6
    # Note that formula args need to be separated by semi-colon. Not comma.
    # The semi-colon is used as the argument separator in the formula, internationally.

    doc = None
    try:
        doc = CalcDoc.create_doc(loader=loader)
        sheet = doc.sheets[0]
        cell = sheet["A1"]
        cell.value = 10
        assert cell.value == 10

        cell = sheet["B1"]
        cell.value = 20
        assert cell.value == 20

        cell = sheet["C1"]
        cell.value = "=SUM(A1:B1)"
        # cell.value = '=TEXT(156,"0.00")'
        # ro = RangeObj.from_range("A1")
        # arr = sheet.get_array(range_name="A1:A1")
        assert cell.value == pytest.approx(30)

        cell = sheet["A2"]
        cell.value = "test"
        assert cell.value == "test"

        cell = sheet["A3"]
        cell.value = "Hello"

        cell = sheet["B3"]
        cell.value = "World"

        cell = sheet["C3"]
        cell.value = '=CONCATENATE(A3;" ";B3)'
        assert cell.value == "Hello World"

        cell = sheet["A3"]

        call_result = doc.call_fun("TEXT", 156, "0.00")
        assert call_result == "156.00"
        cell.value = '=TEXT(156;"0.00")'
        assert cell.value == "156.00"

        cell = sheet["A4"]
        cell.value = '="good " & "morning"'
        assert cell.value == "good morning"

    finally:
        if doc is not None:
            doc.close()


def test_doc_text_formula(copy_fix_calc, loader) -> None:
    from ooodev.calc import CalcDoc

    # https://ask.libreoffice.org/t/how-to-get-the-results-of-a-formula-for-text-based-formulas/116938/6
    # Note that formula args need to be separated by semi-colon. Not comma.
    # The semi-colon is used as the argument separator in the formula, internationally.

    doc_path = copy_fix_calc("small_totals.ods")
    doc = None
    try:
        doc = CalcDoc.open_doc(fnm=doc_path, loader=loader)
        sheet = doc.sheets[2]  # formulas
        cell = sheet["A1"]
        assert cell.value == "156.00"  # =TEXT(156,"0.00")

        cell = sheet["A2"]
        assert cell.value == "good morning"  # ="good " & "morning"

        cell = sheet["C3"]
        assert cell.value == "HelloWorld"  # =CONCATENATE(A3,B3)

    finally:
        if doc is not None:
            doc.close()
