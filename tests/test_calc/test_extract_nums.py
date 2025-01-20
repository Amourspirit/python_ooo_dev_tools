from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc


def test_extract_small_totals(copy_fix_calc, loader, capsys: pytest.CaptureFixture) -> None:
    doc_path = copy_fix_calc("small_totals.ods")
    doc = Calc.open_doc(fnm=str(doc_path), loader=loader)
    visible = False
    delay = 0
    if visible:
        GUI.set_visible(visible=visible, doc=doc)
    sheet = Calc.get_sheet(doc=doc, idx=0)

    try:
        # basic data extraction
        assert Calc.get_val(sheet=sheet, cell_name="A1") == "Stud. No."

        cell = Calc.get_cell(sheet, "A2")
        a2_type = Calc.get_type_string(cell)
        a2_value = Calc.get_num(cell)
        assert a2_type == "VALUE"
        assert a2_value == pytest.approx(22001.0, rel=1e-2)

        cell = Calc.get_cell(sheet=sheet, cell_name="E2")
        e2_type = Calc.get_type_string(cell)
        e2_value = Calc.get_val(sheet=sheet, cell_name="E2")
        assert e2_type == "FORMULA"
        # in version 0.50.1 get_val() returns the value of the formula, previously it was the formula string
        # assert e2_value == "=SUM(B2:D2)/100"
        assert e2_value == pytest.approx(0.843875, rel=1e-7)

        data = Calc.get_array(sheet, "A1:E10")
        assert len(data) == 10
        capsys.readouterr()  # clear buffer
        Calc.print_array(data)
        captured = capsys.readouterr()
        cstr: str = captured.out
        clst = cstr.splitlines()
        assert clst[0] == "Row x Column size: 10 x 5"
        assert clst[1] == "Stud. No.  Proj/20  Mid/35  Fin/45  Total%"
        assert clst[2].startswith("22001.0")
        assert clst[10].strip() == "Proj/20  Mid/35  Fin/45  Total%"
        # Row x Column size: 10 x 5
        # Stud. No.  Proj/20  Mid/35  Fin/45  Total%
        # 22001.0  16.4583333333333  30.9166666666667  37.0125  0.843875
        # 22028.0  11.875  23.0416666666667  25.4625  0.603791666666667
        # 22048.0  13.9583333333333  19.25  25.9875  0.591958333333333
        # 23715.0  12.0833333333333  18.6666666666667  20.475  0.51225
        # 23723.0  17.2916666666667  27.7083333333333  36.225  0.81225
        # 24277.0  0.0  16.0416666666667  19.6875  0.357291666666667
        #   11.9444444444444  22.6041666666667  27.475  0.620236111111111
        #   0.597222222222222  0.645833333333334  0.610555555555556
        #   Proj/20  Mid/35  Fin/45  Total%

        ids = Calc.get_float_array(sheet=sheet, range_name="A2:A7")
        capsys.readouterr()  # clear buffer
        Calc.print_array(ids)
        captured = capsys.readouterr()
        expected = """Row x Column size: 6 x 1
22001.0
22028.0
22048.0
23715.0
23723.0
24277.0

"""
        assert captured.out == expected
        # use convert_to_doubles insetead of convert_to_floats
        # convert_to_floats is tested in test_small_totals.py
        projs = Calc.convert_to_doubles(Calc.get_col(sheet=sheet, range_name="B2:B7"))
        assert projs[0] == pytest.approx(16.5, rel=1e-2)
        assert projs[1] == pytest.approx(11.9, rel=1e-2)
        assert projs[2] == pytest.approx(14.0, rel=1e-2)
        assert projs[3] == pytest.approx(12.1, rel=1e-2)
        assert projs[4] == pytest.approx(17.3, rel=1e-2)
        assert projs[5] == pytest.approx(0.0, rel=1e-2)

        stud = Calc.convert_to_doubles(Calc.get_row(sheet, "A4:E4"))
        assert stud[0] == pytest.approx(22048.0, rel=1e-2)
        assert stud[1] == pytest.approx(14.0, rel=1e-2)
        assert stud[2] == pytest.approx(19.3, rel=1e-2)
        assert stud[3] == pytest.approx(26.0, rel=1e-2)
        assert stud[4] == pytest.approx(0.59, rel=1e-2)

        Lo.delay(delay)
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)  # type: ignore
