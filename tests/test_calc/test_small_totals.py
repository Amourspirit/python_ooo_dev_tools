from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from thefuzz import fuzz
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc


def test_small_totals(copy_fix_calc, loader, capsys: pytest.CaptureFixture) -> None:
    from com.sun.star.sheet import XCellRangesQuery
    from com.sun.star.sheet import CellFlags  # const

    doc_path = copy_fix_calc("small_totals.ods")
    doc = Calc.open_doc(fnm=str(doc_path), loader=loader)
    assert doc is not None, "Could not open small_totals.ods"
    visible = False
    delay = 0
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)
    val = Calc.get_val(sheet=sheet, cell_name="A1")
    assert val == "Stud. No."

    cell = Calc.get_cell(sheet=sheet, cell_name="A2")
    type_str = Calc.get_type_string(cell=cell)
    type_enum = Calc.get_type_enum(cell=cell)
    assert type_str == "VALUE"
    assert type_enum == Calc.CellTypeEnum.VALUE
    val = Calc.get_num(cell)  # 22001
    assert val == pytest.approx(22001.0, rel=1e-4)

    cell = Calc.get_cell(sheet=sheet, cell_name="E2")
    type_str = Calc.get_type_string(cell=cell)
    type_enum = Calc.get_type_enum(cell=cell)
    assert type_str == "FORMULA"
    assert type_enum == Calc.CellTypeEnum.FORMULA
    val = Calc.get_val(sheet=sheet, cell_name="E2")
    assert val == "=SUM(B2:D2)/100"

    data = Calc.get_array(sheet=sheet, range_name="A1:E10")
    capsys.readouterr()  # clear buffer
    Calc.print_array(data)
    captured = capsys.readouterr()
    expected = """Row x Column size: 10 x 5
Stud. No.  Proj/20  Mid/35  Fin/45  Total%
22001.0  16.4583333333333  30.9166666666667  37.0125  0.843875
22028.0  11.875  23.0416666666667  25.4625  0.603791666666667
22048.0  13.9583333333333  19.25  25.9875  0.591958333333333
23715.0  12.0833333333333  18.6666666666667  20.475  0.51225
23723.0  17.2916666666667  27.7083333333333  36.225  0.81225
24277.0  0.0  16.0416666666667  19.6875  0.357291666666667
  11.9444444444444  22.6041666666667  27.475  0.620236111111111
  0.597222222222222  0.645833333333334  0.610555555555556  
  Proj/20  Mid/35  Fin/45  Total%

"""
    ratio = fuzz.ratio(captured.out, expected)
    assert ratio > 90

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

    projs = Calc.convert_to_floats(Calc.get_col(sheet=sheet, range_name="B2:B7"))
    assert projs[0] == pytest.approx(16.5, rel=1e-2)
    assert projs[1] == pytest.approx(11.9, rel=1e-2)
    assert projs[2] == pytest.approx(14.0, rel=1e-2)
    assert projs[3] == pytest.approx(12.1, rel=1e-2)
    assert projs[4] == pytest.approx(17.3, rel=1e-2)
    assert projs[5] == pytest.approx(0.0, rel=1e-2)

    stud = Calc.convert_to_floats(Calc.get_row(sheet, "A4:E4"))
    assert stud[0] == pytest.approx(22048.0, rel=1e-2)
    assert stud[1] == pytest.approx(14.0, rel=1e-2)
    assert stud[2] == pytest.approx(19.3, rel=1e-2)
    assert stud[3] == pytest.approx(26.0, rel=1e-2)
    assert stud[4] == pytest.approx(0.59, rel=1e-2)

    # create a cell range that spans the used area of the sheet
    used_cell_range = Calc.find_used_range(sheet)
    rng_str = Calc.get_range_str(cell_range=used_cell_range)
    assert rng_str == "A1:E10"

    # find cell ranges that cover all the specified data types
    cr_query = Lo.qi(XCellRangesQuery, used_cell_range)
    cell_ranges = cr_query.queryContentCells(CellFlags.VALUE)
    assert cell_ranges is not None
    assert cell_ranges.getRangeAddressesAsString() == "Marks.A2:D7"
    Lo.delay(delay)
    Lo.close(closeable=doc, deliver_ownership=False)
