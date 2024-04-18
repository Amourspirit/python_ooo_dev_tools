"""
Two multiples repeated ops examples, used to fill
table from starting values.
Based on code in Dev Guide's SpreadSheetSample.java example.
See:
    https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Multiple_Operations
"""

from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from typing import cast, TYPE_CHECKING
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc

if TYPE_CHECKING:

    from com.sun.star.sheet import FillDirection as UnoFillDirection  # enum not a legal uno import
    from com.sun.star.sheet import TableOperationMode as UnoTableOperationMode  # enum not a legal uno import


def test_multiples(loader) -> None:
    from com.sun.star.sheet import XMultipleOperation
    from ooodev.utils.uno_enum import UnoEnum

    FillDirection = cast("UnoFillDirection", UnoEnum("com.sun.star.sheet.FillDirection"))
    TableOperationMode = cast("UnoTableOperationMode", UnoEnum("com.sun.star.sheet.TableOperationMode"))
    # cast using string as uno enum's are not a legal import in python
    doc = Calc.create_doc(loader=loader)
    visible = False
    delay = 0  # 2000
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)

    # fill top-left corner
    Calc.set_val(value="=A5^B4", sheet=sheet, cell_name="A4")
    Calc.set_val(value=1, sheet=sheet, cell_name="A5")
    Calc.set_val(value=1, sheet=sheet, cell_name="B4")

    # fill down left edge and to the right along top
    Calc.get_cell_series(sheet=sheet, range_name="A5:A9").fillAuto(FillDirection.TO_BOTTOM, 1)
    Calc.get_cell_series(sheet=sheet, range_name="B4:F4").fillAuto(FillDirection.TO_RIGHT, 1)

    # specify where the formula is located
    formula_range = Calc.get_address(sheet=sheet, range_name="A4")

    col_cell = Calc.get_cell_address(sheet=sheet, cell_name="A5")  # Starting col cell
    row_cell = Calc.get_cell_address(sheet=sheet, cell_name="B4")  # staring row cell

    cell_range = Calc.get_cell_range(sheet=sheet, range_name="A4:F9")  # multiple ops data

    mult_op = Lo.qi(XMultipleOperation, cell_range)
    # fill the table along both axes
    mult_op.setTableOperation(formula_range, TableOperationMode.BOTH, col_cell, row_cell)

    # a row of trig functions
    Calc.set_val(value="=SIN(A11)", sheet=sheet, cell_name="B11")
    Calc.set_val(value="=COS(A11)", sheet=sheet, cell_name="C11")
    Calc.set_val(value="=TAN(A11)", sheet=sheet, cell_name="D11")

    arr = Calc.get_float_array(sheet=sheet, range_name="A4:F9")
    assert arr[0][0] == 1.0
    assert arr[0][1] == 1.0
    assert arr[0][2] == 2.0
    assert arr[0][3] == 3.0
    assert arr[0][4] == 4.0
    assert arr[0][5] == 5.0

    assert arr[1][0] == 1.0
    assert arr[1][1] == 1.0
    assert arr[1][2] == 1.0
    assert arr[1][3] == 1.0
    assert arr[1][4] == 1.0
    assert arr[1][5] == 1.0

    assert arr[2][0] == 2.0
    assert arr[2][1] == 2.0
    assert arr[2][2] == 4.0
    assert arr[2][3] == 8.0
    assert arr[2][4] == 16.0
    assert arr[2][5] == 32.0

    assert arr[3][0] == 3.0
    assert arr[3][1] == 3.0
    assert arr[3][2] == 9.0
    assert arr[3][3] == 27.0
    assert arr[3][4] == 81.0
    assert arr[3][5] == 243.0

    assert arr[4][0] == 4.0
    assert arr[4][1] == 4.0
    assert arr[4][2] == 16.0
    assert arr[4][3] == 64.0
    assert arr[4][4] == 256.0
    assert arr[4][5] == 1024.0

    assert arr[5][0] == 5.0
    assert arr[5][1] == 5.0
    assert arr[5][2] == 25.0
    assert arr[5][3] == 125.0
    assert arr[5][4] == 625.0
    assert arr[5][5] == 3125.0
    # Lo.delay(delay)

    # initial values on the left, going down
    Calc.set_val(value=0, sheet=sheet, cell_name="A12")
    Calc.set_val(value=0.2, sheet=sheet, cell_name="A13")

    # finish off by filling down
    Calc.get_cell_series(sheet=sheet, range_name="A12:A16").fillAuto(FillDirection.TO_BOTTOM, 2)

    # specify where the formulas are located
    formula_range = Calc.get_address(sheet=sheet, range_name="B11:D11")

    col_cell = Calc.get_cell_address(sheet=sheet, cell_name="A11")  # rowCell not needed

    cell_range = Calc.get_cell_range(sheet=sheet, range_name="A12:D16")

    mult_op = Lo.qi(XMultipleOperation, cell_range)
    mult_op.setTableOperation(formula_range, TableOperationMode.COLUMN, col_cell, row_cell)

    Calc.highlight_range(sheet=sheet, headline=" Two Multiple Ops Examples", range_name="A3:F16")
    # a row of trig functions
    arr = Calc.get_float_array(sheet=sheet, range_name="A12:D16")
    assert arr[0][0] == 0.0
    assert arr[0][1] == 0.0
    assert arr[0][2] == 1.0
    assert arr[0][3] == 0.0

    assert arr[1][0] == pytest.approx(0.2, rel=1e-2)
    assert arr[1][1] == pytest.approx(0.19866933079506122, rel=1e-4)
    assert arr[1][2] == pytest.approx(0.9800665778412416, rel=1e-4)
    assert arr[1][3] == pytest.approx(0.2027100355086725, rel=1e-4)

    assert arr[2][0] == pytest.approx(0.4, rel=1e-2)
    assert arr[2][1] == pytest.approx(0.3894183423086505, rel=1e-4)
    assert arr[2][2] == pytest.approx(0.9210609940028851, rel=1e-4)
    assert arr[2][3] == pytest.approx(0.4227932187381618, rel=1e-4)

    assert arr[3][0] == pytest.approx(0.6, rel=1e-2)
    assert arr[3][1] == pytest.approx(0.5646424733950355, rel=1e-4)
    assert arr[3][2] == pytest.approx(0.8253356149096782, rel=1e-4)
    assert arr[3][3] == pytest.approx(0.6841368083416924, rel=1e-4)

    assert arr[4][0] == pytest.approx(0.8, rel=1e-2)
    assert arr[4][1] == pytest.approx(0.7173560908995228, rel=1e-4)
    assert arr[4][2] == pytest.approx(0.6967067093471654, rel=1e-4)
    assert arr[4][3] == pytest.approx(1.0296385570503641, rel=1e-4)
    Lo.delay(delay)
    Lo.close(closeable=doc, deliver_ownership=False)
