from __future__ import annotations
from typing import cast, TYPE_CHECKING
import pytest
if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.uno_util import UnoEnum
from ooodev.office.calc import Calc
if TYPE_CHECKING:
    # Enums are not valid uno import. Guard with type checking
    from com.sun.star.sheet import FillDirection as UnoFillDirection # enum
    from com.sun.star.sheet import FillMode as UnoFillMode # enum
    from com.sun.star.sheet import FillDateMode as UnoFillDateMode # enum



def test_fill_series(loader) -> None:
    # from com.sun.star.sheet.FillDirection import TO_RIGHT, TO_LEFT, TO_TOP
    # from com.sun.star.sheet.FillMode import LINEAR as FM_LINEAR, DATE as FM_DATE, AUTO as FM_AUTO, GROWTH as FM_GROWTH
    # from com.sun.star.sheet.FillDateMode import FILL_DATE_MONTH
    #
    # Enums are not valid uno imports, for typing support cast and wrap import in quotes
    FillDirection = cast("UnoFillDirection", UnoEnum(type_name="com.sun.star.sheet.FillDirection"))
    FillMode = cast("UnoFillMode", UnoEnum(type_name="com.sun.star.sheet.FillMode"))
    FillDateMode = cast("UnoFillDateMode", UnoEnum(type_name="com.sun.star.sheet.FillDateMode"))

    doc = Calc.create_doc(loader=loader)
    visible = False
    delay = 0 # 1000
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)

    # set first two values of three rows
    Calc.set_val(sheet=sheet, cell_name="B7", value=2)
    Calc.set_val(sheet=sheet, cell_name="A7", value=1)    # 1. ascending

    Calc.set_date(sheet=sheet, cell_name="A8", day=28, month=2, year=2015)   #2. dates, descending
    Calc.set_date(sheet=sheet, cell_name="B8", day=28, month=1, year=2015)

    Calc.set_val(sheet=sheet, cell_name="A9", value=6)     # 3. descending
    Calc.set_val(sheet=sheet, cell_name="B9", value=4)
    Lo.delay(delay)
    #Autofill, using first 2 cells to right to determine progressions
    series = Calc.get_cell_series(sheet=sheet, range_name="A7:G9")
    series.fillAuto(FillDirection.TO_RIGHT, 2)
    arr = Calc.get_array(sheet=sheet, range_name="A7:G9")
    assert arr[0][2] == pytest.approx(3, rel=1e-2)
    assert arr[0][3] == pytest.approx(4, rel=1e-2)
    assert arr[0][4] == pytest.approx(5, rel=1e-2)
    assert arr[0][5] == pytest.approx(6, rel=1e-2)
    assert arr[0][6] == pytest.approx(7, rel=1e-2)
    
    assert arr[1][2] == 42001.0
    assert arr[1][3] == 41971.0
    assert arr[1][4] == 41940.0
    assert arr[1][5] == 41910.0
    assert arr[1][6] == 41879.0
    
    assert arr[2][2] == pytest.approx(2, rel=1e-2)
    assert arr[2][3] == pytest.approx(0, rel=1e-2)
    assert arr[2][4] == pytest.approx(-2, rel=1e-2)
    assert arr[2][5] == pytest.approx(-4, rel=1e-2)
    assert arr[2][6] == pytest.approx(-6, rel=1e-2)
    Lo.delay(delay)

    # ----------------------------------------
    Calc.set_val(sheet=sheet, cell_name="A2", value=1)
    Calc.set_val(sheet=sheet, cell_name="A3", value=4)
    Lo.delay(delay)
    # Fill 2 rows; 2nd row is not filled completely since
    # the end value is reached
    series = Calc.get_cell_series(sheet=sheet, range_name="A2:E3")
    series.fillSeries(FillDirection.TO_RIGHT, FillMode.LINEAR, Calc.NO_DATE, 2, 9)
                #   ignore date mode; step == 2; end at 9
    arr = Calc.get_array(sheet=sheet, range_name="A2:E3")
    assert arr[0][1] == 3.0
    assert arr[0][2] == 5.0
    assert arr[0][3] == 7.0
    assert arr[0][4] == 9.0
    
    assert arr[1][1] == 6.0
    assert arr[1][2] == 8.0
    assert arr[1][3] == ""
    assert arr[1][4] == ""
    Lo.delay(delay)
    # ----------------------------------------
    Calc.set_date(sheet=sheet, cell_name="A4", day=20, month=11, year=2015)    # day, month, year
    Lo.delay(delay)
    # fill by adding one month to date; day is unchanged
    series = Calc.get_cell_series(sheet=sheet, range_name="A4:E4")
    series.fillSeries(FillDirection.TO_RIGHT, FillMode.DATE, FillDateMode.FILL_DATE_MONTH, 1, Calc.MAX_VALUE)
    arr = Calc.get_array(sheet=sheet, range_name="A4:E4")
    assert arr[0][0] == 42328.0
    assert arr[0][1] == 42358.0
    assert arr[0][2] == 42389.0
    assert arr[0][3] == 42420.0
    assert arr[0][4] == 42449.0
    Lo.delay(delay)
    # ----------------------------------------
    Calc.set_val(sheet=sheet, cell_name="E5", value="Text 10")   # start in the middle of a row
    Lo.delay(delay)
    # Fill from right to left with text+value in steps of +10
    series = Calc.get_cell_series(sheet=sheet, range_name="A5:E5")
    series.fillSeries(FillDirection.TO_LEFT, FillMode.LINEAR, Calc.NO_DATE, 10, Calc.MAX_VALUE)
    arr = Calc.get_array(sheet=sheet, range_name="A5:E5")
    assert arr[0][0] == 'Text 50'
    assert arr[0][1] == 'Text 40'
    assert arr[0][2] == 'Text 30'
    assert arr[0][3] == 'Text 20'
    assert arr[0][4] == 'Text 10'
    Lo.delay(delay)
    # ----------------------------------------
    Calc.set_val(sheet=sheet, cell_name="A6", value="Jan")
    Lo.delay(delay)
    # Fill with values generated automatically from first entry
    series = Calc.get_cell_series(sheet=sheet, range_name="A6:E6")
    series.fillSeries(FillDirection.TO_RIGHT, FillMode.AUTO, Calc.NO_DATE, 1, Calc.MAX_VALUE)
    # series.fillAuto(TO_RIGHT, 1)  # does the same
    arr = Calc.get_array(sheet=sheet, range_name="A6:E6")
    assert arr[0][0] == 'Jan'
    assert arr[0][1] == 'Feb'
    assert arr[0][2] == 'Mar'
    assert arr[0][3] == 'Apr'
    assert arr[0][4] == 'May'
    Lo.delay(delay)
    # ----------------------------------------
    Calc.set_val(sheet=sheet, cell_name="G6", value=10)
    
    # Fill from  bottom to top with a geometric series (*2)
    series = Calc.get_cell_series(sheet=sheet, range_name="G2:G6")
    series.fillSeries(FillDirection.TO_TOP, FillMode.GROWTH, Calc.NO_DATE, 2, Calc.MAX_VALUE)
    arr = Calc.get_array(sheet=sheet, range_name="G2:G6")
    assert arr[0][0] == 160.0
    assert arr[1][0] == 80.0
    assert arr[2][0] == 40.0
    assert arr[3][0] == 20.0
    assert arr[4][0] == 10.0
    Lo.delay(delay)
    Lo.close(closeable=doc, deliver_ownership=False)
