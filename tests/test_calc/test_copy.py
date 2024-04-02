from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.props import Props
from ooodev.office.calc import Calc


def test_copy(loader) -> None:
    from com.sun.star.sheet import XCellRangeMovement
    from com.sun.star.sheet import XSheetPageBreak
    from com.sun.star.util import Date

    doc = Calc.create_doc(loader=loader)
    assert doc is not None, "Could not create new document"
    visible = False
    delay = 0  # 1_000
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.insert_sheet(doc=doc, name="A new sheet", idx=0x7FFF)
    Calc.set_active_sheet(doc, sheet)

    # Copy a cell range
    Calc.highlight_range(sheet=sheet, cell_range="A1:B3", headline="Copy from")
    Calc.highlight_range(sheet=sheet, cell_range="D1:E3", headline="To")
    Calc.set_val(sheet=sheet, cell_name="A2", value=123)
    Calc.set_val(sheet=sheet, cell_name="B2", value=345)
    Calc.set_val(sheet=sheet, cell_name="A3", value="=SUM(A2:B2)")
    Calc.set_val(sheet=sheet, cell_name="B3", value="=FORMULA(A3)")

    cr_move = Lo.qi(XCellRangeMovement, sheet)
    source_range = Calc.get_address(sheet=sheet, range_name="A2:B3")
    dest_cell = Calc.get_cell_address(sheet=sheet, cell_name="D2")
    cr_move.copyRange(dest_cell, source_range)

    c_vals = Calc.get_array(sheet=sheet, range_name="D2:E3")
    assert c_vals[0][0] == pytest.approx(123.0, rel=1e-4)
    assert c_vals[0][1] == pytest.approx(345.0, rel=1e-4)
    assert c_vals[1][0] == pytest.approx(468.0, rel=1e-4)
    assert c_vals[1][1] == "=SUM(D2:E2)"

    # automatic column page breaks
    sp_break = Lo.qi(XSheetPageBreak, sheet)
    page_breaks = sp_break.getColumnPageBreaks()
    break_arr = []
    for page_break in page_breaks:
        if not page_break.ManualBreak:
            break_arr.append(page_break)
    assert len(break_arr) > 0  # length is 9

    IsIterationEnabled = Props.get_property(doc, "IsIterationEnabled")
    assert IsIterationEnabled == False
    IterationCount = Props.get_property(doc, "IterationCount")
    assert IterationCount == 100
    NullDate: Date = Props.get_property(doc, "NullDate")
    assert NullDate.Day == 30
    assert NullDate.Month == 12
    assert NullDate.Year == 1899

    Lo.delay(delay)
    Lo.close(closeable=doc, deliver_ownership=False)
