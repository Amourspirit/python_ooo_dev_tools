from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc, GeneralFunction

from com.sun.star.sheet import CellFlags  # const
from com.sun.star.sheet import XArrayFormulaRange
from com.sun.star.sheet import XCellRangeData
from com.sun.star.sheet import XCellRangesQuery
from com.sun.star.sheet import XSheetCellRange
from com.sun.star.sheet import XSheetOperation
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheetDocument
from com.sun.star.sheet import XUsedAreaCursor


def test_cell_ranges(loader) -> None:
    doc = Calc.create_doc(loader=loader)
    visible = False
    delay = 0  # 2000
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc)

    try:
        do_cell_range(sheet=sheet)
        do_cell_cursor(sheet=sheet)
        do_cell_selection(doc=doc, sheet=sheet)
        Lo.delay(delay)
    finally:
        Lo.close(doc)


def do_cell_range(sheet: XSpreadsheet) -> None:
    Calc.highlight_range(sheet=sheet, headline="Range Data Example", range_name="A2:C23")
    vals = (
        ("Name", "Fruit", "Quantity"),
        ("Alice", "Apples", 3),
        ("Alice", "Oranges", 7),
        ("Bob", "Apples", 3),
        ("Alice", "Apples", 9),
        ("Bob", "Apples", 5),
        ("Bob", "Oranges", 6),
        ("Alice", "Oranges", 3),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 1),
        ("Bob", "Oranges", 2),
        ("Bob", "Oranges", 7),
        ("Bob", "Apples", 1),
        ("Alice", "Apples", 8),
        ("Alice", "Oranges", 8),
        ("Alice", "Apples", 7),
        ("Bob", "Apples", 1),
        ("Bob", "Oranges", 9),
        ("Bob", "Oranges", 3),
        ("Alice", "Oranges", 4),
        ("Alice", "Apples", 9),
    )
    range_name = "A3:C23"
    Calc.set_array(values=vals, sheet=sheet, name=range_name)  # or just "A3"

    cell_range = Calc.get_cell_range(sheet=sheet, range_name=range_name)
    cr_data = Lo.qi(XCellRangeData, cell_range)
    # Calc.print_address(cell_range)

    # Sheet operation using the range of the crData
    sheet_op = Lo.qi(XSheetOperation, cr_data)
    avg = sheet_op.computeFunction(GeneralFunction.AVERAGE)
    assert avg == 5.2

    # Array formulas
    range_rows = Calc.get_cell_range(sheet=sheet, range_name="E3:G5")
    Calc.highlight_range(sheet=sheet, headline=" Array Formula Example", range_name="E2:G5")
    af_range = Lo.qi(XArrayFormulaRange, range_rows)
    # Insert a 3x3 unit matrix
    af_range.setArrayFormula("=A3:C5")
    assert af_range.getArrayFormula() == "{=A3:C5}"

    # Cell Ranges Query
    cr_query = Lo.qi(XCellRangesQuery, cell_range)
    cell_ranges = cr_query.queryContentCells(CellFlags.STRING)
    rng_addr_str = cell_ranges.getRangeAddressesAsString()
    assert rng_addr_str == "Sheet1.A3:B23,Sheet1.C3"


def do_cell_cursor(sheet: XSpreadsheet) -> None:
    # Find the array formula using a cell cursor
    xrange = Calc.get_cell_range(sheet=sheet, range_name="E4")
    cell_range = Lo.qi(XSheetCellRange, xrange)
    cursor = sheet.createCursorByRange(cell_range)
    cursor.collapseToCurrentArray()

    xarray = Lo.qi(XArrayFormulaRange, cursor)
    rng_str = Calc.get_range_str(cell_range)
    arr_formula = xarray.getArrayFormula()
    assert rng_str == "E4:E4"
    assert arr_formula == "{=A3:C5}"

    # Find the used area
    ua_cursor = Lo.qi(XUsedAreaCursor, cursor)
    ua_cursor.gotoStartOfUsedArea(False)
    ua_cursor.gotoEndOfUsedArea(True)

    # ua_cursor and cursor are interfaces of the same object -
    # so modifying ua_cursor takes effect on cursor and its cell_range:
    ua_str = Calc.get_range_str(cell_range=cell_range)
    assert ua_str == "E4:E4"


def do_cell_selection(doc: XSpreadsheetDocument, sheet: XSpreadsheet) -> None:
    range_name = "A3:C23"
    # test cell selection
    addr = Calc.set_selected_addr(doc=doc, sheet=sheet, range_name=range_name)
    assert Calc.get_range_str(addr) == range_name
    # test deselection
    # when deselection addr will be the current selected cell
    addr = Calc.set_selected_addr(doc=doc, sheet=sheet)
    assert addr.StartColumn == addr.EndColumn
    assert addr.StartRow == addr.EndRow
