from __future__ import annotations
from typing import List
import pytest

if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc

from com.sun.star.table import CellAddress
from com.sun.star.table import CellRangeAddress  # struct
from com.sun.star.sheet import XCellAddressable
from com.sun.star.sheet import XSheetCellRangeContainer
from com.sun.star.sheet import XSpreadsheet


def test_cell_ranges(loader) -> None:
    doc = Calc.create_doc(loader=loader)
    visible = False
    delay = 0  # 2000
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)

    try:
        create_table(sheet=sheet)
        do_cell_ranges()
        Lo.delay(delay)
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)


def create_table(sheet: XSpreadsheet) -> None:
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
    Calc.set_array(vals, sheet, "A10:C30")


def do_cell_ranges() -> None:
    src_con = Lo.create_instance_msf(XSheetCellRangeContainer, "com.sun.star.sheet.SheetCellRanges")
    rng_str = insert_range(src_con, 0, 0, 0, 0, False)
    # A1:A1  -- empty
    assert rng_str == "Sheet1.A1"
    rng_str = insert_range(src_con, 0, 1, 0, 2, True)
    # A2:A3  -- empty
    assert rng_str == "Sheet1.A1:A3"
    rng_str = insert_range(src_con, 1, 0, 1, 2, False)
    # B1:B3   -- empty
    assert rng_str == "Sheet1.A1:A3,Sheet1.B1:B3"

    rng_str = insert_range(src_con, 0, 9, 1, 29, False)
    # A10:B30  -- not empty
    assert rng_str == "Sheet1.A1:A3,Sheet1.B1:B3,Sheet1.A10:B30"
    # Query the list of filled cells
    cells_ea = src_con.getCells()
    cell_enum = cells_ea.createEnumeration()
    addrs: List[CellAddress] = []
    while cell_enum.hasMoreElements():
        xaddr = Lo.qi(XCellAddressable, cell_enum.nextElement())
        addrs.append(xaddr.getCellAddress())

    assert len(addrs) == 42
    assert addrs[0].Column == 0
    assert addrs[0].Row == 9
    assert addrs[41].Column == 1
    assert addrs[41].Row == 29


def insert_range(
    src_con: XSheetCellRangeContainer, start_col: int, start_row: int, end_col: int, end_row: int, is_merged: bool
) -> str:
    """Inserts a cell range address into a cell range container"""
    addr = CellRangeAddress()
    addr.Sheet = 0
    addr.StartColumn = start_col
    addr.StartRow = start_row
    addr.EndColumn = end_col
    addr.EndRow = end_row

    src_con.addRangeAddress(addr, is_merged)
    return src_con.getRangeAddressesAsString()
