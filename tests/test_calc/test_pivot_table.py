from __future__ import annotations
import pytest
from typing import cast, TYPE_CHECKING

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.utils.uno_util import UnoEnum
from ooodev.office.calc import Calc

from com.sun.star.sheet import XDataPilotTable
from com.sun.star.sheet import XSpreadsheet

if TYPE_CHECKING:
    # uno enums are not able to be imported at runtime
    from com.sun.star.sheet import DataPilotFieldOrientation as UnoDataPilotFieldOrientation  # enum


def test_pivot(loader) -> None:
    doc = Calc.create_doc(loader=loader)
    visible = False
    delay = 0  # 3_000
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)
    try:
        create_table(sheet=sheet)
        create_pivot_table(sheet=sheet)

        total = Calc.get_num(sheet=sheet, cell_name="H9")
        assert total == pytest.approx(104.0, rel=1e-4)
        Lo.delay(delay)
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)


def create_table(sheet: XSpreadsheet) -> XDataPilotTable | None:
    Calc.highlight_range(sheet=sheet, range_name="A1:C22", headline="Data Used by Pivot Table")
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

    Calc.set_array(sheet=sheet, name="A2:C22", values=vals)


def create_pivot_table(sheet: XSpreadsheet) -> XDataPilotTable | None:
    # cast and wrap in quotes as uno enums can not be imported in this fashion. Give typing support
    DPFO = cast("UnoDataPilotFieldOrientation", UnoEnum("com.sun.star.sheet.DataPilotFieldOrientation"))
    Calc.highlight_range(sheet=sheet, range_name="E1:H1", headline="Pivot Table")
    Calc.set_col_width(sheet=sheet, width=40, idx=4)
    # E column; in mm

    dp_tables = Calc.get_pilot_tables(sheet)
    dp_desc = dp_tables.createDataPilotDescriptor()

    # set source range (use data range from CellRange test)
    src_addr = Calc.get_address(sheet=sheet, range_name="A2:C22")
    dp_desc.setSourceRange(src_addr)

    # settings for fields
    fields = dp_desc.getDataPilotFields()

    # properties defined in DataPilotField
    # use first column as column field

    props = Lo.find_container_props(con=fields, nm="Name")
    Props.set_property(prop_set=props, name="Orientation", value=DPFO.COLUMN)

    # use second column as row field
    props = Lo.find_container_props(con=fields, nm="Fruit")
    Props.set_property(prop_set=props, name="Orientation", value=DPFO.ROW)

    # use third column as data field, calculating the sum
    props = Lo.find_container_props(con=fields, nm="Quantity")
    Props.set_property(prop_set=props, name="Orientation", value=DPFO.DATA)
    Props.set_property(prop_set=props, name="Function", value=Calc.GeneralFunction.SUM)

    # select output position
    dest_addr = Calc.get_cell_address(sheet=sheet, cell_name="E3")
    dp_tables.insertNewByName("DataPilotExample", dest_addr, dp_desc)
