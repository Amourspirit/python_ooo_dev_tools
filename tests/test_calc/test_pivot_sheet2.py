from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.props import Props
from ooodev.office.calc import Calc

from com.sun.star.sheet import XDataPilotTables, XDataPilotTable
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet.DataPilotFieldOrientation import COLUMN as FO_COLUMN, ROW as FO_ROW, DATA as FO_DATA, PAGE as FO_PAGE

def test_pivot(copy_fix_calc, loader) -> None:
    doc_path = copy_fix_calc("pivottable2.ods")
    doc = Calc.open_doc(fnm=str(doc_path), loader=loader)
    assert doc is not None, "Could not open pivottable2.ods"
    visible = False
    delay = 0
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)
    dp_sheet = Calc.insert_sheet(doc=doc, name="Pivot Table", idx=1)
    ptbl = create_pivot_table(sheet=sheet, dp_sheet=dp_sheet)
    assert ptbl is not None
    assert ptbl.Name == 'PivotTableExample'
    Calc.set_active_sheet(doc=doc, sheet=dp_sheet)
    total = Calc.get_num(sheet=dp_sheet, cell_name="E10")
    assert total == pytest.approx(7844.0, rel=1e-4)
    Lo.save_doc(doc=doc, fnm=str(doc_path))
    Lo.delay(delay)
    Lo.close_doc(doc=doc)

def create_pivot_table(sheet: XSpreadsheet, dp_sheet: XSpreadsheet) -> XDataPilotTable | None:
    cell_range = Calc.find_used_range(sheet)
    print(f"The used area is: { Calc.get_range_str(cell_range)}")
    print()
    
    dp_tables = Calc.get_pilot_tables(sheet)
    dp_desc = dp_tables.createDataPilotDescriptor()
    dp_desc.setSourceRange(Calc.get_address(cell_range))
    
    # XIndexAccess fields = dpDesc.getDataPilotFields();
    fields = dp_desc.getHiddenFields()
    field_names = Lo.get_container_names(con=fields)
    print(f"Field Names ({len(field_names)}):")
    for name in field_names:
        print(f"  {name}")
    
    # properties defined in DataPilotField
    # set page field
    props = Lo.find_container_props(con=fields, nm="Date")
    Props.set_property(prop_set=props, name="Orientation", value=FO_PAGE)
    
    # set column field
    props = Lo.find_container_props(con=fields, nm="Store")
    Props.set_property(prop_set=props, name="Orientation", value=FO_COLUMN)
    
    # set 1st row field
    props = Lo.find_container_props(con=fields, nm="Book")
    Props.set_property(prop_set=props, name="Orientation", value=FO_ROW)
    
    
    # set data field, calculating the sum
    props = Lo.find_container_props(con=fields, nm="Units Sold")
    Props.set_property(prop_set=props, name="Orientation", value=FO_DATA)
    Props.set_property(prop_set=props, name="Function", value=Calc.GeneralFunction.SUM)
    
    
    
    # place onto sheet
    dest_addr = Calc.get_cell_address(sheet=dp_sheet, cell_name="A1")
    dp_tables.insertNewByName("PivotTableExample", dest_addr, dp_desc)
    Calc.set_col_width(sheet=dp_sheet, width=60, idx=0);   # A column; in mm

    # Usually the table is not fully updated. The cells are often
    # drawn with #VALUE! contents (?).

    # This can be fixed by explicitly refreshing the table, but it has to
    # be accessed via the sheet or the tables container is considered
    # empty, and the table is not found.
    
    dp_tables2 = Calc.get_pilot_tables(sheet=dp_sheet)
    return refresh_table(dp_tables=dp_tables2, table_name="PivotTableExample")

def refresh_table(dp_tables: XDataPilotTables, table_name: str) -> XDataPilotTable | None:
    nms = dp_tables.getElementNames()
    print(f"No. of DP tables: {len(nms)}")
    for nm in nms:
        print(f"  {nm}")
    
    dp_table = Calc.get_pilot_table(dp_tables=dp_tables, name=table_name)
    if dp_table is not None:
        dp_table.refresh()
    return dp_table