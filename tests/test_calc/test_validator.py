"""
Validata data entry so between 0 and 5.

   Based on code in Dev Guide's SpreadSheetSample.java example.
   and see the dev guide "Data Validation" section:
     https://wiki.openoffice.org/w/index.php?title=Documentation/DevGuide/Spreadsheets/Other_Table_Operations
"""
from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast
if __name__ == "__main__":
    pytest.main([__file__])
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.uno_util import UnoEnum
from ooodev.utils.props import Props
from ooodev.office.calc import Calc

from com.sun.star.table import XCellRange
from com.sun.star.beans import XPropertySet
from com.sun.star.sheet import XSheetCondition

if TYPE_CHECKING:
    from com.sun.star.sheet import ValidationType as UnoValidationType # enum
    from com.sun.star.sheet import ValidationAlertStyle as UnoValidationAlertStyle # enum
    from com.sun.star.sheet import ConditionOperator as UnoConditionOperator # enum

def test_validator(loader) -> None:
    # this test is most useful when visible.
    # validation is set on A3, B3 and C3.
    # if input value is not 0.0 to 5.0 then validation fails and a message box is shown.
    doc = Calc.create_doc(loader=loader)
    visible = False
    delay = 0 # 30_000
    if visible:
        GUI.set_visible(is_visible=visible, odoc=doc)
    sheet = Calc.get_sheet(doc=doc, index=0)
    try:
        Calc.highlight_range(sheet=sheet, headline="Validation Example", cell_range="A1:C3")
        Calc.set_val(value="Insert 3 values between 0.0 and 5.0:",sheet=sheet, cell_name="A2")
        
        valid_range = Calc.get_cell_range(sheet=sheet, range_name="A3:C3")
        validate(cell_range=valid_range, start=0.0, end=5.0)
        Lo.delay(delay)
    finally:
        Lo.close(closeable=doc, deliver_ownership=False)

def validate(cell_range:XCellRange, start: float, end: float) -> None:
    # get cell range properties
    ValidationType = cast("UnoValidationType", UnoEnum("com.sun.star.sheet.ValidationType"))
    ValidationAlertStyle = cast("UnoValidationAlertStyle", UnoEnum("com.sun.star.sheet.ValidationAlertStyle"))
    ConditionOperator = cast("UnoConditionOperator", UnoEnum("com.sun.star.sheet.ConditionOperator"))
    cr_props = Lo.qi(XPropertySet, cell_range)
    
    # change validation properties of cell range
    # "Validation" is defined inside SheetCellRange
    vprops = Lo.qi(XPropertySet, cr_props.getPropertyValue("Validation"))
    Props.show_props("Validation", vprops)
        # see "Data Validation" in the Dev Guid
        # these props are defined in the TableValidation service
    
    vprops.setPropertyValue("Type", ValidationType.DECIMAL)
    vprops.setPropertyValue("ShowErrorMessage", True)
    vprops.setPropertyValue("ErrorMessage", "This is an invalid value!")
    vprops.setPropertyValue("ErrorAlertStyle", ValidationAlertStyle.STOP)
    
    # set condition
    cond = Lo.qi(XSheetCondition, vprops)
    cond.setOperator(ConditionOperator.BETWEEN)
    cond.setFormula1(f"{start}")
    cond.setFormula2(f"{end}")
    
    # store updated validation props in cell range
    cr_props.setPropertyValue("Validation", vprops)