from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno

from ooodev.format.calc.direct.cell.cell_protection import CellProtection
from ooodev.format import Styler
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

if TYPE_CHECKING:
    from com.sun.star.table import CellProperties  # service


def test_calc(loader) -> None:
    delay = 0
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    try:
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_200_PERCENT)

        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        style = CellProtection(hide_all=False, hide_formula=False, protected=True, hide_print=True)

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        struct = cp.CellProtection
        assert struct.IsLocked == True
        assert struct.IsFormulaHidden == False
        assert struct.IsHidden == False
        assert struct.IsPrintHidden == True

        f_style = CellProtection.from_obj(cell)
        assert f_style.prop_hide_all == style.prop_hide_all
        assert f_style.prop_hide_formula == style.prop_hide_formula
        assert f_style.prop_protected == style.prop_protected
        assert f_style.prop_hide_print == style.prop_hide_print
        # ====================================================

        cell_obj = Calc.get_cell_obj("B1")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        style = CellProtection().hide_formula.protected

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        struct = cp.CellProtection
        assert struct.IsLocked == True
        assert struct.IsFormulaHidden == True
        assert struct.IsHidden == False
        assert struct.IsPrintHidden == False

        f_style = CellProtection.from_obj(cell)
        assert f_style.prop_hide_all == style.prop_hide_all
        assert f_style.prop_hide_formula == style.prop_hide_formula
        assert f_style.prop_protected == style.prop_protected
        assert f_style.prop_hide_print == style.prop_hide_print
        # ====================================================

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
