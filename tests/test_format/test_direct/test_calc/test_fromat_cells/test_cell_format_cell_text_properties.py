from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno

from ooodev.format.calc.direct.format_cells.alignment import Properties, TextDirectionKind
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
        style = Properties(wrap_auto=True, hyphen_active=True, direction=TextDirectionKind.PAGE)

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        assert cp.IsTextWrapped
        assert cp.ParaIsHyphenation
        assert cp.WritingMode == int(TextDirectionKind.PAGE)

        f_style = Properties.from_obj(cell)
        assert f_style.prop_wrap_auto == style.prop_wrap_auto
        assert f_style.prop_hyphen_active == style.prop_hyphen_active
        assert f_style.prop_direction == style.prop_direction
        # ====================================================

        cell_obj = Calc.get_cell_obj("B1")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        style = Properties(wrap_auto=False, shrink_to_fit=True, direction=TextDirectionKind.LR_TB)

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        assert cp.IsTextWrapped == False
        assert cp.ShrinkToFit
        assert cp.WritingMode == int(TextDirectionKind.LR_TB)

        f_style = Properties.from_obj(cell)
        assert f_style.prop_wrap_auto == style.prop_wrap_auto
        assert f_style.prop_shirnk_to_fit == style.prop_shirnk_to_fit
        assert f_style.prop_direction == style.prop_direction
        # ====================================================

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
