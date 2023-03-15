from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.calc.direct.background import Color
from ooodev.format.calc.direct.borders import (
    Borders,
    Shadow,
    Side,
    BorderLineKind,
    ShadowLocation,
    Padding,
)
from ooodev.format import CommonColor, Styler
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

if TYPE_CHECKING:
    from com.sun.star.table import CellProperties  # service
    from com.sun.star.table import CellRange  # service


def test_calc_background(loader) -> None:
    delay = 0  # 0 if Lo.bridge_connector.headless else 5_000
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
        color = Color(CommonColor.LIGHT_BLUE)
        cb = Borders(border_side=Side())
        # color.apply(cell)
        Styler.apply(cell, color, cb)
        cp = cast("CellProperties", cell)
        assert cp.CellBackColor == CommonColor.LIGHT_BLUE

        color = color.empty
        clr = color.empty
        clr.apply(cell)
        cp = cast("CellProperties", cell)
        assert cp.CellBackColor == -1

        rng_obj = Calc.get_range_obj("A3:F8")
        cr = Calc.get_cell_range(sheet, rng_obj)
        color = Color(CommonColor.LIGHT_GREEN)
        color.apply(cr)
        cr = cast("CellRange", cr)
        assert cr.CellBackColor == CommonColor.LIGHT_GREEN

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
