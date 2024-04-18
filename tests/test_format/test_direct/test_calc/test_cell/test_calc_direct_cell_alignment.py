from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooo.dyn.table.cell_vert_justify2 import CellVertJustify2
from ooo.dyn.table.cell_hori_justify import CellHoriJustify

from ooodev.format.calc.direct.cell.alignment import TextAlign, HoriAlignKind, VertAlignKind
from ooodev.format import Styler
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.units.unit_mm100 import UnitMM100

if TYPE_CHECKING:
    from com.sun.star.table import CellProperties  # service
    from com.sun.star.table import CellRange  # service


def test_calc(loader) -> None:
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
        style = TextAlign(hori_align=HoriAlignKind.CENTER, vert_align=VertAlignKind.MIDDLE)

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        assert cp.HoriJustify == CellHoriJustify.CENTER
        assert cp.VertJustify == CellVertJustify2.CENTER
        assert cp.HoriJustifyMethod == 0
        assert cp.VertJustifyMethod == 0

        f_style = TextAlign.from_obj(cell)
        assert f_style.prop_hori_align == HoriAlignKind.CENTER
        assert f_style.prop_vert_align == VertAlignKind.MIDDLE
        assert f_style.prop_indent.get_value_mm100() == 0

        # ====================================================

        cell_obj = Calc.get_cell_obj("B1")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        indent = UnitMM100.from_pt(10)
        style = TextAlign(hori_align=HoriAlignKind.LEFT, indent=indent, vert_align=VertAlignKind.TOP)

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        assert cp.HoriJustify == CellHoriJustify.LEFT
        assert cp.VertJustify == CellVertJustify2.TOP
        assert cp.HoriJustifyMethod == 0
        assert cp.VertJustifyMethod == 0

        f_style = TextAlign.from_obj(cell)
        assert f_style.prop_hori_align == HoriAlignKind.LEFT
        assert f_style.prop_vert_align == VertAlignKind.TOP
        assert f_style.prop_indent.get_value_mm100() in range(indent.value - 2, indent.value + 3)  # +- 2

        # ===================Distributed==========================

        cell_obj = Calc.get_cell_obj("A2")
        Calc.set_val(value="distributed text", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        style = TextAlign(hori_align=HoriAlignKind.DISTRIBUTED, vert_align=VertAlignKind.DISTRIBUTED)

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        assert cp.HoriJustify == CellHoriJustify.BLOCK
        assert cp.VertJustify == CellVertJustify2.BLOCK
        assert cp.HoriJustifyMethod == 1
        assert cp.VertJustifyMethod == 1

        f_style = TextAlign.from_obj(cell)
        assert f_style.prop_hori_align == HoriAlignKind.DISTRIBUTED
        assert f_style.prop_vert_align == VertAlignKind.DISTRIBUTED

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
