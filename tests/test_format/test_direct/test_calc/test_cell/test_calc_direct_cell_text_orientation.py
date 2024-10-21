from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooo.dyn.table.cell_orientation import CellOrientation

# from com.sun.star.table import CellOrientation

from ooodev.format.calc.direct.cell.alignment import Angle, EdgeKind, TextOrientation
from ooodev.format import Styler
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo

if TYPE_CHECKING:
    from com.sun.star.table import CellProperties  # service


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
        style = TextOrientation(vert_stack=False, rotation=Angle(10), edge=EdgeKind.INSIDE)

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        assert cp.RotateAngle == Angle(10).get_angle100()
        assert cp.RotateReference == EdgeKind.INSIDE.value
        assert cp.Orientation == CellOrientation.STANDARD

        f_style = TextOrientation.from_obj(cell)
        assert f_style.prop_edge == style.prop_edge
        assert f_style.prop_rotation == style.prop_rotation
        assert f_style.prop_vert_stacked == style.prop_vert_stacked
        # ====================================================

        cell_obj = Calc.get_cell_obj("B1")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        style = TextOrientation(vert_stack=False, rotation=Angle(-10), edge=EdgeKind.INSIDE)

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        assert cp.RotateAngle == 35000
        assert cp.RotateAngle == Angle(-10).get_angle100()
        assert cp.RotateReference == EdgeKind.INSIDE.value
        assert cp.Orientation == CellOrientation.STANDARD

        f_style = TextOrientation.from_obj(cell)
        assert f_style.prop_edge == style.prop_edge
        assert f_style.prop_rotation == style.prop_rotation
        assert f_style.prop_vert_stacked == style.prop_vert_stacked
        # ====================================================

        cell_obj = Calc.get_cell_obj("D3")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        style = TextOrientation(vert_stack=False, rotation=20, edge=EdgeKind.LOWER)

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        assert cp.RotateAngle == Angle(20).get_angle100()
        assert cp.RotateReference == EdgeKind.LOWER.value
        assert cp.Orientation == CellOrientation.STANDARD

        f_style = TextOrientation.from_obj(cell)
        assert f_style.prop_edge == style.prop_edge
        assert f_style.prop_rotation == style.prop_rotation
        assert f_style.prop_vert_stacked == style.prop_vert_stacked
        # ====================================================

        cell_obj = Calc.get_cell_obj("C6")
        Calc.set_val(value="Stacked", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        style = TextOrientation(vert_stack=True)

        Styler.apply(cell, style)
        cp = cast("CellProperties", cell)
        assert cp.Orientation == CellOrientation.STACKED

        f_style = TextOrientation.from_obj(cell)
        assert f_style.prop_vert_stacked == style.prop_vert_stacked
        # ====================================================

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
