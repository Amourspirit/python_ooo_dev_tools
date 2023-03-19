from __future__ import annotations
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format import Styler
from ooodev.format.calc.modify.cell.alignment import TextOrientation, EdgeKind, StyleCellKind, Angle
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Calc.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        sheet = Calc.get_active_sheet()

        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)

        style = TextOrientation(
            vert_stack=False, rotation=Angle(10), edge=EdgeKind.INSIDE, style_name=StyleCellKind.DEFAULT
        )
        Styler.apply(doc, style)

        f_style = TextOrientation.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_edge == style.prop_inner.prop_edge
        assert f_style.prop_inner.prop_rotation == style.prop_inner.prop_rotation
        assert f_style.prop_inner.prop_vert_stacked == style.prop_inner.prop_vert_stacked

        # ====================================================================

        cell_obj = Calc.get_cell_obj("B1")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)

        style = TextOrientation(
            vert_stack=False, rotation=Angle(-10), edge=EdgeKind.INSIDE, style_name=StyleCellKind.DEFAULT
        )
        Styler.apply(doc, style)

        f_style = TextOrientation.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_edge == style.prop_inner.prop_edge
        assert f_style.prop_inner.prop_rotation == style.prop_inner.prop_rotation
        assert f_style.prop_inner.prop_vert_stacked == style.prop_inner.prop_vert_stacked

        # ====================================================================

        cell_obj = Calc.get_cell_obj("D3")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)

        style = TextOrientation(vert_stack=False, rotation=20, edge=EdgeKind.LOWER, style_name=StyleCellKind.DEFAULT)
        Styler.apply(doc, style)

        f_style = TextOrientation.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_edge == style.prop_inner.prop_edge
        assert f_style.prop_inner.prop_rotation == style.prop_inner.prop_rotation
        assert f_style.prop_inner.prop_vert_stacked == style.prop_inner.prop_vert_stacked

        # ====================================================================

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
