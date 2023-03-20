from __future__ import annotations
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format import Styler
from ooodev.format.calc.modify.cell.background import Color, StyleCellKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.calc import Calc
from ooodev.utils.color import StandardColor


def test_write(loader) -> None:
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

        style = Color(color=StandardColor.GOLD, style_name=StyleCellKind.DEFAULT)
        Styler.apply(doc, style)

        f_style = Color.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_color == StandardColor.GOLD
        assert f_style.prop_inner.prop_color == style.prop_inner.prop_color

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
