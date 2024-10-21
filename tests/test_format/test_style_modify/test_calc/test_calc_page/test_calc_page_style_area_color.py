from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.calc.modify.page.area import Color
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.color import StandardColor
from ooodev.office.calc import Calc


def test_calc(loader) -> None:
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

        style = Color(color=StandardColor.BLUE_LIGHT2)
        style.apply(doc)
        props = style.get_style_props(doc)

        assert props.getPropertyValue("BackColor") == StandardColor.BLUE_LIGHT2

        f_style = Color.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_color == style.prop_inner.prop_color

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
