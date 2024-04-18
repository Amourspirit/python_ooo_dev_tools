from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.calc.modify.page.header import Header, CalcStylePageKind
from ooodev.format.calc.modify.page.header.area import Img, PresetImageKind
from ooodev.format import Styler
from ooodev.utils.color import StandardColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc


def test_calc(loader) -> None:
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

        header_style = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=10.0,
            spacing=3.0,
            margin_left=1.5,
            margin_right=2.0,
            style_name=CalcStylePageKind.DEFAULT,
        )
        style = Img.from_preset(
            preset=PresetImageKind.BRICK_WALL,
            style_name=header_style.prop_style_name,
            style_family=header_style.prop_style_family_name,
        )
        Styler.apply(doc, header_style, style)
        props = style.get_style_props(doc)

        graphic = props.getPropertyValue(style.prop_inner._props.back_graphic)
        assert graphic is not None

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
