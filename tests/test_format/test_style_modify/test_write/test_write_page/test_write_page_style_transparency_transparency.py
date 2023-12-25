from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.area import Color
from ooodev.format.writer.modify.page.transparency import Transparency, Intensity
from ooodev.format import Styler
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.color import StandardColor
from ooodev.office.write import Write


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)
        if not Lo.bridge_connector.headless:
            Write.append_para(cursor=cursor, text=para_text)

        style_color = Color(color=StandardColor.BLUE_LIGHT2)
        style = Transparency(value=Intensity(88))
        Styler.apply(doc, style_color, style)
        props = style_color.get_style_props(doc)

        assert props.getPropertyValue("FillColor") == StandardColor.BLUE_LIGHT2

        f_style_color = Color.from_style(doc=doc, style_name=style_color.prop_style_name)
        assert f_style_color.prop_inner.prop_color == style_color.prop_inner.prop_color

        f_style = Transparency.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_value == Intensity(88)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
