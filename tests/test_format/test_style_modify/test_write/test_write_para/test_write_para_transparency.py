from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.para.area import Color
from ooodev.format.writer.modify.para.transparency import Transparency, Intensity
from ooodev.format.styler import Styler
from ooodev.format import StandardColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=para_text)

        style_color = Color(color=StandardColor.BLUE_LIGHT3)
        style = Transparency(value=Intensity(80))
        Styler.apply(doc, style_color, style)
        props = style_color.get_style_props(doc)
        fp = cast("FillProperties", props)
        assert fp.FillColor == StandardColor.BLUE_LIGHT3
        assert fp.FillTransparence == 80

        f_style = Transparency.from_style(doc)
        assert f_style.prop_inner.prop_value == style.prop_inner.prop_value
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
