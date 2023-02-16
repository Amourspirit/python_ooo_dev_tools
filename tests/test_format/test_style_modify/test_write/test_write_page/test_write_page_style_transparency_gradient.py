from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.area import Color
from ooodev.format.writer.modify.page.transparency import (
    Gradient,
    StylePageKind,
    Intensity,
    IntensityRange,
    Angle,
    Offset,
    GradientStyle,
)
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

        style_color = Color(color=StandardColor.YELLOW_DARK2)
        style = Gradient(
            style=GradientStyle.LINEAR,
            angle=Angle(30),
            grad_intensity=IntensityRange(0, 100),
        )
        Styler.apply(doc, style_color, style)
        props = style_color.get_style_props(doc)

        assert props.getPropertyValue("FillColor") == StandardColor.YELLOW_DARK2

        f_style_color = Color.from_style(doc=doc, style_name=style_color.prop_style_name)
        assert f_style_color.prop_inner.prop_color == style_color.prop_inner.prop_color

        f_style = Gradient.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_inner == style.prop_inner.prop_inner

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
