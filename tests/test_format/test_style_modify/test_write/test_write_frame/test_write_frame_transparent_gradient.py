from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.frame.transparency import (
    Gradient,
    GradientStyle,
    Angle,
    Intensity,
    IntensityRange,
    Offset,
    StyleFrameKind,
)
from ooodev.format.writer.modify.frame.area import Color
from ooodev.format import Styler
from ooodev.utils.color import StandardColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
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

        style = Gradient(
            style=GradientStyle.LINEAR,
            angle=Angle(45),
            grad_intensity=IntensityRange(0, 100),
            style_name=StyleFrameKind.FRAME,
        )
        color_style = Color(
            color=StandardColor.GREEN_DARK3,
            style_name=style.prop_style_name,
            style_family=style.prop_style_family_name,
        )

        Styler.apply(doc, color_style, style)
        # props = style.get_style_props(doc)

        f_style = Gradient.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_inner == style.prop_inner.prop_inner

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
