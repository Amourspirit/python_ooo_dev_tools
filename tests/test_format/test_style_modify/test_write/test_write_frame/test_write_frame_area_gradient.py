from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.frame.area import (
    Gradient,
    InnerGradient,
    GradientStyle,
    PresetGradientKind,
    StyleFrameKind,
    Offset,
    ColorRange,
    IntensityRange,
)
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.color import StandardColor


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

        style = Gradient.from_preset(preset=PresetGradientKind.DEEP_OCEAN, style_name=StyleFrameKind.FRAME)

        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = Gradient.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_inner == style.prop_inner.prop_inner

        inner = InnerGradient.from_preset(preset=PresetGradientKind.SUNSHINE)
        style.prop_inner = inner

        style.apply(doc)
        f_style = Gradient.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_inner == style.prop_inner.prop_inner

        inner = InnerGradient(
            style=GradientStyle.LINEAR,
            step_count=0,
            offset=Offset(0, 0),
            angle=45,
            border=80,
            grad_color=ColorRange(StandardColor.RED_DARK2, StandardColor.BLUE_LIGHT3),
            grad_intensity=IntensityRange(100, 100),
        )
        style.prop_inner = inner
        style.apply(doc)
        f_style = Gradient.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_inner == style.prop_inner.prop_inner

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
