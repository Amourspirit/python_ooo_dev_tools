"""Test for ooodev.format.draw.direct.area.Img"""
from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.draw import Draw, DrawDoc

from ooodev.format.draw.modify.area import Img as FillImg
from ooodev.format.draw.modify.area import PresetImageKind


def test_draw(loader) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = DrawDoc(Draw.create_draw_doc())
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_75_PERCENT)
    try:
        slide = doc.get_slide(idx=0)

        width = 36
        height = 36
        x = int(width / 2)
        y = int(height / 2)

        rect = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        rect.set_string("Hello World!")
        style_modify = FillImg.from_preset(preset=PresetImageKind.POOL)
        doc.apply_styles(style_modify)

        f_style = FillImg.from_style(
            doc=doc.component,
            style_name=style_modify.prop_style_name,
            style_family=style_modify.prop_style_family_name,
        )
        assert f_style.prop_style_name == style_modify.prop_style_name

        Lo.delay(delay)
    finally:
        doc.close_doc()
