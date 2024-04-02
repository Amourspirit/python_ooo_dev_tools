from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.draw import Draw, DrawDoc
from ooodev.format.draw.modify.area import Color as FillColor
from ooodev.format.draw.modify.area import FamilyGraphics, DrawStyleFamilyKind
from ooodev.utils.color import StandardColor


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

        width = 100
        height = 100
        x = int(width / 2)
        y = int(height / 2)

        style = FillColor(
            color=StandardColor.LIME_LIGHT2,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        doc.apply_styles(style)

        _ = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        # props = style.get_style_props(doc.component)

        f_style = FillColor.from_style(
            doc=doc.component,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        assert f_style.prop_inner.prop_color == StandardColor.LIME_LIGHT2

        Lo.delay(delay)
    finally:
        doc.close_doc()
