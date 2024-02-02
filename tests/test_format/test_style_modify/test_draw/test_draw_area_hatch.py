from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.draw import Draw, DrawDoc
from ooodev.format.draw.modify.area import Hatch, PresetHatchKind
from ooodev.format.draw.modify.area import FamilyGraphics, DrawStyleFamilyKind
from ooodev.format import StandardColor


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

        style = Hatch.from_preset(
            preset=PresetHatchKind.GREEN_30_DEGREES,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        doc.apply_styles(style)

        _ = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        # props = style.get_style_props(doc.component)

        f_style = Hatch.from_style(
            doc=doc.component,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        assert f_style.prop_style_name == str(FamilyGraphics.DEFAULT_DRAWING_STYLE)
        assert f_style.prop_inner.prop_inner_hatch.prop_color == StandardColor.GREEN

        Lo.delay(delay)
    finally:
        doc.close_doc()
