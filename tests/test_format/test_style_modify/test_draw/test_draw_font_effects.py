from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.lo import Lo
from ooodev.draw import Draw, DrawDoc, ZoomKind
from ooodev.format.draw.modify.font import FontEffects, FontLine, FontUnderlineEnum
from ooodev.format.draw.modify.area import FamilyGraphics, DrawStyleFamilyKind
from ooodev.utils.color import StandardColor, Color


def test_draw(loader) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = DrawDoc(Draw.create_draw_doc())
    if not Lo.bridge_connector.headless:
        doc.set_visible()
        Lo.delay(500)
        doc.zoom(ZoomKind.ZOOM_75_PERCENT)
    try:
        slide = slide = doc.get_slide(idx=0)

        width = 100
        height = 100
        x = int(width / 2)
        y = int(height / 2)

        style = FontEffects(
            color=StandardColor.BLUE_LIGHT1,
            underline=FontLine(line=FontUnderlineEnum.DOUBLE),
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        doc.apply_styles(style)

        _ = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        props = style.get_style_props(doc.component)
        assert props.getPropertyValue("CharUnderline") == FontUnderlineEnum.DOUBLE

        f_style = FontEffects.from_style(
            doc=doc.component, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_color is not None
        assert f_style.prop_inner.prop_color == StandardColor.BLUE_LIGHT1
        assert f_style.prop_inner.prop_underline == FontLine(line=FontUnderlineEnum.DOUBLE, color=Color(-1))

        Lo.delay(delay)
    finally:
        doc.close_doc()
