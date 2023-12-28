from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.lo import Lo
from ooodev.draw import Draw, DrawDoc, ZoomKind
from ooodev.format.draw.modify.font import FontOnly
from ooodev.format.draw.modify.font import FamilyGraphics, DrawStyleFamilyKind
from ooodev.format.inner.direct.write.char.font.font_only import FontOnly as InnerFontOnly


def test_draw(loader) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = DrawDoc(Draw.create_draw_doc())
    if not Lo.bridge_connector.headless:
        doc.set_visible()
        Lo.delay(500)
        doc.zoom(ZoomKind.ZOOM_75_PERCENT)
    try:
        slide = doc.get_slide(idx=0)

        width = 100
        height = 100
        x = int(width / 2)
        y = int(height / 2)

        style = FontOnly(
            name=InnerFontOnly.default.prop_name,
            size=18.0,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        doc.apply_styles(style)

        _ = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        # props = style.get_style_props(doc.component)

        f_style = FontOnly.from_style(
            doc=doc.component, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_name == InnerFontOnly.default.prop_name
        assert f_style.prop_inner.prop_size is not None
        assert f_style.prop_inner.prop_size.value == 18.0

        Lo.delay(delay)
    finally:
        doc.close_doc()
