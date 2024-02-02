from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.loader.lo import Lo
from ooodev.draw import Draw, DrawDoc, ZoomKind
from ooodev.format.draw.modify.indent_space import LineSpacing, ModeKind
from ooodev.format.draw.modify import FamilyGraphics, DrawStyleFamilyKind


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

        style = LineSpacing(
            mode=ModeKind.PROPORTIONAL,
            value=87,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        doc.apply_styles(style)

        _ = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        # props = style.get_style_props(doc.component)

        f_style = LineSpacing.from_style(
            doc=doc.component, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_mode is not None
        assert f_style.prop_inner.prop_mode == ModeKind.PROPORTIONAL
        assert f_style.prop_inner.prop_value is not None
        assert f_style.prop_inner.prop_value == 87

        Lo.delay(delay)
    finally:
        doc.close_doc()
