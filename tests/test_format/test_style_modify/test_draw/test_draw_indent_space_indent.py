from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.utils.lo import Lo
from ooodev.draw import Draw, DrawDoc, ZoomKind
from ooodev.format.draw.modify.indent_space import Indent
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

        style = Indent(
            before=10.5,
            after=5.5,
            first=10.2,
            # auto=True,
            style_name=FamilyGraphics.DEFAULT_DRAWING_STYLE,
            style_family=DrawStyleFamilyKind.GRAPHICS,
        )
        doc.apply_styles(style)

        _ = slide.draw_rectangle(x=x, y=y, width=width, height=height)
        # props = style.get_style_props(doc.component)

        f_style = Indent.from_style(
            doc=doc.component, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_after is not None
        assert f_style.prop_inner.prop_after.get_value_mm100() in [550 - 2 + i for i in range(5)]
        assert f_style.prop_inner.prop_before is not None
        assert f_style.prop_inner.prop_before.get_value_mm100() in [1050 - 2 + i for i in range(5)]
        assert f_style.prop_inner.prop_first is not None
        assert f_style.prop_inner.prop_first.get_value_mm100() in [1020 - 2 + i for i in range(5)]

        Lo.delay(delay)
    finally:
        doc.close_doc()
