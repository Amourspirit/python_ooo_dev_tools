from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.frame.borders import (
    StyleFrameKind,
    Side,
    Sides,
    BorderLineStyleEnum,
    LineSize,
    InnerSides,
)
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
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

        side = Side(line=BorderLineStyleEnum.DOUBLE, color=StandardColor.RED_DARK3, width=LineSize.MEDIUM)

        style = Sides(all=side, style_name=StyleFrameKind.FRAME)
        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = Sides.from_style(doc=doc, style_name=style.prop_style_name)
        f_side = f_style.prop_inner.prop_left
        assert f_side.prop_color == side.prop_color
        assert f_side.prop_width == pytest.approx(side.prop_width, rel=1e2)

        side = Side(line=BorderLineStyleEnum.DOUBLE_THIN, color=StandardColor.BLUE_DARK1, width=LineSize.THIN)

        inner = InnerSides(all=side)
        # _ = style.prop_inner
        style.prop_inner = inner
        style.apply(doc)

        f_style = Sides.from_style(doc=doc, style_name=style.prop_style_name)
        f_side = f_style.prop_inner.prop_left
        assert f_side.prop_color == side.prop_color
        assert f_side.prop_width == pytest.approx(side.prop_width, rel=1e2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
