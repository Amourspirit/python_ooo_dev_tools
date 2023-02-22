from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.header import Header, StylePageKind
from ooodev.format.writer.modify.page.header.borders import (
    Sides,
    Side,
    SideFlags,
    LineSize,
    StylePageKind,
    BorderLineStyleEnum,
)
from ooodev.format import Styler
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.color import StandardColor
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

        header_style = Header(
            on=True,
            shared_first=True,
            shared=True,
            height=10.0,
            spacing=3.0,
            spacing_dyn=True,
            margin_left=1.5,
            margin_right=2.0,
        )
        side = Side(line=BorderLineStyleEnum.DOUBLE, color=StandardColor.RED_DARK3, width=LineSize.MEDIUM)

        style = Sides(all=side)
        Styler.apply(doc, header_style, style)
        # props = style.get_style_props(doc)

        f_style = Sides.from_style(doc, style.prop_style_name)
        f_side = f_style.prop_inner.prop_left
        assert f_side.prop_color == side.prop_color
        assert f_side.prop_width == pytest.approx(side.prop_width, rel=1e2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
