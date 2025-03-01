from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.page.borders import (
    Sides,
    Side,
    LineSize,
    BorderLineKind,
)
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
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

        side = Side(line=BorderLineKind.DOUBLE, color=StandardColor.RED_DARK3, width=LineSize.MEDIUM)

        style = Sides(all=side)
        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = Sides.from_style(doc, style.prop_style_name)
        f_side = f_style.prop_inner.prop_left
        assert f_side.prop_color == side.prop_color
        assert f_side.prop_width.value == pytest.approx(side.prop_width.value, rel=1e-2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
