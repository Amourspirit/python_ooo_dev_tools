from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.char.borders import Side, Sides, BorderLineKind, LineSize
from ooodev.format import StandardColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        Write.append_para(cursor=cursor, text=para_text)

        style = Sides(
            border_side=Side(line=BorderLineKind.DOUBLE, color=StandardColor.DEFAULT_BLUE, width=LineSize.MEDIUM)
        )
        style.apply(doc)
        props = style.get_style_props(doc)
        left = props.getPropertyValue("CharLeftBorder")
        assert style.prop_inner.prop_left == left

        f_style = Sides.from_style(doc)
        assert f_style.prop_inner.prop_left.prop_color == StandardColor.DEFAULT_BLUE
        assert f_style.prop_inner.prop_top.prop_color == StandardColor.DEFAULT_BLUE
        assert f_style.prop_inner.prop_right.prop_color == StandardColor.DEFAULT_BLUE
        assert f_style.prop_inner.prop_bottom.prop_color == StandardColor.DEFAULT_BLUE

        assert f_style.prop_inner.prop_left.prop_line == BorderLineKind.DOUBLE
        assert f_style.prop_inner.prop_top.prop_line == BorderLineKind.DOUBLE
        assert f_style.prop_inner.prop_right.prop_line == BorderLineKind.DOUBLE
        assert f_style.prop_inner.prop_bottom.prop_line == BorderLineKind.DOUBLE

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
