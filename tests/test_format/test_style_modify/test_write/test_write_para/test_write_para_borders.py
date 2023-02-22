from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.borders import (
    Borders,
    InnerPadding,
    InnerShadow,
    ShadowLocation,
    ShadowFormat,
    Side,
    SideFlags,
    Sides,
    BorderLineStyleEnum,
    LineSize,
)
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

        style = Borders(
            border_side=Side(line=BorderLineStyleEnum.DOUBLE, color=StandardColor.DEFAULT_BLUE, width=LineSize.MEDIUM),
            padding=InnerPadding(all=3.0),
            shadow=InnerShadow(location=ShadowLocation.BOTTOM_RIGHT, width=2.0),
            merge=False,
        )
        style.apply(doc)
        props = style.get_style_props(doc=doc)
        assert props.getPropertyValue("ParaIsConnectBorder") == False

        f_style = Borders.from_style(doc)
        assert f_style.prop_inner.prop_inner_sides.prop_left.prop_color == StandardColor.DEFAULT_BLUE
        assert f_style.prop_inner.prop_inner_sides.prop_right.prop_color == StandardColor.DEFAULT_BLUE
        assert f_style.prop_inner.prop_inner_sides.prop_top.prop_color == StandardColor.DEFAULT_BLUE
        assert f_style.prop_inner.prop_inner_sides.prop_bottom.prop_color == StandardColor.DEFAULT_BLUE

        assert f_style.prop_inner.prop_inner_padding.prop_left == pytest.approx(3.0, rel=1e2)
        assert f_style.prop_inner.prop_inner_padding.prop_right == pytest.approx(3.0, rel=1e2)
        assert f_style.prop_inner.prop_inner_padding.prop_top == pytest.approx(3.0, rel=1e2)
        assert f_style.prop_inner.prop_inner_padding.prop_bottom == pytest.approx(3.0, rel=1e2)

        assert f_style.prop_inner.prop_inner_shadow.prop_width == pytest.approx(2.0, rel=1.0e2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
