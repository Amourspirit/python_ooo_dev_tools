from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.para.borders import (
    Borders,
    InnerPadding,
    InnerShadow,
    ShadowLocation,
    Side,
    BorderLineKind,
    LineSize,
)
from ooodev.format import StandardColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
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
            border_side=Side(line=BorderLineKind.DOUBLE, color=StandardColor.DEFAULT_BLUE, width=LineSize.MEDIUM),
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

        assert f_style.prop_inner.prop_inner_padding.prop_left.value == pytest.approx(3.0, rel=1e-2)
        assert f_style.prop_inner.prop_inner_padding.prop_right.value == pytest.approx(3.0, rel=1e-2)
        assert f_style.prop_inner.prop_inner_padding.prop_top.value == pytest.approx(3.0, rel=1e-2)
        assert f_style.prop_inner.prop_inner_padding.prop_bottom.value == pytest.approx(3.0, rel=1e-2)

        assert f_style.prop_inner.prop_inner_shadow.prop_width.value == pytest.approx(2.0, rel=1.0e-2)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
