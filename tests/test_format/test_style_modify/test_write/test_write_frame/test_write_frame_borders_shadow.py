from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.frame.borders import StyleFrameKind, Shadow, ShadowLocation, InnerShadow
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
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

        style = Shadow(
            location=ShadowLocation.BOTTOM_RIGHT,
            color=StandardColor.GRAY_LIGHT1,
            width=2.4,
            style_name=StyleFrameKind.FRAME,
        )
        style.apply(doc)
        # props = style.get_style_props(doc)
        f_style = Shadow.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_width.value == pytest.approx(style.prop_inner.prop_width.value, rel=1e-2)
        assert f_style.prop_inner.prop_color == StandardColor.GRAY_LIGHT1
        assert f_style.prop_inner.prop_location == ShadowLocation.BOTTOM_RIGHT

        inner = InnerShadow(location=ShadowLocation.TOP_RIGHT, color=StandardColor.RED_LIGHT2, width=1.9)
        style.prop_inner = inner
        style.apply(doc)
        # props = style.get_style_props(doc)
        f_style = Shadow.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_width.value == pytest.approx(style.prop_inner.prop_width.value, rel=1e-2)
        assert f_style.prop_inner.prop_color == StandardColor.RED_LIGHT2
        assert f_style.prop_inner.prop_location == ShadowLocation.TOP_RIGHT

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
