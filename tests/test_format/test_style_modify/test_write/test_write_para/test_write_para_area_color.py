from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.para.area import Color
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

        style = Color(color=StandardColor.BLUE_LIGHT3)
        style.apply(doc)
        props = style.get_style_props(doc)
        assert props.getPropertyValue("FillColor") == StandardColor.BLUE_LIGHT3

        f_style = Color.from_style(doc)
        assert f_style.prop_inner.prop_color == StandardColor.BLUE_LIGHT3
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
