from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.modify.frame.area import StyleFrameKind, Pattern, PresetPatternKind
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
        GUI.zoom(GUI.ZoomEnum.ENTIRE_PAGE)
    try:
        cursor = Write.get_cursor(doc)
        if not Lo.bridge_connector.headless:
            Write.append_para(cursor=cursor, text=para_text)

        style = Pattern.from_preset(preset=PresetPatternKind.HORIZONTAL_BRICK, style_name=StyleFrameKind.FRAME)

        style.apply(doc)
        # props = style.get_style_props(doc)
        f_style = Pattern.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_tile == style.prop_inner.prop_tile
        assert f_style.prop_inner.prop_stretch == style.prop_inner.prop_stretch

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
