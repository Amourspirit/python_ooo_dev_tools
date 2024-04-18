from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.area.color import Color, StylePageKind
from ooodev.format import CommonColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


def _test_write(loader, para_text) -> None:
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

        style = Color(CommonColor.LIGHT_GREEN)
        style.apply(doc)

        cobj = Color.from_style(doc, style.prop_style_name)
        assert cobj.prop_color == style.prop_color

        style = Color(CommonColor.DIM_GRAY, style_name=StylePageKind.FIRST_PAGE)
        style.apply(doc)

        cobj = Color.from_style(doc, style.prop_style_name)
        assert cobj.prop_color == style.prop_color

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
