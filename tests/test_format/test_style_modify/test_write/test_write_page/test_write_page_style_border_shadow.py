from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.borders import Shadow, ShadowLocation, StylePageKind
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

        style = Shadow(location=ShadowLocation.BOTTOM_RIGHT, color=StandardColor.GRAY_LIGHT1, width=2.4)
        style.apply(doc)
        # props = style.get_style_props(doc)
        f_style = Shadow.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_width.value == pytest.approx(style.prop_inner.prop_width.value, rel=1e-2)
        assert f_style.prop_inner.prop_color == StandardColor.GRAY_LIGHT1

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
