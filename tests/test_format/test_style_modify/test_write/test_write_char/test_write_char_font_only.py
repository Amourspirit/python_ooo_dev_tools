from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.char.font import FontOnly, FontLang
from ooodev.format.writer.direct.char.font import InnerFontOnly as DirectFontOnly
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

        style = FontOnly(name=DirectFontOnly.default.prop_name, size=14.0)
        style.apply(doc)
        props = style.get_style_props(doc)
        assert props.getPropertyValue("CharFontName") == DirectFontOnly.default.prop_name

        f_style = FontOnly.from_style(doc)
        assert f_style.prop_inner.prop_name == DirectFontOnly.default.prop_name
        assert f_style.prop_inner.prop_size == 14.0
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
