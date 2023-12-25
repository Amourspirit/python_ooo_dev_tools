from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.font import FontOnly, InnerFontOnly, FontLang, StyleParaKind

# from ooodev.format.writer.direct.char.font import FontOnly as DirectFontOnly
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

        style = FontOnly(name=InnerFontOnly.default.prop_name, size=18.0, style_name=StyleParaKind.CAPTION)
        style.apply(doc)
        props = style.get_style_props(doc)
        assert props.getPropertyValue("CharFontName") == InnerFontOnly.default.prop_name

        f_style = FontOnly.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_name == InnerFontOnly.default.prop_name
        assert f_style.prop_inner.prop_size is not None
        assert f_style.prop_inner.prop_size.value == 18.0
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
