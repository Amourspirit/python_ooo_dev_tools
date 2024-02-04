from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.font import FontEffects, FontLine, FontUnderlineEnum, StyleParaKind
from ooodev.utils.color import StandardColor, Color
from ooodev.loader.lo import Lo
from ooodev.write import Write, WriteDoc, ZoomKind


def test_write(loader, para_text) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = WriteDoc(Write.create_doc())
    if not Lo.bridge_connector.headless:
        doc.set_visible()
        Lo.delay(500)
        doc.zoom(ZoomKind.ZOOM_75_PERCENT)
    try:
        cursor = doc.get_cursor()
        cursor.append_para(text=para_text)

        style = FontEffects(
            color=StandardColor.BLUE_LIGHT1,
            underline=FontLine(line=FontUnderlineEnum.DOUBLE),
            style_name=StyleParaKind.CAPTION,
        )
        doc.apply_styles(style)
        props = style.get_style_props(doc.component)
        assert props.getPropertyValue("CharUnderline") == FontUnderlineEnum.DOUBLE

        f_style = FontEffects.from_style(doc=doc.component, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_color == StandardColor.BLUE_LIGHT1
        assert f_style.prop_inner.prop_underline == FontLine(line=FontUnderlineEnum.DOUBLE, color=Color(-1))
        Lo.delay(delay)
    finally:
        doc.close_doc()
