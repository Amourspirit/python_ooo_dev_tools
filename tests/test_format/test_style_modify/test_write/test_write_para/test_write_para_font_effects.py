from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.font import (
    FontEffects,
    FontStrikeoutEnum,
    FontReliefEnum,
    FontUnderlineEnum,
    CaseMapEnum,
    StyleParaKind,
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

        style = FontEffects(
            color=StandardColor.BLUE_LIGHT1, underine=FontUnderlineEnum.DOUBLE, style_name=StyleParaKind.CAPTION
        )
        style.apply(doc)
        props = style.get_style_props(doc)
        assert props.getPropertyValue("CharUnderline") == FontUnderlineEnum.DOUBLE

        f_style = FontEffects.from_style(doc=doc, style_name=style.prop_style_name)
        assert f_style.prop_inner.prop_color == StandardColor.BLUE_LIGHT1
        assert f_style.prop_inner.prop_underline == FontUnderlineEnum.DOUBLE
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
