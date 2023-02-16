from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.tabs import StyleParaKind, Tabs, TabAlign, FillCharKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write

if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


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

        pos = 14.5
        style = Tabs(position=pos, align=TabAlign.DECIMAL, decimal_char=",", style_name=StyleParaKind.CAPTION)
        style.apply(doc)
        props = style.get_style_props(doc)
        # pp = cast("ParagraphProperties", props)
        tb = Tabs.find(doc=doc, position=pos, style_name=style.prop_style_name)
        assert tb is not None
        assert tb.prop_inner.prop_decimal_char == ","
        assert tb.prop_inner.prop_align == TabAlign.DECIMAL

        f_style = Tabs.from_style(
            doc=doc, index=1, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_decimal_char == ","
        assert f_style.prop_inner.prop_align == TabAlign.DECIMAL

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
