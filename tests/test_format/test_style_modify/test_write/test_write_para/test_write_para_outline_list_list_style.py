from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.para.outline_list import ListStyle, StyleListKind, StyleParaKind
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
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

        style = ListStyle(list_style=StyleListKind.NUM_ABC, style_name=StyleParaKind.CAPTION)
        style.apply(doc)
        props = style.get_style_props(doc)
        pp = cast("ParagraphProperties", props)
        assert pp.NumberingStyleName == str(StyleListKind.NUM_ABC)

        f_style = ListStyle.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_list_style == str(StyleListKind.NUM_ABC)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
