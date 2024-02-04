from __future__ import annotations
from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.page import LayoutSettings, PageStyleLayout, NumberingTypeEnum, StyleParaKind
from ooodev.utils.gui import GUI
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

        style = LayoutSettings(
            layout=PageStyleLayout.MIRRORED,
            numbers=NumberingTypeEnum.CHARS_UPPER_LETTER,
            ref_style=StyleParaKind.TEXT_BODY,
        )
        style.apply(doc)
        props = style.get_style_props(doc)

        f_style = LayoutSettings.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_layout == style.prop_inner.prop_layout
        assert f_style.prop_inner.prop_numbers == style.prop_inner.prop_numbers
        assert f_style.prop_inner.prop_ref_style == style.prop_inner.prop_ref_style

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
