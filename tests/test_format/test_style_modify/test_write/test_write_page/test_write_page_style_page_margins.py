from __future__ import annotations
from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.page.page import Margins
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

        style = Margins(left=10, right=10, top=18, bottom=18, gutter=8)
        style.apply(doc)
        props = style.get_style_props(doc)
        for attrib in style.prop_inner._props:
            val = cast(int, style.prop_inner._get(attrib))
            assert props.getPropertyValue(attrib) in range(val - 2, val + 3)

        f_style = Margins.from_style(doc, style.prop_style_name)

        for attrib in style.prop_inner._props:
            val = cast(int, style.prop_inner._get(attrib))
            assert f_style.prop_inner._get(attrib) in range(val - 2, val + 3)
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
