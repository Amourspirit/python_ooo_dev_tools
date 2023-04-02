from __future__ import annotations
from typing import cast
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.style import Page, WriterStylePageKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.color import CommonColor


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

        style = Page(name=WriterStylePageKind.FIRST_PAGE)
        # style.apply(cursor)
        Write.append_para(cursor=cursor, text=para_text, styles=(style,))

        f_style = Page.from_obj(cursor)
        assert f_style.prop_name == style.prop_name
        assert f_style.prop_name == str(WriterStylePageKind.FIRST_PAGE)

        xprops = style.get_style_props()
        assert xprops is not None
        xprops.setPropertyValue("BackColor", CommonColor.CORAL)
        val = cast(int, xprops.getPropertyValue("BackColor"))
        assert val == CommonColor.CORAL

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
