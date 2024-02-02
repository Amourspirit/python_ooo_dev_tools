from __future__ import annotations
import pytest

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.modify.frame.options import StyleFrameKind, Properties, TextDirectionKind
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.exceptions import ex as mEx


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

        style = Properties(
            editable=True, printable=True, txt_direction=TextDirectionKind.LR_TB, style_name=StyleFrameKind.FRAME
        )

        style.apply(doc)
        # props = style.get_style_props(doc)

        f_style = Properties.from_style(doc, style.prop_style_name)
        assert f_style.prop_inner.prop_editable == style.prop_inner.prop_editable
        assert f_style.prop_inner.prop_printable == style.prop_inner.prop_printable
        assert f_style.prop_inner.prop_txt_direction == style.prop_inner.prop_txt_direction

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
