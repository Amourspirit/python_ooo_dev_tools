from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.para.transparency import Transparency

# from ooodev.format.inner.direct.write.para.area.color import Color
from ooodev.format.writer.direct.para.area import Color
from ooodev.format.writer.style.para import Para
from ooodev.format import StandardColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


def test_write(loader, para_text) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        cursor_p = Write.get_paragraph_cursor(cursor)

        Write.append_para(cursor=cursor, text=para_text)
        cursor_p.gotoEnd(False)

        dc = Color(StandardColor.LIME)
        tp = Transparency(52)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc, tp))
        Para.default.apply(cursor)
        cursor_p.gotoEnd(False)
        cursor_p.gotoPreviousParagraph(False)
        cursor_p.gotoStartOfParagraph(False)
        cursor_p.gotoEndOfParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillTransparence == tp.prop_value.value

        cursor_p.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
