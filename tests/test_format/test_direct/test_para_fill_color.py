from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.para.area import Color
from ooodev.format import CommonColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write

from ooo.dyn.drawing.fill_style import FillStyle

if TYPE_CHECKING:
    from com.sun.star.drawing import FillProperties  # service


def test_write(loader, para_text) -> None:
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    # LibreOffice seems to have an unresolved bug with Background color.
    # https://bugs.documentfoundation.org/show_bug.cgi?id=99125
    # see Also: https://forum.openoffice.org/en/forum/viewtopic.php?p=417389&sid=17b21c173e4a420b667b45a2949b9cc5#p417389
    # The solution to these issues is to apply FillColor to Paragraph cursors TextParagraph.

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        p_len = len(para_text)

        dc = Color(CommonColor.LIME_GREEN)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))

        cursor_p = Write.get_paragraph_cursor(cursor)
        cursor_p.gotoPreviousParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        # note: it is necessary to reast fp each time cursor_p is moved
        for attr in dc.get_attrs():
            assert getattr(fp, attr) == dc._get(attr)
        cursor_p.gotoEnd(False)

        dc = Color(CommonColor.LIGHT_BLUE)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        # dc.dispatch_reset()

        cursor_p.gotoEnd(False)
        cursor_p.gotoPreviousParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        for attr in dc.get_attrs():
            assert getattr(fp, attr) == dc._get(attr)
        cursor_p.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text)
        # dc.dispatch_reset()

        cursor_p.gotoEnd(False)
        cursor_p.gotoPreviousParagraph(True)
        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillStyle == FillStyle.NONE
        cursor_p.gotoEnd(False)

        # test applying to cursor
        dc = Color(CommonColor.AQUAMARINE)
        dc.apply(cursor_p.TextParagraph)

        for _ in range(3):
            Write.append_para(cursor=cursor, text=para_text)
            cursor_p.gotoEnd(False)
            cursor_p.gotoPreviousParagraph(True)
            fp = cast("FillProperties", cursor_p.TextParagraph)
            for attr in dc.get_attrs():
                assert getattr(fp, attr) == dc._get(attr)
            cursor_p.gotoEnd(False)

        dc.default.apply(cursor_p.TextParagraph)

        fp = cast("FillProperties", cursor_p.TextParagraph)
        assert fp.FillStyle == FillStyle.NONE

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
