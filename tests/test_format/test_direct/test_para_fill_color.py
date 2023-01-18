from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.direct.para.fill_color import FillColor
from ooodev.format import CommonColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


@pytest.mark.skip_headless("Requires Dispatch")
def test_write(loader, para_text) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    delay = 0
    # delay = 0 if Lo.bridge_connector.headless else 3_000

    # Fillcolor is done via dispatch commands for Writer.
    # LibreOffice seems to have an unresolved bug with Background color.
    # https://bugs.documentfoundation.org/show_bug.cgi?id=99125
    # see Also: https://forum.openoffice.org/en/forum/viewtopic.php?p=417389&sid=17b21c173e4a420b667b45a2949b9cc5#p417389

    # Unable to access properties to test. This test exist as a visual test only.

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        p_len = len(para_text)

        dc = FillColor(CommonColor.LIME_GREEN)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        # assert pp.ParaBackColor == CommonColor.LIME_GREEN
        cursor.gotoEnd(False)

        dc = FillColor(CommonColor.LIGHT_BLUE)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        # dc.dispatch_reset()

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        # assert pp.DropCapWholeWord == False
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text)
        # dc.dispatch_reset()

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        # assert pp.DropCapWholeWord == False
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
