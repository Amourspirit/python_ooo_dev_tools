from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, Any, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.direct.para.drop_cap import DropCaps, StyleCharKind
from ooodev.format.direct.structs.drop_cap_struct import DropCapStruct
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

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        p_len = len(para_text)

        dc = DropCaps(count=1)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        # dc.dispatch_reset()

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        assert pp.DropCapCharStyleName == ""
        assert pp.DropCapWholeWord == False
        inner_dc = cast(DropCapStruct, dc._get_style("drop_cap")[0])
        assert inner_dc == pp.DropCapFormat
        cursor.gotoEnd(False)

        dc = DropCaps(count=1, style=StyleCharKind.DROP_CAPS)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        # dc.dispatch_reset()

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.DropCapCharStyleName == StyleCharKind.DROP_CAPS.value
        assert pp.DropCapWholeWord == False
        inner_dc = cast(DropCapStruct, dc._get_style("drop_cap")[0])
        assert inner_dc == pp.DropCapFormat
        cursor.gotoEnd(False)

        dc = DropCaps(count=5, lines=5, style=StyleCharKind.DROP_CAPS)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        # dc.dispatch_reset()

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.DropCapCharStyleName == StyleCharKind.DROP_CAPS.value
        assert pp.DropCapWholeWord == False
        inner_dc = cast(DropCapStruct, dc._get_style("drop_cap")[0])
        assert inner_dc == pp.DropCapFormat
        cursor.gotoEnd(False)

        dc = DropCaps(count=3, whole_word=True)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        # dc.dispatch_reset()

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.DropCapCharStyleName == ""
        assert pp.DropCapWholeWord == True
        inner_dc = cast(DropCapStruct, dc._get_style("drop_cap")[0])
        assert inner_dc == pp.DropCapFormat
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
