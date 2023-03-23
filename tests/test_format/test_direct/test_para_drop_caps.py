from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.para.drop_caps import DropCaps, StyleCharKind
from ooodev.format.inner.direct.structs.drop_cap_struct import DropCapStruct
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


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

        dc = DropCaps(count=1)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        cursor_p = Write.get_paragraph_cursor(cursor)
        cursor_p.gotoEnd(False)
        cursor_p.gotoPreviousParagraph(True)
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        assert pp.DropCapCharStyleName == ""
        assert pp.DropCapWholeWord == False
        inner_dc = cast(DropCapStruct, dc._get_style_inst("drop_cap"))
        assert inner_dc == pp.DropCapFormat
        cursor_p.gotoEnd(False)

        dc = DropCaps(count=1, style=StyleCharKind.DROP_CAPS)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        cursor_p.gotoEnd(False)
        cursor_p.gotoPreviousParagraph(True)
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        assert pp.DropCapCharStyleName == StyleCharKind.DROP_CAPS.value
        assert pp.DropCapWholeWord == False
        inner_dc = cast(DropCapStruct, dc._get_style_inst("drop_cap"))
        assert inner_dc == pp.DropCapFormat
        cursor_p.gotoEnd(False)

        dc = DropCaps(count=5, lines=5, style=StyleCharKind.DROP_CAPS)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        cursor_p.gotoEnd(False)
        cursor_p.gotoPreviousParagraph(True)
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        assert pp.DropCapCharStyleName == StyleCharKind.DROP_CAPS.value
        assert pp.DropCapWholeWord == False
        inner_dc = cast(DropCapStruct, dc._get_style_inst("drop_cap"))
        assert inner_dc == pp.DropCapFormat
        cursor_p.gotoEnd(False)

        dc = DropCaps(count=3, whole_word=True)
        Write.append_para(cursor=cursor, text=para_text, styles=(dc,))
        cursor_p.gotoEnd(False)
        cursor_p.gotoPreviousParagraph(True)
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        assert pp.DropCapCharStyleName == ""
        assert pp.DropCapWholeWord == True
        inner_dc = cast(DropCapStruct, dc._get_style_inst("drop_cap"))
        assert inner_dc == pp.DropCapFormat
        cursor_p.gotoEnd(False)

        # set drop cap on cursor
        dc = DropCaps(count=3, whole_word=True, style=StyleCharKind.DROP_CAPS)
        dc.apply(cursor_p.TextParagraph)
        for _ in range(2):
            Write.append_para(cursor=cursor, text=para_text)
            cursor_p.gotoEnd(False)
            cursor_p.gotoPreviousParagraph(True)
            pp = cast("ParagraphProperties", cursor_p.TextParagraph)
            assert pp.DropCapCharStyleName == StyleCharKind.DROP_CAPS.value
            assert pp.DropCapWholeWord == True
            inner_dc = cast(DropCapStruct, dc._get_style_inst("drop_cap"))
            assert inner_dc == pp.DropCapFormat
            cursor_p.gotoEnd(False)

        dc.default.apply(cursor_p.TextParagraph)
        pp = cast("ParagraphProperties", cursor_p.TextParagraph)
        assert pp.DropCapCharStyleName == ""
        assert pp.DropCapWholeWord == False

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
