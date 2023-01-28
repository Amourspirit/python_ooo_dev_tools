from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.direct.para.text_flow import Breaks, BreakType
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_props() -> None:
    brk = Breaks(type=BreakType.PAGE_BEFORE)
    assert brk.prop_type == BreakType.PAGE_BEFORE
    assert brk._get("BreakType") == BreakType.PAGE_BEFORE

    brk = Breaks(type=BreakType.PAGE_BEFORE, style="HTML")
    assert brk.prop_type == BreakType.PAGE_BEFORE
    assert brk._get("BreakType") == BreakType.PAGE_BEFORE
    assert brk.prop_style == "HTML"
    assert brk._get("PageDescName") == "HTML"

    brk = Breaks(type=BreakType.PAGE_BEFORE, style="HTML", num=8)
    assert brk.prop_type == BreakType.PAGE_BEFORE
    assert brk._get("BreakType") == BreakType.PAGE_BEFORE
    assert brk.prop_style == "HTML"
    assert brk._get("PageDescName") == "HTML"
    assert brk.prop_num == 8
    assert brk._get("PageNumberOffset") == 8

    brk = Breaks(type=BreakType.PAGE_BEFORE, num=8)
    assert brk.prop_type == BreakType.PAGE_BEFORE
    assert brk._get("BreakType") == BreakType.PAGE_BEFORE
    assert brk.prop_style == None
    assert brk.prop_num == None

    # PAGE_AFTER does not allow style or num
    brk = Breaks(type=BreakType.PAGE_AFTER, style="HTML", num=8)
    assert brk.prop_type == BreakType.PAGE_AFTER
    assert brk._get("BreakType") == BreakType.PAGE_AFTER
    assert brk.prop_style == None
    assert brk.prop_num == None


def test_default() -> None:
    # brk = cast(Breaks, Breaks.default)
    brk = Breaks.default
    assert brk.prop_type == BreakType.NONE
    assert brk.prop_style == None
    assert brk.prop_num == None


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
        p_len = len(para_text)
        Write.append_para(cursor=cursor, text="Starting here...")

        brk = Breaks(type=BreakType.PAGE_BEFORE)
        Write.append_para(cursor=cursor, text=para_text, styles=(brk,))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        pp = cast("ParagraphProperties", cursor)
        assert pp.BreakType == BreakType.PAGE_BEFORE
        cursor.gotoEnd(False)

        brk = Breaks(type=BreakType.PAGE_BEFORE, style="Right Page")
        Write.append_para(cursor=cursor, text=para_text, styles=(brk,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.BreakType == BreakType.PAGE_BEFORE
        assert pp.PageDescName == "Right Page"
        cursor.gotoEnd(False)

        brk = Breaks(type=BreakType.PAGE_BEFORE, style="Right Page", num=5)
        Write.append_para(cursor=cursor, text=para_text, styles=(brk,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.BreakType == BreakType.PAGE_BEFORE
        assert pp.PageDescName == "Right Page"
        assert pp.PageNumberOffset == 5
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
