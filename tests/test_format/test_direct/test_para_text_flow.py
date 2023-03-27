from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno

# from ooodev.format.inner.direct.write.para.text_flow import TextFlow, BreakType
from ooodev.format.writer.direct.para.text_flow import BreakType, Breaks, FlowOptions, Hyphenation
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_write(loader, para_text) -> None:
    # minimal testing is fine here as each part of TextFlow is tested via Breaks, Hyphenation and FlowOptions test.
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
        # test FlowOptions
        Write.append_para(cursor=cursor, text=para_text, styles=(FlowOptions(orphans=4),))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        pp = cast("ParagraphProperties", cursor)
        assert pp.ParaOrphans == 4
        cursor.gotoEnd(False)

        # test Hyphenation
        Write.append_para(cursor=cursor, text=para_text, styles=(Hyphenation(auto=True, no_caps=True),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaIsHyphenation
        assert pp.ParaHyphenationNoCaps
        cursor.gotoEnd(False)

        # test Breaks
        Write.append_para(cursor=cursor, text=para_text, styles=(Breaks(type=BreakType.PAGE_BEFORE),))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.BreakType == BreakType.PAGE_BEFORE
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
