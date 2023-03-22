from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno

# from ooodev.format.inner.direct.write.para.outline_list import OutlineList, LevelKind, StyleListKind
from ooodev.format.writer.direct.para.outline_list import Outline, LineNum, ListStyle, LevelKind, StyleListKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_write(loader, para_text) -> None:
    # minimal testing is fine here as each part of OutlineList is tested via Outline, ListStyle and LineNum test.
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
        # Test Outline
        Write.append_para(cursor=cursor, text=para_text, styles=(Outline(level=LevelKind.LEVEL_01),))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        pp.OutlineLevel == int(LevelKind.LEVEL_01)
        cursor.gotoEnd(False)

        # test ListStyle
        start_pos = Write.get_position(cursor)
        ls = ListStyle(list_style=StyleListKind.NUM_123)
        ls.apply(cursor)
        for i in range(1, 6):
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_123.value
        cursor.gotoEnd(False)

        ls.default.apply(cursor)

        # test LineNum
        Write.append_para(cursor=cursor, text=para_text, styles=(LineNum(num_start=6),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaLineNumberCount == True
        assert pp.ParaLineNumberStartValue == 6
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
