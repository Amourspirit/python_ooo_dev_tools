from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.para.indent_spacing import IndentSpacing, ModeKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_write(loader, para_text) -> None:
    # minimal testing is fine here as each part of IndentSpacing is tested via Indent, Spacing and Line Spacing test.
    delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        p_len = len(para_text)
        amt = 3.5
        # test Indent
        Write.append_para(cursor=cursor, text=para_text, styles=(IndentSpacing(before=amt),))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        assert pp.ParaLeftMargin in [round(amt * 100) - 2 + i for i in range(5)]  # plus or minus 2
        cursor.gotoEnd(False)

        # test LineSpacing
        Write.append_para(cursor=cursor, text=para_text, styles=(IndentSpacing(mode=ModeKind.PORPORTINAL, value=96),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Mode == int(ModeKind.PORPORTINAL)
        assert ls.Height == 96
        cursor.gotoEnd(False)

        # test Spacing
        amt = 2.0
        Write.append_para(cursor=cursor, text=para_text, styles=(IndentSpacing(below=amt),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaBottomMargin in [round(amt * 100) - 2 + i for i in range(5)]  # plus or minus 2
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
