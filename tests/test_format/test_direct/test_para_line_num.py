from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.direct.para.outline_list import LineNum
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_props() -> None:
    ln = LineNum()
    assert ln._get("ParaLineNumberStartValue") == 0
    assert ln._get("ParaLineNumberCount") == True
    assert ln.prop_num_start == 0

    ln = LineNum(-1)
    assert ln._get("ParaLineNumberStartValue") == 0
    assert ln._get("ParaLineNumberCount") == False
    assert ln.prop_num_start == 0

    ln = LineNum(num_start=1)
    assert ln._get("ParaLineNumberStartValue") == 1
    assert ln._get("ParaLineNumberCount") == True
    assert ln.prop_num_start == 1

    ln = LineNum(num_start=5)
    assert ln._get("ParaLineNumberStartValue") == 5
    assert ln._get("ParaLineNumberCount") == True
    assert ln.prop_num_start == 5


def test_default() -> None:
    ln = LineNum().default
    assert ln._get("ParaLineNumberStartValue") == 0
    assert ln._get("ParaLineNumberCount") == True
    assert ln.prop_num_start == 0


def test_include() -> None:
    ln = LineNum().include
    assert ln._get("ParaLineNumberCount") == True

    # if is already include then it value should not change
    ln = LineNum(3).include
    assert ln._get("ParaLineNumberCount") == True
    assert ln.prop_num_start == 3


def test_exclude() -> None:
    ln = LineNum().exclude
    assert ln._get("ParaLineNumberCount") == False


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
        Write.append_para(cursor=cursor, text="Start paragraph ...")
        Write.append_para(cursor=cursor, text=para_text, styles=(LineNum(0),))

        pp = cast("ParagraphProperties", cursor)
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaLineNumberCount
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineNum(-1),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaLineNumberCount == False
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineNum(6),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaLineNumberCount == True
        assert pp.ParaLineNumberStartValue == 6
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineNum().exclude,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaLineNumberCount == False
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineNum().include,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaLineNumberCount == True
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineNum().restart_numbers,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaLineNumberCount == True
        assert pp.ParaLineNumberStartValue == 1
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineNum().default,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaLineNumberCount == True
        assert pp.ParaLineNumberStartValue == 0
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
