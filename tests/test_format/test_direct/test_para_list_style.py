from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.direct.para.outline_list import ListStyle, StyleListKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_props() -> None:
    ls = ListStyle(list_style=StyleListKind.LIST_01)
    assert ls.prop_list_style == StyleListKind.LIST_01.value
    assert ls._get("NumberingStyleName") == StyleListKind.LIST_01.value
    assert ls.prop_num_start is None

    ls = ListStyle(list_style=StyleListKind.NONE)
    assert ls.prop_list_style == ""
    assert ls._get("NumberingStyleName") == ""
    assert ls.prop_num_start == -1
    assert ls._get("ParaIsNumberingRestart") == False

    ls = ListStyle(list_style="", num_start=2)
    assert ls.prop_list_style == ""
    assert ls._get("NumberingStyleName") == ""
    assert ls.prop_num_start == -1
    assert ls._get("ParaIsNumberingRestart") == False

    ls = ListStyle(num_start=-1)
    assert ls.prop_num_start == -1
    assert ls._get("NumberingStartValue") == -1
    assert ls._get("ParaIsNumberingRestart") == False
    assert ls.prop_list_style is None

    ls = ListStyle(num_start=-2)
    assert ls.prop_num_start == -1
    assert ls._get("NumberingStartValue") == -1
    assert ls._get("ParaIsNumberingRestart") == True
    assert ls.prop_list_style is None

    ls = ListStyle(num_start=0)
    assert ls.prop_num_start == 0
    assert ls._get("NumberingStartValue") == 0
    assert ls._get("ParaIsNumberingRestart") == True
    assert ls.prop_list_style is None

    ls = ListStyle(num_start=5)
    assert ls.prop_num_start == 5
    assert ls._get("NumberingStartValue") == 5
    assert ls._get("ParaIsNumberingRestart") == True
    assert ls.prop_list_style is None


def test_default() -> None:
    # brk = cast(Breaks, Breaks.default)
    ls = ListStyle.default
    assert ls._get("NumberingStyleName") == ""
    assert ls._get("NumberingStartValue") == -1
    assert ls._get("ParaIsNumberingRestart") == False


def test_write(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        ls = ListStyle(list_style=StyleListKind.LIST_01)
        ls.apply(cursor)
        start_pos = 0
        for i in range(1, 6):
            Write.append_para(cursor=cursor, text=f"Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos, True)
        pp = cast("ParagraphProperties", cursor)
        assert pp.NumberingStyleName == StyleListKind.LIST_01.value
        cursor.gotoEnd(False)

        ListStyle.default.apply(cursor)
        Write.append_para(cursor, "Moving on...")

        start_pos = Write.get_position(cursor)
        ls = ListStyle(list_style=StyleListKind.NUM_123)
        ls.apply(cursor)
        for i in range(1, 6):
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_123.value
        cursor.gotoEnd(False)

        ListStyle.default.apply(cursor)
        Write.append_para(cursor, "Moving on...")

        start_pos = Write.get_position(cursor)
        ls.apply(cursor)
        for i in range(1, 4):
            if i == 1:
                ls.restart_numbers.apply(cursor)
                assert pp.ParaIsNumberingRestart == True
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_123.value
        cursor.gotoEnd(False)

        ls.restart_numbers.apply(cursor)
        start_pos = Write.get_position(cursor)
        for i in range(4, 8):
            if i == 4:
                ls.restart_numbers.apply(cursor)
                assert pp.ParaIsNumberingRestart == True
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_123.value
        cursor.gotoEnd(False)
        # restart numbering
        ls_rs = ls.fmt_num_start(-2)
        ls_rs.apply(cursor)
        assert pp.ParaIsNumberingRestart == True
        Write.append_para(cursor=cursor, text="Num Point 8")
        Write.append_para(cursor=cursor, text="Num Point 9")

        ListStyle.default.apply(cursor)
        Write.append_para(cursor, "Moving on...")

        start_pos = Write.get_position(cursor)
        ls = ListStyle(list_style=StyleListKind.NUM_ABC, num_start=-2)
        ls.apply(cursor)
        for i in range(1, 4):
            if i == 1:
                ls.restart_numbers.apply(cursor)
                assert pp.ParaIsNumberingRestart == True
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_ABC.value
        cursor.gotoEnd(False)

        ListStyle.default.apply(cursor)
        Write.append_para(cursor, "Moving on...")

        start_pos = Write.get_position(cursor)
        ls = ListStyle(list_style=StyleListKind.NUM_abc, num_start=-2)
        ls.apply(cursor)
        for i in range(1, 4):
            if i == 1:
                ls.restart_numbers.apply(cursor)
                assert pp.ParaIsNumberingRestart == True
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_abc.value
        cursor.gotoEnd(False)

        ListStyle.default.apply(cursor)
        Write.append_para(cursor, "Moving on...")

        start_pos = Write.get_position(cursor)
        ls = ListStyle(list_style=StyleListKind.NUM_IVX, num_start=-2)
        ls.apply(cursor)
        for i in range(1, 4):
            if i == 1:
                ls.restart_numbers.apply(cursor)
                assert pp.ParaIsNumberingRestart == True
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_IVX.value
        cursor.gotoEnd(False)

        ListStyle.default.apply(cursor)
        Write.append_para(cursor, "Moving on...")

        start_pos = Write.get_position(cursor)
        ls = ListStyle(list_style=StyleListKind.NUM_ivx, num_start=-2)
        ls.apply(cursor)
        for i in range(1, 4):
            if i == 1:
                ls.restart_numbers.apply(cursor)
                assert pp.ParaIsNumberingRestart == True
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_ivx.value
        cursor.gotoEnd(False)

        ListStyle.default.apply(cursor)
        Write.append_para(cursor, "Moving on...")

        start_pos = Write.get_position(cursor)
        ls = ListStyle(list_style=StyleListKind.LIST_02, num_start=-2)
        ls.apply(cursor)
        for i in range(1, 4):
            if i == 1:
                ls.restart_numbers.apply(cursor)
                assert pp.ParaIsNumberingRestart == True
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.LIST_02.value
        cursor.gotoEnd(False)

        ListStyle.default.apply(cursor)
        Write.append_para(cursor, "Moving on...")

        start_pos = Write.get_position(cursor)
        ls = ListStyle(list_style=StyleListKind.LIST_03, num_start=-2)
        ls.apply(cursor)
        for i in range(1, 4):
            if i == 1:
                ls.restart_numbers.apply(cursor)
                assert pp.ParaIsNumberingRestart == True
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.LIST_03.value
        cursor.gotoEnd(False)

        ListStyle.default.apply(cursor)
        Write.append_para(cursor, "Moving on...")

        start_pos = Write.get_position(cursor)
        ls = ListStyle(list_style=StyleListKind.LIST_04, num_start=-2)
        ls.apply(cursor)
        for i in range(1, 4):
            if i == 1:
                ls.restart_numbers.apply(cursor)
                assert pp.ParaIsNumberingRestart == True
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.LIST_04.value
        cursor.gotoEnd(False)

        ListStyle.default.apply(cursor)
        Write.append_para(cursor, "Moving on...")

        start_pos = Write.get_position(cursor)
        ls = ListStyle(list_style=StyleListKind.LIST_05, num_start=-2)
        ls.apply(cursor)
        for i in range(1, 4):
            if i == 1:
                ls.restart_numbers.apply(cursor)
                assert pp.ParaIsNumberingRestart == True
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.LIST_05.value
        cursor.gotoEnd(False)

        ListStyle.default.apply(cursor)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
