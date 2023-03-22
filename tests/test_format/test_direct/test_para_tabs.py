from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.para.tabs import Tabs, TabAlign, FillCharKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_write(loader, para_text) -> None:
    # Tabs inherits from Tab and tab is tested in test_struct_tab
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        txt = "\t" + para_text
        p_len = len(txt)

        pos = 8.0
        Write.append_para(cursor=cursor, text=txt, styles=(Tabs(position=pos, fill_char="^"),))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        tb = Tabs.find(cursor, pos)
        assert tb is not None
        assert tb.prop_fill_char == "^"
        cursor.gotoEnd(False)

        pos = 14.5
        tb = Tabs(position=pos, align=TabAlign.DECIMAL, decimal_char=",")
        Write.append_para(cursor=cursor, text=txt, styles=(tb,))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        tb = Tabs.find(cursor, pos)
        ts_len = len(pp.ParaTabStops)
        assert tb is not None
        assert tb.prop_decimal_char == ","
        assert tb.prop_align == TabAlign.DECIMAL

        # update existing tabstop (based on position)
        pos = 14.5
        tb = Tabs.find(cursor, pos)
        assert tb is not None
        tb.prop_fill_char = FillCharKind.UNDER_SCORE
        tb.prop_align = TabAlign.RIGHT
        tb.apply(cursor)
        assert len(pp.ParaTabStops) == ts_len
        tb = Tabs.find(cursor, pos)
        assert tb is not None
        assert tb.prop_fill_char == str(FillCharKind.UNDER_SCORE)
        assert tb.prop_align == TabAlign.RIGHT

        # remove tabstop
        result = Tabs.remove(cursor, tb)
        assert result
        assert len(pp.ParaTabStops) == ts_len - 1
        ts_len = len(pp.ParaTabStops)
        cursor.gotoEnd(False)

        # remove all tabs
        Tabs.remove_all(cursor)
        tb = Tabs.from_obj(cursor, 0)
        assert tb is not None
        assert len(pp.ParaTabStops) == 1
        Write.append_para(cursor=cursor, text=para_text)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
