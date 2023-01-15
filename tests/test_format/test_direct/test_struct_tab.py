from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.direct.structs.tab import Tab, TabAlign, FillCharKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_props() -> None:
    tab = Tab(position=10.0)
    assert tab.prop_position == 10.0
    assert tab._get("Position") == 1000

    tab = Tab(align=TabAlign.LEFT)
    assert tab.prop_align == TabAlign.LEFT
    assert tab._get("Alignment") == TabAlign.LEFT

    tab = Tab(align=TabAlign.DECIMAL, decimal_char="*")
    assert tab.prop_decimal_char == "*"
    assert tab._get("DecimalChar") == "*"

    tab = Tab(fill_char="#")
    assert tab.prop_fill_char == "#"
    assert tab._get("FillChar") == "#"

    tab = Tab()
    tab.prop_align = TabAlign.DEFAULT
    assert tab.prop_align == TabAlign.DEFAULT

    tab.prop_position = 10.0
    assert tab.prop_position == 10.0

    tab.prop_decimal_char = "#"
    assert tab.prop_decimal_char == "#"

    tab.prop_fill_char = FillCharKind.DASH
    assert tab.prop_fill_char == str(FillCharKind.DASH)


def test_write(loader, para_text) -> None:
    delay = 0 if Lo.bridge_connector.headless else 3_000

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
        Write.append_para(cursor=cursor, text=txt, styles=(Tab(position=pos, fill_char="^"),))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        # assert pp.ParaIsHyphenation
        rng = get_range(pos)
        idx = -1
        assert len(pp.ParaTabStops) == 2
        for i, ts in enumerate(pp.ParaTabStops):
            if ts.Position in rng:
                idx = i
                break
        assert idx > -1
        ts = pp.ParaTabStops[idx]
        assert ts.FillChar == "^"
        cursor.gotoEnd(False)

        pos = 0.0
        tb = Tab()
        Write.append_para(cursor=cursor, text="\t" + para_text, styles=(tb,))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        rng = get_range(pos)
        idx = -1
        assert len(pp.ParaTabStops) == 3
        for i, ts in enumerate(pp.ParaTabStops):
            if ts.Position in rng:
                idx = i
                break
        assert idx > -1
        ts = pp.ParaTabStops[idx]
        assert ts.FillChar == " "
        assert ts.Alignment == TabAlign.LEFT
        cursor.gotoEnd(False)

        pos = 14.5
        tb = Tab(position=pos, align=TabAlign.DECIMAL, decimal_char=",")
        Write.append_para(cursor=cursor, text="\t" + para_text, styles=(tb,))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        rng = get_range(pos)
        idx = -1
        assert len(pp.ParaTabStops) == 4
        for i, ts in enumerate(pp.ParaTabStops):
            if ts.Position in rng:
                idx = i
                break
        assert idx > -1
        ts = pp.ParaTabStops[idx]
        assert ts.DecimalChar == ","
        assert ts.Alignment == TabAlign.DECIMAL
        cursor.gotoEnd(False)

        # update existing tabstop (based on position)
        pos = 14.5
        tb = Tab(position=pos, align=TabAlign.RIGHT, fill_char="@")
        Write.append_para(cursor=cursor, text="\t" + para_text, styles=(tb,))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        rng = get_range(pos)
        idx = -1
        assert len(pp.ParaTabStops) == 4
        for i, ts in enumerate(pp.ParaTabStops):
            if ts.Position in rng:
                idx = i
                break
        assert idx > -1
        ts = pp.ParaTabStops[idx]
        assert ts.FillChar == "@"
        assert ts.Alignment == TabAlign.RIGHT
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def get_range(pos: float) -> range:
    i = round(pos * 100)
    return range(i - 2, i + 3)
