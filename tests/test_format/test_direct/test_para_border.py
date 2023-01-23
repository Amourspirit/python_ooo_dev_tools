from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.direct.para.borders import (
    Borders,
    Side,
    BorderLineStyleEnum,
    BorderShadow,
    ShadowLocation,
    BorderPadding,
    LineSize,
)
from ooodev.format import CommonColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


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

        side = Side(line=BorderLineStyleEnum.DOUBLE, width=0.75)
        bdr = Borders(border_side=side)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        pp = cast("ParagraphProperties", cursor)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.DASH_DOT, color=CommonColor.DARK_RED)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.DOUBLE_THIN, color=CommonColor.DARK_RED, width=LineSize.MEDIUM)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.THINTHICK_SMALLGAP, color=CommonColor.DARK_RED)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.THINTHICK_MEDIUMGAP, color=CommonColor.DARK_RED)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.THINTHICK_LARGEGAP, color=CommonColor.DARK_RED)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.THICKTHIN_SMALLGAP, color=CommonColor.BLUE_VIOLET)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.THICKTHIN_MEDIUMGAP, color=CommonColor.BLUE_VIOLET)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.THICKTHIN_LARGEGAP, color=CommonColor.BROWN)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.ENGRAVED, color=CommonColor.CADET_BLUE)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.OUTSET, color=CommonColor.DARK_GREEN)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.INSET, color=CommonColor.DARK_GREEN)
        bdr = Borders(border_side=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        side = Side(line=BorderLineStyleEnum.DOUBLE, color=CommonColor.GREEN)
        shadow = BorderShadow(location=ShadowLocation.BOTTOM_RIGHT)
        bdr = Borders(border_side=side, shadow=shadow, merge=True)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert shadow == pp.ParaShadowFormat
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder
        cursor.gotoEnd(False)
        Borders.default.apply(cursor)

        bdr = Borders(
            border_side=Side(line=BorderLineStyleEnum.DOUBLE_THIN, color=CommonColor.BLUE),
            shadow=BorderShadow(location=ShadowLocation.BOTTOM_RIGHT),
            padding=BorderPadding(left=2.0, right=1.5, top=3.1, bottom=4.2),
        )
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        # assert pp.BreakType == BreakType.PAGE_BEFORE
        # assert pp.PageDescName == "Right Page"
        # assert pp.PageNumberOffset == 5
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
