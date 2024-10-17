from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.para.borders import (
    Borders,
    Side,
    BorderLineKind,
    Shadow,
    ShadowLocation,
    Padding,
    LineSize,
)
from ooodev.format import CommonColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
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

        side = Side(line=BorderLineKind.DOUBLE, width=0.75)
        bdr = Borders(all=side)
        bdr_default = bdr.default.copy()
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        pp = cast("ParagraphProperties", cursor)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.DASH_DOT, color=CommonColor.DARK_RED)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.DOUBLE_THIN, color=CommonColor.DARK_RED, width=LineSize.MEDIUM)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.THINTHICK_SMALLGAP, color=CommonColor.DARK_RED)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.THINTHICK_MEDIUMGAP, color=CommonColor.DARK_RED)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.THINTHICK_LARGEGAP, color=CommonColor.DARK_RED)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.THICKTHIN_SMALLGAP, color=CommonColor.BLUE_VIOLET)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.THICKTHIN_MEDIUMGAP, color=CommonColor.BLUE_VIOLET)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.THICKTHIN_LARGEGAP, color=CommonColor.BROWN)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.ENGRAVED, color=CommonColor.CADET_BLUE)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.OUTSET, color=CommonColor.DARK_GREEN)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.INSET, color=CommonColor.DARK_GREEN)
        bdr = Borders(all=side, merge=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(bdr,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert side == pp.LeftBorder
        assert side == pp.TopBorder
        assert side == pp.RightBorder
        assert side == pp.BottomBorder
        assert pp.ParaIsConnectBorder == False
        cursor.gotoEnd(False)
        bdr_default.apply(cursor)

        side = Side(line=BorderLineKind.DOUBLE, color=CommonColor.GREEN)
        shadow = Shadow(location=ShadowLocation.BOTTOM_RIGHT)
        bdr = Borders(all=side, shadow=shadow, merge=True)
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
        bdr_default.apply(cursor)

        bdr = Borders(
            all=Side(line=BorderLineKind.DOUBLE_THIN, color=CommonColor.BLUE),
            shadow=Shadow(location=ShadowLocation.BOTTOM_RIGHT),
            padding=Padding(left=2.0, right=1.5, top=3.1, bottom=4.2),
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
