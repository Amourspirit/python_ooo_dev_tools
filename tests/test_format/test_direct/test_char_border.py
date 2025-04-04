from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.char.borders import Borders, BorderLineKind, Side, LineSize
from ooodev.format import CommonColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.format.writer.direct.char.font import (
    Font,
    FontLine,
    FontUnderlineEnum,
    FontFamilyEnum,
)

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_char_borders(loader) -> None:
    delay = 0  # 0 if Lo.bridge_connector.headless else 5_000
    from ooodev.office.write import Write

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        # note: If the  first word of the document is styled then it also will set the paragraph style
        # this will result in the style automatically being applied to new pargraphs as they are written.
        cursor = Write.get_cursor(doc)
        Write.append(cursor, "Starting")
        pos = Write.get_position(cursor)
        Write.append(cursor, "Hello")
        ft = Font(
            name="Lucida Console",
            overline=FontLine(line=FontUnderlineEnum.BOLDWAVE, color=CommonColor.CHARTREUSE),
            size=40,
            family=FontFamilyEnum.SCRIPT,
        )
        side = Side(color=CommonColor.RED)
        border = Borders(all=side)
        Write.style_left(cursor=cursor, pos=pos, styles=(border, ft))
        cursor.gotoEnd(False)
        cursor.goLeft(5, True)
        cp = cast("CharacterProperties", cursor)
        assert side == cp.CharLeftBorder
        assert side == cp.CharRightBorder
        assert side == cp.CharTopBorder
        assert side == cp.CharBottomBorder

        cursor.gotoEnd(False)

        # cursor.goLeft(5, True)
        # Style.apply_style(cursor, style)
        # cursor.gotoEnd(False)
        Write.append_para(cursor, " Non-Formatted")

        # cursor.ParaStyleName = "Default Paragraph Style"
        side = Side(color=CommonColor.BLUE, width=5.0)
        border = Borders(all=side)
        pos = Write.get_position(cursor)
        Write.append(cursor, "World")
        Write.style_left(cursor=cursor, pos=pos, styles=(border,))
        cp = cast("CharacterProperties", cursor)
        cursor.gotoEnd(False)
        cursor.goLeft(5, True)
        assert side == cp.CharLeftBorder
        assert side == cp.CharRightBorder
        assert side == cp.CharTopBorder
        assert side == cp.CharBottomBorder

        cursor.gotoEnd(False)

        Write.end_paragraph(cursor)
        # reset the paragraph style
        Write.style_left(cursor=cursor, pos=0, prop_name="ParaStyleName", prop_val="Standard")

        # using 1.05 for this test. LibreOffice chnages 1.1 to 1.05 in  this case.
        side = Side(color=CommonColor.DARK_ORANGE, line=BorderLineKind.DOUBLE, width=1.05)
        border = Borders(all=side)
        ft = Font(
            size=30.0,
            b=True,
            i=True,
            u=True,
            color=CommonColor.BLUE,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.GREEN),
        )
        Write.append(cursor=cursor, text="Nice Day", styles=(border, ft))
        cursor.gotoEnd(False)
        cursor.goLeft(5, True)
        assert side == cp.CharLeftBorder
        assert side == cp.CharRightBorder
        assert side == cp.CharTopBorder
        assert side == cp.CharBottomBorder
        cursor.gotoEnd(False)

        Write.end_paragraph(cursor)
        # reset the paragraph style
        Write.style_left(cursor=cursor, pos=0, prop_name="ParaStyleName", prop_val="Standard")

        txt = "adding\nmultiple\nlines"
        side = Side(color=CommonColor.CADET_BLUE, width=LineSize.MEDIUM)
        Write.append(cursor=cursor, text=txt, styles=(Borders(all=side),))

        cursor.gotoEnd(False)
        cursor.goLeft(len(txt), True)
        assert side == cp.CharLeftBorder
        assert side == cp.CharRightBorder
        assert side == cp.CharTopBorder
        assert side == cp.CharBottomBorder
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
