from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.char.borders import Borders, BorderLineStyleEnum, Side
from ooodev.styles import CommonColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.styles.char.font import (
    Font,
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
            overline=FontUnderlineEnum.BOLDWAVE,
            size=40,
            overline_color=CommonColor.CHARTREUSE,
            family=FontFamilyEnum.SCRIPT,
        )
        border = Borders(border_side=Side(color=CommonColor.RED))
        Write.style_left(cursor=cursor, pos=pos, styles=(border, ft))
        cursor.gotoEnd(False)
        cursor.goLeft(5, True)
        cp = cast("CharacterProperties", cursor)
        assert cp.CharLeftBorder.Color == CommonColor.RED
        assert cp.CharLeftBorder.LineWidth == 26

        assert cp.CharRightBorder.Color == CommonColor.RED
        assert cp.CharRightBorder.LineWidth == 26

        assert cp.CharTopBorder.Color == CommonColor.RED
        assert cp.CharTopBorder.LineWidth == 26

        assert cp.CharBottomBorder.Color == CommonColor.RED
        assert cp.CharBottomBorder.LineWidth == 26
        cursor.gotoEnd(False)

        # cursor.goLeft(5, True)
        # Style.apply_style(cursor, style)
        # cursor.gotoEnd(False)
        Write.append_para(cursor, " Non-Formatted")

        # cursor.ParaStyleName = "Default Paragraph Style"
        border = Borders(border_side=Side(color=CommonColor.BLUE, width=5.0))
        pos = Write.get_position(cursor)
        Write.append(cursor, "World")
        Write.style_left(cursor=cursor, pos=pos, styles=(border,))
        cp = cast("CharacterProperties", cursor)
        cursor.gotoEnd(False)
        cursor.goLeft(5, True)
        assert cp.CharLeftBorder.Color == CommonColor.BLUE
        assert cp.CharLeftBorder.LineWidth == 176

        assert cp.CharRightBorder.Color == CommonColor.BLUE
        assert cp.CharRightBorder.LineWidth == 176

        assert cp.CharTopBorder.Color == CommonColor.BLUE
        assert cp.CharTopBorder.LineWidth == 176

        assert cp.CharBottomBorder.Color == CommonColor.BLUE
        assert cp.CharBottomBorder.LineWidth == 176
        cursor.gotoEnd(False)

        Write.end_paragraph(cursor)
        # reset the paragraph style
        Write.style_left(cursor=cursor, pos=0, prop_name="ParaStyleName", prop_val="Standard")

        border = Borders(border_side=Side(color=CommonColor.DARK_ORANGE, line=BorderLineStyleEnum.DOUBLE, width=1.1))
        ft = Font(size=30.0, b=True, i=True, u=True, color=CommonColor.BLUE, underine_color=CommonColor.GREEN)
        Write.append(cursor=cursor, text="Nice Day", styles=(border, ft))
        cursor.gotoEnd(False)
        cursor.goLeft(5, True)
        assert cp.CharLeftBorder.Color == CommonColor.DARK_ORANGE
        assert cp.CharLeftBorder.LineWidth == 39
        assert cp.CharLeftBorder.LineStyle == BorderLineStyleEnum.DOUBLE.value

        assert cp.CharRightBorder.Color == CommonColor.DARK_ORANGE
        assert cp.CharRightBorder.LineWidth == 39
        assert cp.CharRightBorder.LineStyle == BorderLineStyleEnum.DOUBLE.value

        assert cp.CharTopBorder.Color == CommonColor.DARK_ORANGE
        assert cp.CharTopBorder.LineWidth == 39
        assert cp.CharTopBorder.LineStyle == BorderLineStyleEnum.DOUBLE.value

        assert cp.CharBottomBorder.Color == CommonColor.DARK_ORANGE
        assert cp.CharBottomBorder.LineWidth == 39
        assert cp.CharBottomBorder.LineStyle == BorderLineStyleEnum.DOUBLE.value
        cursor.gotoEnd(False)

        Write.end_paragraph(cursor)
        # reset the paragraph style
        Write.style_left(cursor=cursor, pos=0, prop_name="ParaStyleName", prop_val="Standard")

        txt = "adding\nmultiple\nlines"
        Write.append(
            cursor=cursor,
            text=txt,
            styles=(Borders(border_side=Side(color=CommonColor.CADET_BLUE, width=1.3)),),
        )

        cursor.gotoEnd(False)
        cursor.goLeft(len(txt), True)
        assert cp.CharLeftBorder.Color == CommonColor.CADET_BLUE
        assert cp.CharLeftBorder.LineWidth == 46

        assert cp.CharRightBorder.Color == CommonColor.CADET_BLUE
        assert cp.CharRightBorder.LineWidth == 46

        assert cp.CharTopBorder.Color == CommonColor.CADET_BLUE
        assert cp.CharTopBorder.LineWidth == 46

        assert cp.CharBottomBorder.Color == CommonColor.CADET_BLUE
        assert cp.CharBottomBorder.LineWidth == 46
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
