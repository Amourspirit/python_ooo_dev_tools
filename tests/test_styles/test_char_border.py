from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.char.borders import Borders, BorderLineStyleEnum, Side
from ooodev.styles import CommonColor
from ooodev.styles.style_const import POINT_RATIO
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_char_borders(loader, test_headless) -> None:
    delay = 0 if test_headless else 5_000
    from ooodev.office.write import Write
    from ooodev.styles import Style

    doc = Write.create_doc()
    if not test_headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        Write.append(cursor, "Hello")
        style = Borders(border_side=Side(color=CommonColor.RED))
        Write.style_left(cursor=cursor, pos=5, style=style)
        # cursor.goLeft(5, True)
        # Style.apply_style(cursor, style)
        # cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        style = Borders(border_side=Side(color=CommonColor.BLUE, width=5.0))
        Write.append(cursor, "World")
        Write.style_left(cursor=cursor, pos=5, style=style)
        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
