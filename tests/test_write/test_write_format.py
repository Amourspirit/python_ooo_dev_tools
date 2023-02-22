import os
import pytest
from typing import cast, TYPE_CHECKING

# from ooodev.office.write import Write
if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.utils.lo import Lo
from ooodev.utils.gui import GUI
from ooodev.office.write import Write
from ooodev.format.writer.direct.char.highlight import InnerHighlight
from ooodev.format import CommonColor

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_format(loader, run_headless) -> None:
    # breakpoint()

    # delay = 0 if run_headless else 2_000
    delay = 0
    visible = not run_headless

    doc = Write.create_doc()
    try:
        # if not Lo.bridge_connector.headless:
        if visible:
            GUI.set_visible(True, doc)
            Lo.delay(500)
            GUI.zoom(GUI.ZoomEnum.ZOOM_200_PERCENT)

        cursor = Write.get_cursor(doc)
        hl = InnerHighlight(CommonColor.YELLOW_GREEN)
        Write.append(cursor=cursor, text="Highlighting starts ")
        pos = Write.get_position(cursor)
        Write.append_para(cursor=cursor, text="here.")
        Write.style(pos=pos, length=4, styles=(hl,))
        cp = cast("CharacterProperties", cursor)
        cursor.gotoStart(False)
        cursor.goRight(pos, False)
        cursor.goRight(4, True)
        assert cp.CharBackColor == CommonColor.YELLOW_GREEN
        cursor.gotoEnd(False)

        Lo.delay(delay)
        Write.style(pos=pos, length=4, styles=(InnerHighlight.empty,))
        cursor.gotoStart(False)
        cursor.goRight(pos, False)
        cursor.goRight(4, True)
        assert cp.CharBackColor == -1
        cursor.gotoEnd(False)

        Lo.delay(delay)
        Write.style(pos=pos, length=4, styles=(hl,), cursor=cursor)
        cursor.gotoStart(False)
        cursor.goRight(pos, False)
        cursor.goRight(4, True)
        assert cp.CharBackColor == CommonColor.YELLOW_GREEN
        cursor.gotoEnd(False)

        Lo.delay(delay)
        Write.style(pos=pos, length=4, prop_name="CharBackColor", prop_val=CommonColor.YELLOW_GREEN)
        cursor.gotoStart(False)
        cursor.goRight(pos, False)
        cursor.goRight(4, True)
        assert cp.CharBackColor == CommonColor.YELLOW_GREEN
        cursor.gotoEnd(False)

        Lo.delay(delay)
        Write.style(pos=pos, length=4, prop_name="CharBackColor", prop_val=CommonColor.LIGHT_BLUE, cursor=cursor)
        cursor.gotoStart(False)
        cursor.goRight(pos, False)
        cursor.goRight(4, True)
        assert cp.CharBackColor == CommonColor.LIGHT_BLUE
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_format_left(loader, run_headless) -> None:
    # breakpoint()

    # delay = 0 if run_headless else 2_000
    delay = 0
    visible = not run_headless

    doc = Write.create_doc()
    try:
        if visible:
            GUI.set_visible(True, doc)
            Lo.delay(500)
            GUI.zoom(GUI.ZoomEnum.ZOOM_200_PERCENT)

        cursor = Write.get_cursor(doc)
        cp = cast("CharacterProperties", cursor)
        hl = InnerHighlight(CommonColor.YELLOW_GREEN)
        Write.append(cursor=cursor, text="Highlighting starts ")
        pos = Write.get_position(cursor)  # 20
        # pos = 20
        pre_backcolor = cp.CharBackColor
        Write.append(cursor=cursor, text="here")
        Write.style_left(cursor=cursor, pos=pos, styles=(hl,))
        # check that property has been restored after style applied
        assert pre_backcolor == cp.CharBackColor
        Write.append_para(cursor=cursor, text=".")

        cursor.gotoStart(False)
        cursor.goRight(pos, False)
        cursor.goRight(4, True)
        assert cp.CharBackColor == CommonColor.YELLOW_GREEN
        cursor.gotoEnd(False)

        Lo.delay(delay)
        Write.append(cursor=cursor, text="And ")
        pre_backcolor = cp.CharBackColor
        # pos = Write.get_position(cursor) # 30
        pos = 30
        Write.append(cursor=cursor, text="here")
        Write.style_left(cursor=cursor, pos=pos, prop_name="CharBackColor", prop_val=CommonColor.YELLOW_GREEN)
        assert pre_backcolor == cp.CharBackColor
        Write.append_para(cursor=cursor, text=".")
        cursor.gotoStart(False)
        cursor.goRight(pos, False)
        cursor.goRight(4, True)
        assert cp.CharBackColor == CommonColor.YELLOW_GREEN
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
