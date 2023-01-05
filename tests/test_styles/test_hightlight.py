from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.char.highlight import Highlight
from ooodev.styles import CommonColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write
from ooodev.styles.char.font import (
    Font,
    FontUnderlineEnum,
    FontFamilyEnum,
)

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_char_highlight_props() -> None:
    hl = Highlight(color=CommonColor.LIGHT_YELLOW)
    assert hl.prop_color == CommonColor.LIGHT_YELLOW
    assert hl._get("CharBackColor") == CommonColor.LIGHT_YELLOW
    assert hl._get("CharBackTransparent") == False

    hl.prop_color = -1
    assert hl.prop_color == -1
    assert hl._get("CharBackColor") == -1
    assert hl._get("CharBackTransparent")

    hl = Highlight()
    hl.prop_color = -1
    assert hl.prop_color == -1
    assert hl._get("CharBackColor") == -1
    assert hl._get("CharBackTransparent")

    hl.prop_color = CommonColor.AQUA
    assert hl.prop_color == CommonColor.AQUA
    assert hl._get("CharBackColor") == CommonColor.AQUA
    assert hl._get("CharBackTransparent") == False

    hl = Highlight.empty
    hl.color = -1
    assert hl.color == -1
    assert hl._get("CharBackColor") == -1
    assert hl._get("CharBackTransparent")


def test_char_hightlight(loader) -> None:
    # breakpoint()
    delay = 0  # 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        # note: If the  first word of the document is styled then it also will set the paragraph style
        # this will result in the style automatically being applied to new pargraphs as they are written.
        cursor = Write.get_cursor(doc)
        ft = Font(size=40)
        Write.append(cursor, "Starting ", (ft,))
        pos = Write.get_position(cursor)
        hl = Highlight(color=CommonColor.LIGHT_YELLOW)
        Write.append(cursor, "Hello")
        Write.style_left(cursor=cursor, pos=pos, styles=(hl, ft))
        cursor.gotoEnd(False)
        cursor.goLeft(5, True)
        cp = cast("CharacterProperties", cursor)
        assert cp.CharBackColor == CommonColor.LIGHT_YELLOW
        assert cp.CharBackTransparent == False
        cursor.gotoEnd(False)

        Write.style_left(cursor=cursor, pos=pos, styles=(Highlight.empty,))
        cursor.gotoEnd(False)
        cursor.goLeft(5, True)
        cp = cast("CharacterProperties", cursor)
        assert cp.CharBackColor == -1
        assert cp.CharBackTransparent == True
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
