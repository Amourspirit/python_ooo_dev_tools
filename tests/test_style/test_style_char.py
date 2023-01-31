from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.style.char import Char, StyleCharKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


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

        txt = "Hello "
        Write.append(cursor=cursor, text=txt)
        sc = Char(name=StyleCharKind.EXAMPLE)
        Write.append(cursor=cursor, text="World", styles=(sc,))
        cursor.goLeft(5, False)
        cursor.goRight(5, True)
        cp = cast("CharacterProperties", cursor)
        assert cp.CharStyleName == str(StyleCharKind.EXAMPLE)
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Write.append(cursor=cursor, text="What a ")
        sc = Char().source_text
        Write.append_para(cursor=cursor, text="World", styles=(sc,))
        cursor.goLeft(6, False)
        cursor.goRight(5, True)
        assert cp.CharStyleName == str(StyleCharKind.SOURCE_TEXT)
        Char.default.apply(cursor)
        assert cp.CharStyleName == ""
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
