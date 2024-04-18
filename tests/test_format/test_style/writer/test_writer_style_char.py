from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.style import Char, StyleCharKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write
from ooodev.utils.color import CommonColor


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
        # document = Lo.ThisComponent.CurrentController.Frame
        # cursor.PageDescName = "First Page"

        txt = "Hello "
        Write.append(cursor=cursor, text=txt)
        style = Char(name=StyleCharKind.EXAMPLE)
        Write.append(cursor=cursor, text="World", styles=(style,))
        cursor.goLeft(5, False)
        cursor.goRight(5, True)
        cp = cast("CharacterProperties", cursor)
        assert cp.CharStyleName == str(StyleCharKind.EXAMPLE)
        f_style = Char.from_obj(cursor)
        assert f_style.prop_name == style.prop_name
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Write.append(cursor=cursor, text="What a ")
        style = Char().source_text
        Write.append_para(cursor=cursor, text="World", styles=(style,))
        cursor.goLeft(6, False)
        cursor.goRight(5, True)
        assert cp.CharStyleName == str(StyleCharKind.SOURCE_TEXT)
        Char.default.apply(cursor)
        assert cp.CharStyleName == ""
        cursor.gotoEnd(False)

        style = Char(name=StyleCharKind.STANDARD)
        xprops = style.get_style_props()
        assert xprops is not None
        xprops.setPropertyValue("CharBackColor", CommonColor.CORAL)
        val = cast(int, xprops.getPropertyValue("CharBackColor"))
        assert val == CommonColor.CORAL

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
