from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.style import BulletList, StyleListKind
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


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
        sl = BulletList(name=StyleListKind.LIST_01)
        sl.apply(cursor)

        start_pos = 0
        for i in range(1, 11):
            Write.append_para(cursor=cursor, text=f"Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos, True)
        pp = cast("ParagraphProperties", cursor)
        assert pp.NumberingStyleName == StyleListKind.LIST_01.value
        cursor.gotoEnd(False)
        BulletList.default.apply(cursor)
        Write.append_para(cursor=cursor, text="Moving On...")

        sl = BulletList().list_02
        start_pos = Write.get_position(cursor)
        sl.apply(cursor)
        for i in range(1, 11):
            Write.append_para(cursor=cursor, text=f"Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.LIST_02.value
        cursor.gotoEnd(False)
        BulletList.default.apply(cursor)
        Write.append_para(cursor=cursor, text="Moving On...")

        sl = BulletList().list_03
        start_pos = Write.get_position(cursor)
        sl.apply(cursor)
        for i in range(1, 11):
            Write.append_para(cursor=cursor, text=f"Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.LIST_03.value
        cursor.gotoEnd(False)
        BulletList.default.apply(cursor)
        Write.append_para(cursor=cursor, text="Moving On...")

        sl = BulletList(name=StyleListKind.NUM_123)
        start_pos = Write.get_position(cursor)
        sl.apply(cursor)
        for i in range(1, 11):
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_123.value
        cursor.gotoEnd(False)
        BulletList.default.apply(cursor)
        Write.append_para(cursor=cursor, text="Moving On...")

        sl = BulletList(name=StyleListKind.NUM_ABC)
        start_pos = Write.get_position(cursor)
        sl.apply(cursor)
        for i in range(1, 11):
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_ABC.value
        cursor.gotoEnd(False)
        BulletList.default.apply(cursor)
        Write.append_para(cursor=cursor, text="Moving On...")

        sl = BulletList().num_ivx
        start_pos = Write.get_position(cursor)
        sl.apply(cursor)
        for i in range(1, 11):
            Write.append_para(cursor=cursor, text=f"Num Point {i}")
        end_pos = Write.get_position(cursor)
        cursor.goLeft(end_pos - start_pos, True)
        assert pp.NumberingStyleName == StyleListKind.NUM_ivx.value
        cursor.gotoEnd(False)
        BulletList.default.apply(cursor)
        Write.append_para(cursor=cursor, text="Moving On...")

        style = BulletList(name=StyleListKind.NUM_123)
        xprops = style.get_style_props()
        assert xprops is not None

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
