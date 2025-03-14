from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.para.text_flow import FlowOptions
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_props() -> None:
    flo = FlowOptions(orphans=3)
    assert flo.prop_orphans == 3
    assert flo._get("ParaOrphans") == 3

    flo = FlowOptions(widows=4)
    assert flo.prop_widows == 4
    assert flo._get("ParaWidows") == 4

    flo = FlowOptions(keep=True)
    assert flo.prop_keep
    assert flo._get("ParaKeepTogether")

    flo = FlowOptions(no_split=True)
    assert flo.prop_no_split
    assert flo._get("ParaSplit") == False

    flo = FlowOptions(no_split=False)
    assert flo.prop_no_split == False
    assert flo._get("ParaSplit") == True

    # no_split is ommited when orphans or windows
    flo = FlowOptions(orphans=3, no_split=False)
    assert flo.prop_orphans == 3
    assert flo.prop_no_split is None
    assert flo._get("ParaOrphans") == 3

    flo = FlowOptions(widows=4, no_split=True)
    assert flo.prop_widows == 4
    assert flo.prop_no_split is None
    assert flo._get("ParaWidows") == 4

    flo = FlowOptions(orphans=3, widows=5, no_split=False)
    assert flo.prop_orphans == 3
    assert flo.prop_widows == 5
    assert flo.prop_no_split is None


def test_default() -> None:
    flo = FlowOptions().default
    assert flo.prop_orphans == 2
    assert flo.prop_widows == 2
    assert flo.prop_keep == False
    assert flo.prop_no_split == False


def test_keep() -> None:
    flo = FlowOptions().keep
    assert flo.prop_keep


def test_no_split() -> None:
    flo = FlowOptions().no_split
    assert flo.prop_no_split


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
        flo = FlowOptions(orphans=4)
        Write.append_para(cursor=cursor, text=para_text, styles=(flo,))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        pp = cast("ParagraphProperties", cursor)
        assert pp.ParaOrphans == 4
        cursor.gotoEnd(False)

        flo = FlowOptions(widows=3)
        Write.append_para(cursor=cursor, text=para_text, styles=(flo,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaWidows == 3
        cursor.gotoEnd(False)

        flo = FlowOptions(keep=True)
        Write.append_para(cursor=cursor, text=para_text, styles=(flo,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaKeepTogether
        cursor.gotoEnd(False)

        flo = FlowOptions(keep=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(flo,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaKeepTogether == False
        cursor.gotoEnd(False)

        flo = FlowOptions(no_split=True)
        Write.append_para(cursor=cursor, text=para_text, styles=(flo,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaKeepTogether == False
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
