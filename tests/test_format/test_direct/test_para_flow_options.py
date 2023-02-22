from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.para.text_flow import InnerFlowOptions
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_props() -> None:
    flo = InnerFlowOptions(orphans=3)
    assert flo.prop_orphans == 3
    assert flo._get("ParaOrphans") == 3

    flo = InnerFlowOptions(widows=4)
    assert flo.prop_widows == 4
    assert flo._get("ParaWidows") == 4

    flo = InnerFlowOptions(keep=True)
    assert flo.prop_keep
    assert flo._get("ParaKeepTogether")

    flo = InnerFlowOptions(no_split=True)
    assert flo.prop_no_split
    assert flo._get("ParaSplit") == False

    flo = InnerFlowOptions(no_split=False)
    assert flo.prop_no_split == False
    assert flo._get("ParaSplit") == True

    # no_split is ommited when orphans or windows
    flo = InnerFlowOptions(orphans=3, no_split=False)
    assert flo.prop_orphans == 3
    assert flo.prop_no_split is None
    assert flo._get("ParaOrphans") == 3

    flo = InnerFlowOptions(widows=4, no_split=True)
    assert flo.prop_widows == 4
    assert flo.prop_no_split is None
    assert flo._get("ParaWidows") == 4

    flo = InnerFlowOptions(orphans=3, widows=5, no_split=False)
    assert flo.prop_orphans == 3
    assert flo.prop_widows == 5
    assert flo.prop_no_split is None


def test_default() -> None:
    flo = InnerFlowOptions.default
    assert flo.prop_orphans == 2
    assert flo.prop_widows == 2
    assert flo.prop_keep == False
    assert flo.prop_no_split == False


def test_keep() -> None:
    flo = InnerFlowOptions().keep
    assert flo.prop_keep


def test_no_split() -> None:
    flo = InnerFlowOptions().no_split
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
        flo = InnerFlowOptions(orphans=4)
        Write.append_para(cursor=cursor, text=para_text, styles=(flo,))

        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        pp = cast("ParagraphProperties", cursor)
        assert pp.ParaOrphans == 4
        cursor.gotoEnd(False)

        flo = InnerFlowOptions(widows=3)
        Write.append_para(cursor=cursor, text=para_text, styles=(flo,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaWidows == 3
        cursor.gotoEnd(False)

        flo = InnerFlowOptions(keep=True)
        Write.append_para(cursor=cursor, text=para_text, styles=(flo,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaKeepTogether
        cursor.gotoEnd(False)

        flo = InnerFlowOptions(keep=False)
        Write.append_para(cursor=cursor, text=para_text, styles=(flo,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaKeepTogether == False
        cursor.gotoEnd(False)

        flo = InnerFlowOptions(no_split=True)
        Write.append_para(cursor=cursor, text=para_text, styles=(flo,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaKeepTogether == False
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
