from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.para.line_spacing import LineSpacing, ModeKind
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_props() -> None:
    idt = LineSpacing(mode=ModeKind.SINGLE)
    ls_info = idt._get_style("line_spacing")
    ls, gen = ls_info
    assert ls.prop_height == 100
    assert ls.prop_mode == int(ModeKind.SINGLE)
    assert "keys" in gen.kwargs

    idt = LineSpacing(mode=ModeKind.LINES_1_15)
    ls_info = idt._get_style("line_spacing")
    ls, gen = ls_info
    assert ls.prop_height == 115
    assert ls.prop_mode == int(ModeKind.LINES_1_15)
    assert "keys" in gen.kwargs

    idt = LineSpacing(mode=ModeKind.LINES_15)
    ls_info = idt._get_style("line_spacing")
    ls, gen = ls_info
    assert ls.prop_height == 150
    assert ls.prop_mode == int(ModeKind.LINES_15)
    assert "keys" in gen.kwargs

    idt = LineSpacing(mode=ModeKind.DOUBLE)
    ls_info = idt._get_style("line_spacing")
    ls, gen = ls_info
    assert ls.prop_height == 200
    assert ls.prop_mode == int(ModeKind.DOUBLE)
    assert "keys" in gen.kwargs

    idt = LineSpacing(mode=ModeKind.PORPORTINAL, value=98)
    ls_info = idt._get_style("line_spacing")
    ls, gen = ls_info
    assert ls.prop_height == 98
    assert ls.prop_mode == int(ModeKind.PORPORTINAL)
    assert "keys" in gen.kwargs

    idt = LineSpacing(mode=ModeKind.AT_LEAST, value=1.0)
    ls_info = idt._get_style("line_spacing")
    ls, gen = ls_info
    assert ls.prop_height == 100
    assert ls.prop_mode == int(ModeKind.AT_LEAST)
    assert "keys" in gen.kwargs

    idt = LineSpacing(mode=ModeKind.LEADING, value=1.0)
    ls_info = idt._get_style("line_spacing")
    ls, gen = ls_info
    assert ls.prop_height == 100
    assert ls.prop_mode == int(ModeKind.LEADING)
    assert "keys" in gen.kwargs

    idt = LineSpacing(mode=ModeKind.FIXED, value=1.0)
    ls_info = idt._get_style("line_spacing")
    ls, gen = ls_info
    assert ls.prop_height == 100
    assert ls.prop_mode == int(ModeKind.FIXED)
    assert "keys" in gen.kwargs


def test_default() -> None:
    idt = LineSpacing.default
    ls_info = idt._get_style("line_spacing")
    ls, gen = ls_info
    assert ls.prop_height == 100
    assert ls.prop_mode == int(ModeKind.SINGLE)
    assert "keys" in gen.kwargs


def test_type_error() -> None:
    with pytest.raises(TypeError):
        _ = LineSpacing(mode=ModeKind.PORPORTINAL)

    with pytest.raises(TypeError):
        _ = LineSpacing(mode=ModeKind.AT_LEAST)

    with pytest.raises(TypeError):
        _ = LineSpacing(mode=ModeKind.LEADING)

    with pytest.raises(TypeError):
        _ = LineSpacing(mode=ModeKind.FIXED)


def test_write(loader, para_text) -> None:
    delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        p_len = len(para_text)
        Write.append_para(cursor=cursor, text=para_text, styles=(LineSpacing(mode=ModeKind.SINGLE),))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        ls = pp.ParaLineSpacing
        assert ls.Mode == int(ModeKind.SINGLE)
        assert ls.Height == 100
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineSpacing(mode=ModeKind.LINES_1_15),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Mode == int(ModeKind.LINES_1_15)
        assert ls.Height == 115
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineSpacing(mode=ModeKind.LINES_15),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Mode == int(ModeKind.LINES_15)
        assert ls.Height == 150
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineSpacing(mode=ModeKind.DOUBLE),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Mode == int(ModeKind.DOUBLE)
        assert ls.Height == 200
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineSpacing(mode=ModeKind.PORPORTINAL, value=96),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Mode == int(ModeKind.PORPORTINAL)
        assert ls.Height == 96
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineSpacing(mode=ModeKind.AT_LEAST, value=1.0),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Mode == int(ModeKind.AT_LEAST)
        assert ls.Height in range(99, 102)  # 99 - 101
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineSpacing(mode=ModeKind.LEADING, value=1.0),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Mode == int(ModeKind.LEADING)
        assert ls.Height in range(99, 102)  # 99 - 101
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(LineSpacing(mode=ModeKind.FIXED, value=5.0),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Mode == int(ModeKind.FIXED)
        assert ls.Height in range(498, 503)  # 498 - 502
        cursor.gotoEnd(False)

        Write.append_para(
            cursor=cursor, text=para_text, styles=(LineSpacing(mode=ModeKind.LINES_1_15, active_ln_spacing=True),)
        )
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Mode == int(ModeKind.LINES_1_15)
        assert ls.Height == 115
        assert pp.ParaRegisterModeActive
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
