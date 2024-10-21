from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

from ooodev.format.writer.direct.para.indent_space import LineSpacing, ModeKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_props() -> None:
    ls = LineSpacing(mode=ModeKind.SINGLE)
    assert ls.prop_value == 100
    assert ls.prop_mode == ModeKind.SINGLE

    ls = LineSpacing(mode=ModeKind.LINE_1_15)
    assert ls.prop_value == 115
    assert ls.prop_mode == ModeKind.LINE_1_15

    ls = LineSpacing(mode=ModeKind.LINE_1_5)
    assert ls.prop_value == 150
    assert ls.prop_mode == ModeKind.LINE_1_5

    ls = LineSpacing(mode=ModeKind.DOUBLE)
    assert ls.prop_value == 200
    assert ls.prop_mode == ModeKind.DOUBLE

    ls = LineSpacing(mode=ModeKind.PROPORTIONAL, value=98)
    assert ls.prop_value == 98
    assert ls.prop_mode == ModeKind.PROPORTIONAL

    ls = LineSpacing(mode=ModeKind.AT_LEAST, value=1.0)
    assert ls.prop_value == 100
    assert ls.prop_mode == ModeKind.AT_LEAST

    ls = LineSpacing(mode=ModeKind.LEADING, value=1.0)
    assert ls.prop_value == 100
    assert ls.prop_mode == ModeKind.LEADING

    ls = LineSpacing(mode=ModeKind.FIXED, value=1.0)
    assert ls.prop_value == 100
    assert ls.prop_mode == ModeKind.FIXED


def test_default() -> None:
    idt = LineSpacing().default
    assert idt.prop_value == 100
    assert idt.prop_mode == ModeKind.SINGLE


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
        style_ls = LineSpacing(mode=ModeKind.SINGLE)
        Write.append_para(cursor=cursor, text=para_text, styles=(style_ls,))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        ls = pp.ParaLineSpacing
        struct = style_ls._get_style("line_spacing")[0]
        assert struct == ls
        cursor.gotoEnd(False)

        style_ls = LineSpacing(mode=ModeKind.LINE_1_15)
        Write.append_para(cursor=cursor, text=para_text, styles=(style_ls,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        struct = style_ls._get_style("line_spacing")[0]
        assert struct == ls
        cursor.gotoEnd(False)

        style_ls = LineSpacing(mode=ModeKind.LINE_1_5)
        Write.append_para(cursor=cursor, text=para_text, styles=(style_ls,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        struct = style_ls._get_style("line_spacing")[0]
        assert struct == ls
        cursor.gotoEnd(False)

        style_ls = LineSpacing(mode=ModeKind.DOUBLE)
        Write.append_para(cursor=cursor, text=para_text, styles=(style_ls,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        struct = style_ls._get_style("line_spacing")[0]
        assert struct == ls
        cursor.gotoEnd(False)

        style_ls = LineSpacing(mode=ModeKind.PROPORTIONAL, value=96)
        Write.append_para(cursor=cursor, text=para_text, styles=(style_ls,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        struct = style_ls._get_style("line_spacing")[0]
        assert struct == ls
        cursor.gotoEnd(False)

        style_ls = LineSpacing(mode=ModeKind.AT_LEAST, value=1.0)
        Write.append_para(cursor=cursor, text=para_text, styles=(style_ls,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Height in range(99, 103)
        assert ls.Mode == ModeKind.AT_LEAST.get_mode()
        cursor.gotoEnd(False)

        style_ls = LineSpacing(mode=ModeKind.LEADING, value=1.0)
        Write.append_para(cursor=cursor, text=para_text, styles=(style_ls,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Height in range(99, 103)
        assert ls.Mode == ModeKind.LEADING.get_mode()
        cursor.gotoEnd(False)

        style_ls = LineSpacing(mode=ModeKind.FIXED, value=5.0)
        Write.append_para(cursor=cursor, text=para_text, styles=(style_ls,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        assert ls.Height in range(498, 503)
        assert ls.Mode == ModeKind.FIXED.get_mode()
        cursor.gotoEnd(False)

        style_ls = LineSpacing(mode=ModeKind.LINE_1_15, active_ln_spacing=True)
        Write.append_para(cursor=cursor, text=para_text, styles=(style_ls,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        ls = pp.ParaLineSpacing
        struct = style_ls._get_style("line_spacing")[0]
        assert struct == ls
        assert pp.ParaRegisterModeActive
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
