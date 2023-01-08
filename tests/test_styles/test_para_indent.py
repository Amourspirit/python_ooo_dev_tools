from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.para.indent import Indent
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_indent_props() -> None:
    idt = Indent(before=10.2)
    assert idt.prop_before > 10.1 and idt.prop_before < 10.3
    assert idt._get("ParaLeftMargin") == 1020

    idt = Indent(after=10.2)
    assert idt.prop_after > 10.1 and idt.prop_after < 10.3
    assert idt._get("ParaRightMargin") == 1020

    idt = Indent(first=10.2)
    assert idt.prop_first > 10.1 and idt.prop_first < 10.3
    assert idt._get("ParaFirstLineIndent") == 1020

    idt = Indent(auto=True)
    assert idt.prop_auto
    assert idt._get("ParaIsAutoFirstLineIndent")


def test_indent_default() -> None:
    idt = cast(Indent, Indent.default)
    assert idt.prop_after == 0.0
    assert idt.prop_before == 0.0
    assert idt.prop_first == 0.0
    assert idt.prop_auto == False


def test_indent_auto() -> None:
    idt = Indent().auto
    assert idt.prop_auto


def test_alignemnt_write(loader, para_text) -> None:
    delay = 0 if Lo.bridge_connector.headless else 3_000

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        cursor = Write.get_cursor(doc)
        p_len = len(para_text)
        amt = 3.5
        Write.append_para(cursor=cursor, text=para_text, styles=(Indent(before=amt),))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        assert pp.ParaLeftMargin in [round(amt * 100) - 2 + i for i in range(5)]  # plus or minus 2
        cursor.gotoEnd(False)

        amt = 2.0
        Write.append_para(cursor=cursor, text=para_text, styles=(Indent(after=amt),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaRightMargin in [round(amt * 100) - 2 + i for i in range(5)]  # plus or minus 2
        cursor.gotoEnd(False)

        amt = 6.0
        Write.append_para(cursor=cursor, text=para_text, styles=(Indent(first=amt),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaFirstLineIndent in [round(amt * 100) - 2 + i for i in range(5)]  # plus or minus 2
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(Indent().auto,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaIsAutoFirstLineIndent
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
