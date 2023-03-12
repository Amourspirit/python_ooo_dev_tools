from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.direct.para.indent_space.spacing import Spacing
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_spacing_props() -> None:
    idt = Spacing(above=10.2)
    assert idt.prop_above.value > 10.1 and idt.prop_above.value < 10.3
    assert idt._get("ParaTopMargin") == 1020

    idt = Spacing(below=10.2)
    assert idt.prop_below.value > 10.1 and idt.prop_below.value < 10.3
    assert idt._get("ParaBottomMargin") == 1020

    idt = Spacing(style_no_space=True)
    assert idt.prop_style_no_space
    assert idt._get("ParaContextMargin")


def test_spacing_default() -> None:
    idt = Spacing().default
    assert idt.prop_above.value == 0.0
    assert idt.prop_below.value == 0.0
    assert idt.prop_style_no_space == False


def test_spacing_style_no_space() -> None:
    idt = Spacing().style_no_space
    assert idt.prop_style_no_space


def test_spacing_write(loader, para_text) -> None:
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
        amt = 3.5
        Write.append_para(cursor=cursor, text=para_text, styles=(Spacing(above=amt),))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        assert pp.ParaTopMargin in [round(amt * 100) - 2 + i for i in range(5)]  # plus or minus 2
        cursor.gotoEnd(False)

        amt = 2.0
        Write.append_para(cursor=cursor, text=para_text, styles=(Spacing(below=amt),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaBottomMargin in [round(amt * 100) - 2 + i for i in range(5)]  # plus or minus 2
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(Spacing().style_no_space,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaContextMargin
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
