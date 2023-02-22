from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.para.text_flow import InnerHyphenation
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_props() -> None:
    hy = InnerHyphenation(auto=True)
    assert hy.prop_auto
    assert hy._get("ParaIsHyphenation")

    hy = InnerHyphenation(no_caps=True)
    assert hy.prop_no_caps
    assert hy._get("ParaHyphenationNoCaps")

    hy = InnerHyphenation(start_chars=10)
    assert hy.prop_start_chars == 10
    assert hy._get("ParaHyphenationMaxLeadingChars") == 10

    hy = InnerHyphenation(end_chars=12)
    assert hy.prop_end_chars == 12
    assert hy._get("ParaHyphenationMaxTrailingChars") == 12

    hy = InnerHyphenation(max=100)
    assert hy.prop_max == 100
    assert hy._get("ParaHyphenationMaxHyphens") == 100


def test_default() -> None:
    hy = cast(InnerHyphenation, InnerHyphenation.default)
    assert hy.prop_auto == False
    assert hy.prop_no_caps == False
    assert hy.prop_start_chars == 2
    assert hy.prop_end_chars == 2
    assert hy.prop_max == 0


def test_auto() -> None:
    hy = InnerHyphenation().auto
    assert hy.prop_auto


def test_no_caps() -> None:
    hy = InnerHyphenation().no_caps
    assert hy.prop_no_caps


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

        Write.append_para(cursor=cursor, text=para_text, styles=(InnerHyphenation(auto=True),))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        assert pp.ParaIsHyphenation
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(InnerHyphenation(auto=True, no_caps=True),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaIsHyphenation
        assert pp.ParaHyphenationNoCaps
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(InnerHyphenation(auto=True, start_chars=3),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaIsHyphenation
        assert pp.ParaHyphenationMaxLeadingChars == 3
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(InnerHyphenation(end_chars=4).auto,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaIsHyphenation
        assert pp.ParaHyphenationMaxTrailingChars == 4
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(InnerHyphenation(auto=True, max=2),))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaIsHyphenation
        assert pp.ParaHyphenationMaxHyphens == 2
        cursor.gotoEnd(False)

        # apply style directly to cursor
        hy = InnerHyphenation(auto=True, no_caps=True, start_chars=5, end_chars=4, max=8)
        hy.apply(cursor)
        Write.append_para(cursor=cursor, text=para_text)
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        assert pp.ParaIsHyphenation
        assert pp.ParaHyphenationNoCaps
        assert pp.ParaHyphenationMaxLeadingChars == 5
        assert pp.ParaHyphenationMaxTrailingChars == 4
        assert pp.ParaHyphenationMaxHyphens == 8
        cursor.gotoEnd(False)

        # restore cursor
        InnerHyphenation.default.apply(cursor)
        assert pp.ParaIsHyphenation == False
        assert pp.ParaHyphenationNoCaps == False
        assert pp.ParaHyphenationMaxLeadingChars == 2
        assert pp.ParaHyphenationMaxTrailingChars == 2
        assert pp.ParaHyphenationMaxHyphens == 0

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
