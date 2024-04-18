from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.writer.direct.para.outline_list import Outline, LevelKind
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.write import Write


if TYPE_CHECKING:
    from com.sun.star.style import ParagraphProperties  # service


def test_props() -> None:
    ol = Outline(level=LevelKind.TEXT_BODY)
    assert ol._get("OutlineLevel") == LevelKind.TEXT_BODY
    assert ol.prop_level == LevelKind.TEXT_BODY


def test_default() -> None:
    ol = Outline().default
    assert ol._get("OutlineLevel") == LevelKind.TEXT_BODY
    assert ol.prop_level == LevelKind.TEXT_BODY


def test_text_body() -> None:
    ol = Outline().text_body
    assert ol.prop_level == LevelKind.TEXT_BODY


def test_level_01() -> None:
    ol = Outline().level_01
    assert ol.prop_level == LevelKind.LEVEL_01


def test_level_02() -> None:
    ol = Outline().level_02
    assert ol.prop_level == LevelKind.LEVEL_02


def test_level_03() -> None:
    ol = Outline().level_03
    assert ol.prop_level == LevelKind.LEVEL_03


def test_level_04() -> None:
    ol = Outline().level_04
    assert ol.prop_level == LevelKind.LEVEL_04


def test_level_05() -> None:
    ol = Outline().level_05
    assert ol.prop_level == LevelKind.LEVEL_05


def test_level_06() -> None:
    ol = Outline().level_06
    assert ol.prop_level == LevelKind.LEVEL_06


def test_level_07() -> None:
    ol = Outline().level_07
    assert ol.prop_level == LevelKind.LEVEL_07


def test_level_08() -> None:
    ol = Outline().level_08
    assert ol.prop_level == LevelKind.LEVEL_08


def test_level_09() -> None:
    ol = Outline().level_09
    assert ol.prop_level == LevelKind.LEVEL_09


def test_level_10() -> None:
    ol = Outline().level_10
    assert ol.prop_level == LevelKind.LEVEL_10


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
        Write.append_para(cursor=cursor, text=para_text, styles=(Outline(level=LevelKind.LEVEL_01),))

        cursor.goLeft(1, False)
        cursor.gotoStart(True)

        pp = cast("ParagraphProperties", cursor)
        pp.OutlineLevel == int(LevelKind.LEVEL_01)
        cursor.gotoEnd(False)

        Write.append_para(cursor=cursor, text=para_text, styles=(Outline().level_03,))
        cursor.goLeft(p_len + 1, False)
        cursor.goRight(p_len, True)
        pp.OutlineLevel == int(LevelKind.LEVEL_03)
        cursor.gotoEnd(False)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
