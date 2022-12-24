from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.font import (
    Font,
    StrikeOutKind,
    LineKind,
    WeightKind,
    CharSetKnid,
    FamilyKind,
    SlantKind,
)
from ooodev.styles import CommonColor
from ooodev.utils.info import Info
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_font(loader) -> None:
    ft = Font(
        name=Info.get_font_general_name(),
        height=22.0,
        charset=CharSetKnid.SYSTEM,
        family=FamilyKind.MODERN,
        b=True,
        i=True,
        u=True,
        color=CommonColor.BLUE,
        strike=StrikeOutKind.BOLD,
        underine_color=CommonColor.AQUA,
        overline=LineKind.BOLDDASHDOT,
        overline_color=CommonColor.BEIGE,
        super_script=True,
        rotation=90.0,
    )
    assert ft.name == Info.get_font_general_name()
    assert ft.charset == CharSetKnid.SYSTEM
    assert ft.family == FamilyKind.MODERN
    assert ft.b
    assert ft.i
    assert ft.u
    assert ft.weight == WeightKind.BOLD
    assert ft.slant == SlantKind.ITALIC
    assert ft.underline == LineKind.SINGLE
    assert ft.color == CommonColor.BLUE
    assert ft.strike == StrikeOutKind.BOLD
    assert ft.underine_color == CommonColor.AQUA
    assert ft.super_script
    assert ft.height == 22.0
    assert ft.rotation == 90.0
    assert ft.overline == LineKind.BOLDDASHDOT
    assert ft.overline_color == CommonColor.BEIGE

    ft = Font(weight=WeightKind.BOLD, underine=LineKind.BOLDDASH, slant=SlantKind.OBLIQUE, sub_script=True)
    assert ft.weight == WeightKind.BOLD
    assert ft.underline == LineKind.BOLDDASH
    assert ft.slant == SlantKind.OBLIQUE
    assert ft.sub_script


def test_font_cursor(loader, test_headless) -> None:
    delay = 0  # 5_000
    from ooodev.office.write import Write
    from ooodev.styles import Style
    from functools import partial

    doc = Write.create_doc()
    if not test_headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        # red: 16711680
        ft = Font(height=30.0, b=True, i=True, u=True, color=CommonColor.BLUE, underine_color=CommonColor.GREEN)
        cursor = Write.get_cursor(doc)
        style = partial(Style.apply_style, cursor)
        Write.append(cursor, "hello")
        cursor.goLeft(5, True)
        ft.apply_style(cursor)
        cp = cast("CharacterProperties", cursor)
        assert cp.CharColor == CommonColor.BLUE
        assert cp.CharUnderlineColor == CommonColor.GREEN
        assert cp.CharWeight == WeightKind.BOLD.value
        assert cp.CharPosture == SlantKind.ITALIC
        assert cp.CharUnderline == LineKind.SINGLE.value

        # clear attributes or cursor will continue on with font setting just set above.
        Lo.dispatch_cmd("ResetAttributes")
        Lo.delay(500)

        cursor.gotoEnd(False)
        Style.apply_style(cursor, Font(height=40))
        Write.append(cursor, " world")

        cursor.goLeft(5, False)
        cursor.goRight(1, True)
        Style.apply_style(cursor, Font(super_script=True))
        assert cp.CharEscapement == 14_000

        cursor.gotoEnd(False)
        cursor.goLeft(1, True)
        # use partial function
        style(Font(sub_script=True))
        assert cp.CharEscapement == -14_000
        cursor.gotoEnd(False)

        Write.end_paragraph(cursor)

        Write.append(cursor, "BIG")
        cursor.goLeft(3, True)
        style(Font(height=36, rotation=90.0))
        assert cp.CharRotation == 900
        cursor.gotoEnd(False)

        Write.end_paragraph(cursor)
        Lo.dispatch_cmd("ResetAttributes")
        Lo.delay(500)

        Write.append(cursor, "Overline")
        cursor.goLeft(8, True)
        style(Font(overline=LineKind.BOLDWAVE, height=40, overline_color=CommonColor.CHARTREUSE))
        # no documentation found for Overline
        assert cursor.CharOverlineHasColor
        assert cursor.CharOverlineColor == CommonColor.CHARTREUSE
        assert cursor.CharOverline == LineKind.BOLDWAVE
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Lo.dispatch_cmd("ResetAttributes")
        Lo.delay(500)
        Write.append(cursor, "Highlite")
        cursor.goLeft(8, True)
        style(Font(bg_color=CommonColor.LIGHT_YELLOW))
        assert cp.CharBackColor == CommonColor.LIGHT_YELLOW

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_font(loader, test_headless) -> None:
    delay = 5_000
    from ooodev.office.calc import Calc
    from ooodev.styles import Style

    doc = Calc.create_doc()
    try:
        sheet = Calc.get_sheet(doc)
        if not test_headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_200_PERCENT)
        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        cp = cast("CharacterProperties", cell)
        Style.apply_style(cell, Font(b=True, u=True, color=CommonColor.ROYAL_BLUE))
        assert cp.CharWeight == WeightKind.BOLD.value
        assert cp.CharUnderline == LineKind.SINGLE.value
        assert cp.CharColor == CommonColor.ROYAL_BLUE

        cell_obj = Calc.get_cell_obj("A3")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        cp = cast("CharacterProperties", cell)
        Style.apply_style(
            cell,
            Font(
                i=True,
                overline=LineKind.DOUBLE,
                overline_color=CommonColor.RED,
                color=CommonColor.DARK_GREEN,
                bg_color=CommonColor.LIGHT_GRAY,
            ),
        )
        # no documentation found for Overline
        assert cell.CharOverlineHasColor
        assert cell.CharOverlineColor == CommonColor.RED
        assert cell.CharOverline == LineKind.DOUBLE.value
        assert cp.CharPosture == SlantKind.ITALIC
        assert cp.CharColor == CommonColor.DARK_GREEN
        # CharBackColor not supported for cell
        # assert cp.CharBackColor == CommonColor.LIGHT_GRAY

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
