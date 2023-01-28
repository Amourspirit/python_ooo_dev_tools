from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.direct.char.font import (
    Font,
    FontStrikeoutEnum,
    FontUnderlineEnum,
    FontWeightEnum,
    CharSetEnum,
    FontFamilyEnum,
    FontSlant,
    CharSpacingKind,
)
from ooodev.format import CommonColor
from ooodev.utils.unit_convert import UnitConvert
from ooodev.utils.gui import GUI
from ooodev.utils.info import Info
from ooodev.utils.lo import Lo

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_font_chain() -> None:
    ft = Font().bold
    assert ft.prop_weight == FontWeightEnum.BOLD
    ft = ft.italic
    assert ft.prop_weight == FontWeightEnum.BOLD
    assert ft.prop_is_italic
    assert ft.prop_is_bold

    ft = ft.fmt_color(CommonColor.RED)
    assert ft.prop_weight == FontWeightEnum.BOLD
    assert ft.prop_is_italic
    assert ft.prop_is_bold
    assert ft.prop_color == CommonColor.RED

    ft = ft.fmt_bg_color(CommonColor.BLUE)
    assert ft.prop_weight == FontWeightEnum.BOLD
    assert ft.prop_is_italic
    assert ft.prop_is_bold
    assert ft.prop_color == CommonColor.RED
    assert ft.prop_bg_color == CommonColor.BLUE

    ft = ft.underline
    assert ft.prop_weight == FontWeightEnum.BOLD
    assert ft.prop_is_italic
    assert ft.prop_is_bold
    assert ft.prop_color == CommonColor.RED
    assert ft.prop_bg_color == CommonColor.BLUE
    assert ft.prop_is_underline

    ft = ft.strike
    assert ft.prop_weight == FontWeightEnum.BOLD
    assert ft.prop_is_italic
    assert ft.prop_is_bold
    assert ft.prop_color == CommonColor.RED
    assert ft.prop_bg_color == CommonColor.BLUE
    assert ft.prop_is_underline
    assert ft.prop_strike == FontStrikeoutEnum.SINGLE

    ft = ft.shadowed
    assert ft.prop_weight == FontWeightEnum.BOLD
    assert ft.prop_is_italic
    assert ft.prop_is_bold
    assert ft.prop_color == CommonColor.RED
    assert ft.prop_bg_color == CommonColor.BLUE
    assert ft.prop_is_underline
    assert ft.prop_strike == FontStrikeoutEnum.SINGLE
    assert ft.prop_shadowed


def test_font(loader) -> None:
    ft = Font(
        name=Info.get_font_general_name(),
        size=22.0,
        charset=CharSetEnum.SYSTEM,
        family=FontFamilyEnum.MODERN,
        b=True,
        i=True,
        u=True,
        color=CommonColor.BLUE,
        strike=FontStrikeoutEnum.BOLD,
        underine_color=CommonColor.AQUA,
        overline=FontUnderlineEnum.BOLDDASHDOT,
        overline_color=CommonColor.BEIGE,
        superscript=True,
        rotation=90.0,
        spacing=CharSpacingKind.TIGHT,
        shadowed=True,
    )
    assert ft.prop_name == Info.get_font_general_name()
    assert ft.prop_charset == CharSetEnum.SYSTEM
    assert ft.prop_family == FontFamilyEnum.MODERN
    assert ft.prop_is_bold
    assert ft.prop_is_italic
    assert ft.prop_is_underline
    assert ft.prop_weight == FontWeightEnum.BOLD
    assert ft.prop_slant == FontSlant.ITALIC
    assert ft.prop_underline == FontUnderlineEnum.SINGLE
    assert ft.prop_color == CommonColor.BLUE
    assert ft.prop_strike == FontStrikeoutEnum.BOLD
    assert ft.prop_underine_color == CommonColor.AQUA
    assert ft.prop_superscript
    assert ft.prop_size == 22.0
    assert ft.prop_rotation == 90.0
    assert ft.prop_overline == FontUnderlineEnum.BOLDDASHDOT
    assert ft.prop_overline_color == CommonColor.BEIGE
    assert ft.prop_spacing == pytest.approx(CharSpacingKind.TIGHT.value, rel=1e-2)
    assert ft.prop_shadowed

    ft = Font(
        weight=FontWeightEnum.BOLD,
        underine=FontUnderlineEnum.BOLDDASH,
        slant=FontSlant.OBLIQUE,
        subscript=True,
        spacing=2.0,
    )
    assert ft.prop_weight == FontWeightEnum.BOLD
    assert ft.prop_is_underline
    assert ft.prop_underline == FontUnderlineEnum.BOLDDASH
    assert ft.prop_slant == FontSlant.OBLIQUE
    assert ft.prop_subscript
    assert ft.prop_spacing == pytest.approx(2.0, rel=1e-2)


def test_font_cursor(loader) -> None:
    delay = 0  # 0 if Lo.bridge_connector.headless else 5_000
    from ooodev.office.write import Write
    from ooodev.format import Styler
    from functools import partial

    doc = Write.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        ft = Font(size=30.0, b=True, i=True, u=True, color=CommonColor.BLUE, underine_color=CommonColor.GREEN)
        cursor = Write.get_cursor(doc)
        style = partial(Styler.apply, cursor)
        Write.append(cursor, "hello")
        cursor.goLeft(5, True)
        ft.apply(cursor)
        cp = cast("CharacterProperties", cursor)
        assert cp.CharColor == CommonColor.BLUE
        assert cp.CharUnderlineColor == CommonColor.GREEN
        assert cp.CharWeight == FontWeightEnum.BOLD.value
        assert cp.CharPosture == FontSlant.ITALIC
        assert cp.CharUnderline == FontUnderlineEnum.SINGLE.value
        assert cp.CharHeight == pytest.approx(30.0, rel=1e-2)

        # clear attributes or cursor will continue on with font setting just set above.
        Lo.dispatch_cmd("ResetAttributes")
        Lo.delay(500)

        cursor.gotoEnd(False)
        Styler.apply(cursor, Font(size=40))
        Write.append(cursor, " world")

        cursor.goLeft(5, False)
        cursor.goRight(1, True)
        Styler.apply(cursor, Font(superscript=True))
        assert cp.CharEscapement == 14_000

        cursor.gotoEnd(False)
        cursor.goLeft(1, True)
        # use partial function
        style(Font(subscript=True))
        assert cp.CharEscapement == -14_000
        cursor.gotoEnd(False)

        Write.end_paragraph(cursor)

        Write.append(cursor, "BIG")
        cursor.goLeft(3, True)
        style(Font(size=36, rotation=90.0))
        assert cp.CharRotation == 900
        cursor.gotoEnd(False)

        Write.end_paragraph(cursor)
        Lo.dispatch_cmd("ResetAttributes")
        Lo.delay(500)

        Write.append(cursor, "Overline")
        cursor.goLeft(8, True)
        style(Font(overline=FontUnderlineEnum.BOLDWAVE, size=40, overline_color=CommonColor.CHARTREUSE))
        # no documentation found for Overline
        assert cursor.CharOverlineHasColor
        assert cursor.CharOverlineColor == CommonColor.CHARTREUSE
        assert cursor.CharOverline == FontUnderlineEnum.BOLDWAVE
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Lo.dispatch_cmd("ResetAttributes")
        Lo.delay(500)
        Write.append(cursor, "Highlite")
        cursor.goLeft(8, True)
        style(Font(bg_color=CommonColor.LIGHT_YELLOW))
        assert cp.CharBackColor == CommonColor.LIGHT_YELLOW
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Lo.dispatch_cmd("ResetAttributes")
        Lo.delay(500)
        Write.append(cursor, "Very Tight")
        cursor.goLeft(10, True)
        style(Font(spacing=CharSpacingKind.VERY_TIGHT, size=30))
        assert cp.CharKerning == UnitConvert.convert_pt_mm100(CharSpacingKind.VERY_TIGHT.value)
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Write.append(cursor, "Tight")
        cursor.goLeft(5, True)
        style(Font(spacing=CharSpacingKind.TIGHT, size=30))
        assert cp.CharKerning == UnitConvert.convert_pt_mm100(CharSpacingKind.TIGHT.value)
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Write.append(cursor, "Normal")
        cursor.goLeft(6, True)
        style(Font(spacing=CharSpacingKind.NORMAL, size=30))
        assert cp.CharKerning == 0
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Write.append(cursor, "Loose")
        cursor.goLeft(5, True)
        style(Font(spacing=CharSpacingKind.LOOSE, size=30))
        assert cp.CharKerning == UnitConvert.convert_pt_mm100(CharSpacingKind.LOOSE.value)
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Write.append(cursor, "Very Loose")
        cursor.goLeft(10, True)
        style(Font(spacing=CharSpacingKind.VERY_LOOSE, size=30))
        assert cp.CharKerning == UnitConvert.convert_pt_mm100(CharSpacingKind.VERY_LOOSE.value)
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Write.append(cursor, "Custom Spacing 6 pt")
        cursor.goLeft(19, True)
        style(Font(spacing=19.0, size=14))
        assert cp.CharKerning == UnitConvert.convert_pt_mm100(19.0)
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Lo.dispatch_cmd("ResetAttributes")
        Lo.delay(500)

        Write.append(cursor, "Shadowed")
        cursor.goLeft(8, True)
        style(Font(size=40, shadowed=True))
        assert cp.CharShadowed
        cursor.gotoEnd(False)
        Write.end_paragraph(cursor)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_font(loader) -> None:
    # delay = 0 if not Lo.bridge_connector.headless else 5_000
    delay = 0
    from ooodev.office.calc import Calc
    from ooodev.format import Styler

    doc = Calc.create_doc()
    try:
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_200_PERCENT)
        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        cp = cast("CharacterProperties", cell)
        Styler.apply(cell, Font(b=True, u=True, color=CommonColor.ROYAL_BLUE))
        assert cp.CharWeight == FontWeightEnum.BOLD.value
        assert cp.CharUnderline == FontUnderlineEnum.SINGLE.value
        assert cp.CharColor == CommonColor.ROYAL_BLUE

        cell_obj = Calc.get_cell_obj("A3")
        Calc.set_val(value="World", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        cp = cast("CharacterProperties", cell)
        Styler.apply(
            cell,
            Font(
                i=True,
                overline=FontUnderlineEnum.DOUBLE,
                overline_color=CommonColor.RED,
                color=CommonColor.DARK_GREEN,
                bg_color=CommonColor.LIGHT_GRAY,
            ),
        )
        # no documentation found for Overline
        assert cell.CharOverlineHasColor
        assert cell.CharOverlineColor == CommonColor.RED
        assert cell.CharOverline == FontUnderlineEnum.DOUBLE.value
        assert cp.CharPosture == FontSlant.ITALIC
        assert cp.CharColor == CommonColor.DARK_GREEN
        # CharBackColor not supported for cell
        # assert cp.CharBackColor == CommonColor.LIGHT_GRAY

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
