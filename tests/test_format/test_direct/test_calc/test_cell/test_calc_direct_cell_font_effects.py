from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])


# simpler test then test_char_font because it is testing the same font class under the hood.
from ooodev.format.calc.direct.cell.font import (
    FontOnly,
    FontUnderlineEnum,
    FontStrikeoutEnum,
    FontEffects,
    FontLine,
)
from ooodev.format import CommonColor
from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.units.unit_pt import UnitPT

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


def test_calc_font_effects(loader) -> None:
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
        font_size = UnitPT(12.0)
        style_fo = FontOnly(name="Liberation Sans", size=font_size, font_style="Bold")
        style_fe = FontEffects(
            color=CommonColor.BLUE,
            underline=FontLine(line=FontUnderlineEnum.SINGLE, color=CommonColor.RED),
            strike=FontStrikeoutEnum.SINGLE,
        )
        Styler.apply(cell, style_fo, style_fe)
        assert cp.CharStrikeout == FontStrikeoutEnum.SINGLE.value
        assert cp.CharUnderlineColor == CommonColor.RED
        assert cp.CharUnderline == 1
        assert cp.CharColor == CommonColor.BLUE

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
