from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno

# simpler test then test_char_font because it is testing the same font class under the hood.
from ooodev.format.calc.direct.font import (
    Font,
    FontUnderlineEnum,
    FontWeightEnum,
    FontSlant,
)
from ooodev.format import CommonColor
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

if TYPE_CHECKING:
    from com.sun.star.style import CharacterProperties  # service


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
