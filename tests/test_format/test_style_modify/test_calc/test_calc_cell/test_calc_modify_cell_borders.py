from __future__ import annotations
import pytest


if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format import Styler
from ooodev.format.calc.modify.cell.borders import (
    Borders,
    Shadow,
    Side,
    BorderLineKind,
    ShadowLocation,
    Padding,
    StyleCellKind,
    LineSize,
)
from ooodev.utils.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.office.calc import Calc
from ooodev.utils.color import StandardColor
from ooodev.units.unit_mm100 import UnitMM100


def test_write(loader) -> None:
    # delay = 0 if Lo.bridge_connector.headless else 3_000
    delay = 0

    doc = Calc.create_doc()
    if not Lo.bridge_connector.headless:
        GUI.set_visible()
        Lo.delay(500)
        GUI.zoom(GUI.ZoomEnum.ZOOM_150_PERCENT)
    try:
        sheet = Calc.get_active_sheet()

        cell_obj = Calc.get_cell_obj("A1")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)

        shadow = Shadow()

        left100 = UnitMM100.from_pt(4.5)
        right100 = UnitMM100.from_pt(2.5)

        style = Borders(
            left=Side(line=BorderLineKind.DASHED, color=StandardColor.BLUE, width=left100),
            right=Side(line=BorderLineKind.DOUBLE, width=right100),
            shadow=shadow,
            style_name=StyleCellKind.DEFAULT,
        )
        Styler.apply(doc, style)

        f_style = Borders.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_left.prop_width.get_value_mm100() in range(
            left100.value - 2, left100.value + 2
        )  # +- 2
        assert f_style.prop_inner.prop_left.prop_color == StandardColor.BLUE
        assert f_style.prop_inner.prop_right.prop_color == StandardColor.BLACK

        # ===================================================================

        all100 = UnitMM100.from_pt(1.5)
        padding_amt = UnitMM100.from_mm(3)

        d_up = Side(line=BorderLineKind.THINTHICK_SMALLGAP, color=StandardColor.RED_DARK2, width=LineSize.THIN)
        d_dn = Side(line=BorderLineKind.FINE_DASHED, color=StandardColor.GREEN_LIGHT2, width=LineSize.HAIRLINE)

        style = Borders(
            border_side=Side(line=BorderLineKind.SOLID, color=StandardColor.YELLOW_DARK2, width=all100),
            diagonal_up=d_up,
            diagonal_down=d_dn,
            padding=Padding(all=padding_amt),
            style_name=StyleCellKind.DEFAULT,
        )
        Styler.apply(doc, style)

        f_style = Borders.from_style(
            doc=doc, style_name=style.prop_style_name, style_family=style.prop_style_family_name
        )
        assert f_style.prop_inner.prop_left.prop_width.get_value_mm100() in range(
            all100.value - 2, all100.value + 2
        )  # +- 2
        assert f_style.prop_inner.prop_left.prop_color == StandardColor.YELLOW_DARK2
        assert f_style.prop_inner.prop_right.prop_width.get_value_mm100() in range(
            all100.value - 2, all100.value + 2
        )  # +- 2
        assert f_style.prop_inner.prop_right.prop_color == StandardColor.YELLOW_DARK2
        assert f_style.prop_inner.prop_top.prop_width.get_value_mm100() in range(
            all100.value - 2, all100.value + 2
        )  # +- 2
        assert f_style.prop_inner.prop_top.prop_color == StandardColor.YELLOW_DARK2
        assert f_style.prop_inner.prop_bottom.prop_width.get_value_mm100() in range(
            all100.value - 2, all100.value + 2
        )  # +- 2
        assert f_style.prop_inner.prop_bottom.prop_color == StandardColor.YELLOW_DARK2

        assert f_style.prop_inner.prop_diagonal_up.prop_line == BorderLineKind.THINTHICK_SMALLGAP
        assert f_style.prop_inner.prop_diagonal_up.prop_color == StandardColor.RED_DARK2
        assert f_style.prop_inner.prop_diagonal_dn.prop_line == BorderLineKind.FINE_DASHED
        assert f_style.prop_inner.prop_diagonal_dn.prop_color == StandardColor.GREEN_LIGHT2

        assert f_style.prop_inner.prop_padding.prop_left.get_value_mm100() in range(
            padding_amt.value - 2, padding_amt.value + 3
        )  # +- 2
        assert f_style.prop_inner.prop_padding.prop_right.get_value_mm100() in range(
            padding_amt.value - 2, padding_amt.value + 3
        )  # +- 2
        assert f_style.prop_inner.prop_padding.prop_top.get_value_mm100() in range(
            padding_amt.value - 2, padding_amt.value + 3
        )  # +- 2
        assert f_style.prop_inner.prop_padding.prop_bottom.get_value_mm100() in range(
            padding_amt.value - 2, padding_amt.value + 3
        )  # +- 2
        # ===================================================================

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
