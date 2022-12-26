from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.tbl_borders import CellBorder, BorderTable, Shadow, Side, LineStyleKind, ShadowLocation
from ooodev.styles import CommonColor, Style
from ooodev.utils.info import Info
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo

if TYPE_CHECKING:
    from com.sun.star.table import CellProperties  # service


def test_calc_border(loader, test_headless) -> None:
    delay = 0 if test_headless else 5_000
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
        cb = CellBorder(border_side=Side())
        Style.apply_style(cell, cb)
        cp = cast("CellProperties", cell)
        assert cp.TableBorder2.IsLeftLineValid
        assert cp.TableBorder2.LeftLine.LineWidth == 26
        assert cp.TableBorder2.IsRightLineValid
        assert cp.TableBorder2.RightLine.LineWidth == 26
        assert cp.TableBorder2.IsTopLineValid
        assert cp.TableBorder2.TopLine.LineWidth == 26
        assert cp.TableBorder2.IsBottomLineValid
        assert cp.TableBorder2.BottomLine.LineWidth == 26

        shadow = Shadow()

        cell_obj = Calc.get_cell_obj("c1")
        cell = Calc.get_cell(sheet, cell_obj)
        cb = CellBorder(
            left=Side(style=LineStyleKind.DASHED, color=CommonColor.BLUE, width=4.5),
            right=Side(style=LineStyleKind.DOUBLE, width=2.5),
            shadow=shadow,
        )
        Style.apply_style(cell, cb)
        cp = cast("CellProperties", cell)
        assert cp.TableBorder2.IsLeftLineValid
        assert cp.TableBorder2.LeftLine.LineWidth == 450
        assert cp.TableBorder2.LeftLine.LineStyle == LineStyleKind.DASHED.value

        assert cp.TableBorder2.IsRightLineValid
        assert cp.TableBorder2.RightLine.LineWidth == 250
        assert cp.TableBorder2.RightLine.LineStyle == LineStyleKind.DOUBLE.value

        assert cp.ShadowFormat == shadow.get_shadow_format()

        shadow = Shadow(location=ShadowLocation.BOTTOM_LEFT)
        assert cp.ShadowFormat != shadow.get_shadow_format()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
