from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.styles.tbl_borders import (
    Border,
    Shadow,
    Side,
    BorderLineStyleEnum,
    ShadowLocation,
    BorderPadding,
    DEFAULT_BORDER,
)
from ooodev.styles import CommonColor, Style
from ooodev.utils.info import Info
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.styles.style_const import POINT_RATIO

if TYPE_CHECKING:
    from com.sun.star.table import CellProperties  # service
    from com.sun.star.style import ParagraphProperties  # service


def test_calc_border(loader, test_headless) -> None:
    delay = 0 if test_headless else 5_000
    from ooodev.office.calc import Calc

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
        cb = Border(border_side=Side())

        Style.apply_style(cell, cb)
        cp = cast("CellProperties", cell)
        # line width may not be applied exact by LibreOffice.
        assert cp.TableBorder2.IsLeftLineValid
        assert cp.TableBorder2.LeftLine.LineWidth in [
            round(0.75 * POINT_RATIO) - 1 + i for i in range(3)
        ]  # plus or minus 1

        assert cp.TableBorder2.IsRightLineValid
        assert cp.TableBorder2.RightLine.LineWidth == cp.TableBorder2.LeftLine.LineWidth

        assert cp.TableBorder2.IsTopLineValid
        assert cp.TableBorder2.TopLine.LineWidth == cp.TableBorder2.LeftLine.LineWidth

        assert cp.TableBorder2.IsBottomLineValid
        assert cp.TableBorder2.BottomLine.LineWidth == cp.TableBorder2.LeftLine.LineWidth

        shadow = Shadow()

        cell_obj = Calc.get_cell_obj("c1")
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Border(
            left=Side(style=BorderLineStyleEnum.DASHED, color=CommonColor.BLUE, width=4.5),
            right=Side(style=BorderLineStyleEnum.DOUBLE, width=2.5),
            shadow=shadow,
        )
        Style.apply_style(cell, cb)
        cp = cast("CellProperties", cell)
        assert cp.TableBorder2.IsLeftLineValid
        assert cp.TableBorder2.LeftLine.LineWidth in [
            round(4.5 * POINT_RATIO) - 1 + i for i in range(3)
        ]  # plus or minus 1
        assert cp.TableBorder2.LeftLine.LineStyle == BorderLineStyleEnum.DASHED.value

        assert cp.TableBorder2.IsRightLineValid
        assert cp.TableBorder2.RightLine.LineWidth in [
            round(2.5 * POINT_RATIO) - 1 + i for i in range(3)
        ]  # plus or minus 1
        assert cp.TableBorder2.RightLine.LineStyle == BorderLineStyleEnum.DOUBLE.value

        assert cp.ShadowFormat == shadow.get_shadow_format()

        shadow = Shadow(location=ShadowLocation.BOTTOM_LEFT)
        assert cp.ShadowFormat != shadow.get_shadow_format()

        cell_obj = Calc.get_cell_obj("e1")
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Border(diagonal_up=Side(style=BorderLineStyleEnum.DOUBLE_THIN, color=CommonColor.RED))
        Style.apply_style(cell, cb)
        cp = cast("CellProperties", cell)
        assert cp.DiagonalBLTR2.Color == CommonColor.RED
        assert cp.DiagonalBLTR2.LineStyle == BorderLineStyleEnum.DOUBLE_THIN.value

        cell_obj = Calc.get_cell_obj("g1")
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Border(diagonal_down=Side(style=BorderLineStyleEnum.DOTTED, color=CommonColor.BROWN))
        Style.apply_style(cell, cb)
        cp = cast("CellProperties", cell)
        assert cp.DiagonalTLBR2.Color == CommonColor.BROWN
        assert cp.DiagonalTLBR2.LineStyle == BorderLineStyleEnum.DOTTED.value

        side = Side(color=CommonColor.ORANGE_RED)

        cell_obj = Calc.get_cell_obj("A3")
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Border(diagonal_down=side, diagonal_up=side)
        Style.apply_style(cell, cb)
        cp = cast("CellProperties", cell)
        assert cp.DiagonalTLBR2.Color == CommonColor.ORANGE_RED
        assert cp.DiagonalTLBR2.LineStyle == BorderLineStyleEnum.SOLID.value
        assert cp.DiagonalBLTR2.Color == CommonColor.ORANGE_RED
        assert cp.DiagonalBLTR2.LineStyle == BorderLineStyleEnum.SOLID.value

        cell_obj = Calc.get_cell_obj("C3")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Border(diagonal_down=side, diagonal_up=side)
        Style.apply_style(cell, Border(padding=BorderPadding(padding_all=0.7)))
        cp = cast("ParagraphProperties", cell)
        # padding may not apply exact
        assert cp.ParaLeftMargin >= 69 and cp.ParaLeftMargin <= 72
        assert cp.ParaRightMargin == cp.ParaLeftMargin
        assert cp.ParaTopMargin == cp.ParaLeftMargin
        assert cp.ParaBottomMargin == cp.ParaLeftMargin

        cell_obj = Calc.get_cell_obj("E3")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Border(diagonal_down=side, diagonal_up=side)
        Style.apply_style(cell, Border(padding=BorderPadding(left=1.2, right=1.2, top=0.5, bottom=0.5)))
        cp = cast("ParagraphProperties", cell)
        # padding may not apply exact
        assert cp.ParaLeftMargin >= 118 and cp.ParaLeftMargin <= 122
        assert cp.ParaRightMargin == cp.ParaLeftMargin

        assert cp.ParaTopMargin >= 48 and cp.ParaTopMargin <= 52
        assert cp.ParaBottomMargin == cp.ParaTopMargin

        cell_obj = Calc.get_cell_obj("G3")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Border(diagonal_down=side, diagonal_up=side)
        Style.apply_style(cell, DEFAULT_BORDER)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_border_range(loader, test_headless) -> None:
    delay = 0 if test_headless else 5_000
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    try:
        sheet = Calc.get_sheet(doc)
        if not test_headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_200_PERCENT)

        rng_obj = Calc.get_range_obj("B2:G8")
        cr = Calc.get_cell_range(sheet, rng_obj)

        # for some unknown reason LibreOffice is overriding style of horizontal Side. to match outter border.
        cb = Border(
            border_side=Side(style=BorderLineStyleEnum.SOLID, color=CommonColor.BLUE),
            vertical=Side(color=CommonColor.RED, style=BorderLineStyleEnum.DASHED),
            horizontal=Side(color=CommonColor.GREEN, width=1.4, style=BorderLineStyleEnum.DOUBLE),
        )
        Style.apply_style(cr, cb)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
