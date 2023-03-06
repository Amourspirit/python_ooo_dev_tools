from __future__ import annotations
import pytest
from typing import TYPE_CHECKING, cast

if __name__ == "__main__":
    pytest.main([__file__])

import uno
from ooodev.format.calc.direct.borders import (
    Borders,
    Shadow,
    Side,
    BorderLineStyleEnum,
    ShadowLocation,
    Padding,
)
from ooodev.format import CommonColor, Styler
from ooodev.utils.gui import GUI
from ooodev.utils.lo import Lo
from ooodev.utils.unit_convert import UnitConvert

if TYPE_CHECKING:
    from com.sun.star.table import CellProperties  # service
    from com.sun.star.style import ParagraphProperties  # service
    from com.sun.star.table import CellRange  # service


def test_calc_border(loader) -> None:
    delay = 0  # 0 if Lo.bridge_connector.headless else 5_000
    from ooodev.office.calc import Calc

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
        cb = Borders(border_side=Side())

        Styler.apply(cell, cb)
        cp = cast("CellProperties", cell)
        # line width may not be applied exact by LibreOffice.
        assert cp.TableBorder2.IsLeftLineValid
        lw_mm100 = UnitConvert.convert_pt_mm100(0.75)
        assert cp.TableBorder2.LeftLine.LineWidth in range(lw_mm100 - 2, lw_mm100 + 3)  # +- 2

        assert cp.TableBorder2.IsRightLineValid
        assert cp.TableBorder2.RightLine.LineWidth == cp.TableBorder2.LeftLine.LineWidth

        assert cp.TableBorder2.IsTopLineValid
        assert cp.TableBorder2.TopLine.LineWidth == cp.TableBorder2.LeftLine.LineWidth

        assert cp.TableBorder2.IsBottomLineValid
        assert cp.TableBorder2.BottomLine.LineWidth == cp.TableBorder2.LeftLine.LineWidth

        shadow = Shadow()

        cell_obj = Calc.get_cell_obj("c1")
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Borders(
            left=Side(line=BorderLineStyleEnum.DASHED, color=CommonColor.BLUE, width=4.5),
            right=Side(line=BorderLineStyleEnum.DOUBLE, width=2.5),
            shadow=shadow,
        )
        Styler.apply(cell, cb)
        cp = cast("CellProperties", cell)
        assert cp.TableBorder2.IsLeftLineValid
        lw_mm100 = UnitConvert.convert_pt_mm100(4.5)
        assert cp.TableBorder2.LeftLine.LineWidth in range(lw_mm100 - 2, lw_mm100 + 3)  # +- 2
        assert cp.TableBorder2.LeftLine.LineStyle == BorderLineStyleEnum.DASHED.value

        assert cp.TableBorder2.IsRightLineValid
        lw_mm100 = UnitConvert.convert_pt_mm100(2.5)
        assert cp.TableBorder2.RightLine.LineWidth in range(lw_mm100 - 2, lw_mm100 + 3)  # +- 2
        assert cp.TableBorder2.RightLine.LineStyle == BorderLineStyleEnum.DOUBLE.value

        assert cp.ShadowFormat == shadow.get_uno_struct()

        shadow = Shadow(location=ShadowLocation.BOTTOM_LEFT)
        assert cp.ShadowFormat != shadow.get_uno_struct()

        cell_obj = Calc.get_cell_obj("e1")
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Borders(diagonal_up=Side(line=BorderLineStyleEnum.DOUBLE_THIN, color=CommonColor.RED))
        Styler.apply(cell, cb)
        cp = cast("CellProperties", cell)
        assert cp.DiagonalBLTR2.Color == CommonColor.RED
        assert cp.DiagonalBLTR2.LineStyle == BorderLineStyleEnum.DOUBLE_THIN.value

        cell_obj = Calc.get_cell_obj("g1")
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Borders(diagonal_down=Side(line=BorderLineStyleEnum.DOTTED, color=CommonColor.BROWN))
        Styler.apply(cell, cb)
        cp = cast("CellProperties", cell)
        assert cp.DiagonalTLBR2.Color == CommonColor.BROWN
        assert cp.DiagonalTLBR2.LineStyle == BorderLineStyleEnum.DOTTED.value

        side = Side(color=CommonColor.ORANGE_RED)

        cell_obj = Calc.get_cell_obj("A3")
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Borders(diagonal_down=side, diagonal_up=side)
        Styler.apply(cell, cb)
        cp = cast("CellProperties", cell)
        assert cp.DiagonalTLBR2.Color == CommonColor.ORANGE_RED
        assert cp.DiagonalTLBR2.LineStyle == BorderLineStyleEnum.SOLID.value
        assert cp.DiagonalBLTR2.Color == CommonColor.ORANGE_RED
        assert cp.DiagonalBLTR2.LineStyle == BorderLineStyleEnum.SOLID.value

        cell_obj = Calc.get_cell_obj("C3")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Borders(diagonal_down=side, diagonal_up=side)
        Styler.apply(cell, Borders(padding=Padding(all=0.7)))
        cp = cast("ParagraphProperties", cell)
        # padding may not apply exact
        assert cp.ParaLeftMargin >= 69 and cp.ParaLeftMargin <= 72
        assert cp.ParaRightMargin == cp.ParaLeftMargin
        assert cp.ParaTopMargin == cp.ParaLeftMargin
        assert cp.ParaBottomMargin == cp.ParaLeftMargin

        cell_obj = Calc.get_cell_obj("E3")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Borders(diagonal_down=side, diagonal_up=side)
        Styler.apply(cell, Borders(padding=Padding(left=1.2, right=1.2, top=0.5, bottom=0.5)))
        cp = cast("ParagraphProperties", cell)
        # padding may not apply exact
        assert cp.ParaLeftMargin >= 118 and cp.ParaLeftMargin <= 122
        assert cp.ParaRightMargin == cp.ParaLeftMargin

        assert cp.ParaTopMargin >= 48 and cp.ParaTopMargin <= 52
        assert cp.ParaBottomMargin == cp.ParaTopMargin

        cell_obj = Calc.get_cell_obj("G3")
        Calc.set_val(value="Hello", sheet=sheet, cell_obj=cell_obj)
        cell = Calc.get_cell(sheet, cell_obj)
        cb = Borders(diagonal_down=side, diagonal_up=side)
        Styler.apply(cell, Borders.default)

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_calc_border_range(loader) -> None:
    delay = 0  # 0 if Lo.bridge_connector.headless else 5_000
    from ooodev.office.calc import Calc

    doc = Calc.create_doc()
    try:
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_200_PERCENT)

        rng_obj = Calc.get_range_obj("B2:G6")
        cr = Calc.get_cell_range(sheet, rng_obj)

        # for some unknown reason LibreOffice is overriding style of horizontal Side. to match outter border.
        # Soluttion is to create a new border with only horizontal side set after inital range has been set
        cb = Borders(
            border_side=Side(line=BorderLineStyleEnum.SOLID, color=CommonColor.BLUE),
            vertical=Side(color=CommonColor.RED, line=BorderLineStyleEnum.DASHED),
            horizontal=Side(color=CommonColor.GREEN, width=1.4, line=BorderLineStyleEnum.DOUBLE),
        )
        Styler.apply(cr, cb)

        rng = cast("CellRange", cr)

        assert rng.TableBorder2.TopLine.Color == CommonColor.BLUE
        assert rng.TableBorder2.RightLine.LineStyle == BorderLineStyleEnum.SOLID.value

        cb = Borders(horizontal=Side(color=CommonColor.GREEN, width=1.4, line=BorderLineStyleEnum.DOUBLE))
        Styler.apply(cr, cb)

        assert rng.TableBorder2.VerticalLine.Color == CommonColor.RED
        assert rng.TableBorder2.VerticalLine.LineStyle == BorderLineStyleEnum.DASHED.value

        assert rng.TableBorder2.HorizontalLine.Color == CommonColor.GREEN
        assert rng.TableBorder2.HorizontalLine.LineStyle == BorderLineStyleEnum.DOUBLE.value

        rng_obj = Calc.get_range_obj("B8:G12")
        cr = Calc.get_cell_range(sheet, rng_obj)

        if not Lo.bridge_connector.headless:
            Calc.goto_cell(cell_obj=rng_obj.cell_start, doc=doc)

        cb = Borders(border_side=Side(), diagonal_up=Side(color=CommonColor.RED))
        Styler.apply(cr, cb)

        cell = Calc.get_cell(sheet=sheet, cell_obj=rng_obj.cell_start)

        cp = cast("CellProperties", cell)
        assert cp.DiagonalBLTR.Color == CommonColor.RED

        rng_obj = Calc.get_range_obj("B14:G18")
        cr = Calc.get_cell_range(sheet, rng_obj)

        if not Lo.bridge_connector.headless:
            Calc.goto_cell(cell_obj=rng_obj.cell_start, doc=doc)

        cb = Borders(
            border_side=Side(), diagonal_up=Side(color=CommonColor.RED), diagonal_down=Side(color=CommonColor.RED)
        )
        Styler.apply(cr, cb)

        rng_obj = Calc.get_range_obj("c15:F17")
        cr = Calc.get_cell_range(sheet, rng_obj)

        if not Lo.bridge_connector.headless:
            Calc.goto_cell(cell_obj=rng_obj.cell_start, doc=doc)

        Styler.apply(cr, Borders.empty)

        cell = Calc.get_cell(sheet=sheet, cell_obj=rng_obj.cell_start)
        cp = cast("CellProperties", cell)
        assert cp.LeftBorder2.LineWidth == 0
        assert cp.TopBorder2.LineWidth == 0
        assert cp.RightBorder2.LineWidth == 0
        assert cp.BottomBorder.LineWidth == 0

        para = cast("ParagraphProperties", cell)
        assert para.ParaLeftMargin == 35
        assert para.ParaRightMargin == 35
        assert para.ParaTopMargin == 35
        assert para.ParaBottomMargin == 35

        cb = Borders(border_side=Side(color=CommonColor.GREEN).line_double_thin)

        Styler.apply(cr, cb)

        cell = Calc.get_cell(sheet=sheet, cell_obj=rng_obj.cell_start)
        cp = cast("CellProperties", cell)
        assert cp.LeftBorder2.Color == CommonColor.GREEN
        assert cp.LeftBorder2.LineStyle == BorderLineStyleEnum.DOUBLE_THIN
        assert cp.TopBorder2.Color == CommonColor.GREEN
        assert cp.TopBorder2.LineStyle == BorderLineStyleEnum.DOUBLE_THIN

        cell = Calc.get_cell(sheet=sheet, cell_obj=rng_obj.cell_end)
        cp = cast("CellProperties", cell)
        assert cp.RightBorder2.Color == CommonColor.GREEN
        assert cp.RightBorder2.LineStyle == BorderLineStyleEnum.DOUBLE_THIN
        assert cp.BottomBorder2.Color == CommonColor.GREEN
        assert cp.BottomBorder2.LineStyle == BorderLineStyleEnum.DOUBLE_THIN

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
