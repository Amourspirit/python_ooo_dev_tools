from __future__ import annotations
import pytest
from typing import cast
from pathlib import Path

if __name__ == "__main__":
    pytest.main([__file__])


# from com.sun.star.lang import XMultiServiceFactory
# from com.sun.star.container import XNameContainer
from ooodev.utils.kind.chart2_types import ChartTypes

try:
    from ooodev.office.chart2 import Chart2
except ImportError:
    Chart2 = None

from ooodev.gui.gui import GUI
from ooodev.loader.lo import Lo
from ooodev.utils.info import Info


def test_insert_chart(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5) or Chart2 is None:
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    delay = 0
    from ooodev.office.calc import Calc

    fix_path = cast(Path, copy_fix_calc("chartsData.ods"))

    doc = None
    try:
        doc = Calc.open_doc(fix_path)
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

        rng_data = Calc.get_range_obj("A2:B8")
        chart_doc = Chart2.insert_chart(
            sheet=sheet, cells_range=rng_data.get_cell_range_address(), diagram_name=ChartTypes.Column.DEFAULT
        )
        assert chart_doc is not None

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        if doc is not None:
            Lo.close_doc(doc)


def test_insert_chart_named_sheet(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5) or Chart2 is None:
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    delay = 0
    from ooodev.office.calc import Calc

    fix_path = cast(Path, copy_fix_calc("chartsData.ods"))

    doc = Calc.open_doc(fix_path)
    try:
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

        Calc.set_sheet_name(sheet=sheet, name="MySheet")

        rng_data = Calc.get_range_obj("A2:B8")
        chart_doc = Chart2.insert_chart(
            sheet=sheet, cells_range=rng_data.get_cell_range_address(), diagram_name=ChartTypes.Column.DEFAULT
        )
        assert chart_doc is not None

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)


def test_insert_chart_named_sheet_with_spaces(loader, copy_fix_calc) -> None:
    if Info.version_info < (7, 5) or Chart2 is None:
        pytest.skip("Not supported in this version, Requires LibreOffice 7.5 or higher.")

    delay = 0
    from ooodev.office.calc import Calc

    fix_path = cast(Path, copy_fix_calc("chartsData.ods"))

    doc = Calc.open_doc(fix_path)
    try:
        sheet = Calc.get_sheet(doc)
        if not Lo.bridge_connector.headless:
            GUI.set_visible()
            Lo.delay(500)
            Calc.zoom(doc, GUI.ZoomEnum.ZOOM_100_PERCENT)

        Calc.set_sheet_name(sheet=sheet, name="My Special Sheet")

        rng_data = Calc.get_range_obj("A2:B8")
        chart_doc = Chart2.insert_chart(
            sheet=sheet, cells_range=rng_data.get_cell_range_address(), diagram_name=ChartTypes.Column.DEFAULT
        )
        assert chart_doc is not None

        if not Lo.bridge_connector.headless:
            # The chart does not always update correctly after style changes.
            # To refresh the chart, we can do a recalculation by calling the dispatch_recalculate() method.
            Calc.dispatch_recalculate()

        Lo.delay(delay)
    finally:
        Lo.close_doc(doc)
