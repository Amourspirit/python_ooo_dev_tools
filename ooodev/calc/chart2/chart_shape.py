from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from com.sun.star.graphic import XGraphic

from ooodev.mock import mock_g
from ooodev.exceptions import ex as mEx
from ooodev.draw.shapes.ole2_shape import OLE2Shape
from ooodev.format.inner.partial.position_size.draw.position_partial import PositionPartial
from ooodev.format.inner.partial.position_size.draw.size_partial import SizePartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial

# Line Properties not supported by OLE2Shape
# from ooodev.format.inner.partial.borders.draw.line_properties_partial import LinePropertiesPartial

if TYPE_CHECKING:
    from com.sun.star.drawing import XShape
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.table_chart import TableChart
    from ooodev.calc.chart2.chart_image import ChartImage
else:
    TableChart = Any


class ChartShape(OLE2Shape[TableChart], PositionPartial, SizePartial, CalcDocPropPartial, CalcSheetPropPartial):
    def __init__(self, owner: TableChart, component: XShape, lo_inst: LoInst | None = None) -> None:
        OLE2Shape.__init__(self, owner=owner, component=component, lo_inst=lo_inst)
        PositionPartial.__init__(self, factory_name="ooodev.draw.position", component=component, lo_inst=lo_inst)
        SizePartial.__init__(self, factory_name="ooodev.draw.size", component=component, lo_inst=lo_inst)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        CalcSheetPropPartial.__init__(self, obj=owner.calc_sheet)

    def get_image(self) -> ChartImage:
        """Returns the chart image."""
        from .chart_image import ChartImage

        try:
            graphic = self.lo_inst.qi(XGraphic, self.get_property("Graphic"), True)
            return ChartImage(owner=self, component=graphic, lo_inst=self.lo_inst)
        except Exception as e:
            raise mEx.ChartError(f"Failed to get chart image: {e}") from e


if mock_g.FULL_IMPORT:
    from .chart_image import ChartImage
