from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from ooodev.exceptions import ex as mEx
from com.sun.star.graphic import XGraphic
from ooodev.draw.shapes.ole2_shape import OLE2Shape

if TYPE_CHECKING:
    from com.sun.star.drawing import XShape
    from ooodev.loader.inst.lo_inst import LoInst
    from .table_chart import TableChart
    from .chart_image import ChartImage
else:
    TableChart = Any


class ChartShape(OLE2Shape[TableChart]):
    def __init__(self, owner: TableChart, component: XShape, lo_inst: LoInst | None = None) -> None:
        super().__init__(owner=owner, component=component, lo_inst=lo_inst)

    def get_image(self) -> ChartImage:
        """Returns the chart image."""
        from .chart_image import ChartImage

        try:
            graphic = self.lo_inst.qi(XGraphic, self.get_property("Graphic"), True)
            return ChartImage(owner=self, component=graphic, lo_inst=self.lo_inst)
        except Exception as e:
            raise mEx.ChartError(f"Failed to get chart image: {e}") from e
