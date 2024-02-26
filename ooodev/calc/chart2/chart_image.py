from __future__ import annotations
from typing import TYPE_CHECKING
import uno
from ooodev.adapter.graphic.graphic_comp import GraphicComp
from ooodev.loader import lo as mLo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial


if TYPE_CHECKING:
    from com.sun.star.graphic import XGraphic
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.chart_shape import ChartShape


class ChartImage(LoInstPropsPartial, GraphicComp, QiPartial, ServicePartial, CalcDocPropPartial, CalcSheetPropPartial):
    """
    Class for managing Chart2 Chart Image Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, owner: ChartShape, component: XGraphic, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            owner (ChartShape): Chart Shape.
            component (Any): UNO Chart2 Chart Image Component.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        GraphicComp.__init__(self, component)  # type: ignore
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        CalcSheetPropPartial.__init__(self, obj=owner.calc_sheet)
        self._owner = owner

    @property
    def owner(self) -> ChartShape:
        """Chart Shape"""
        return self._owner
