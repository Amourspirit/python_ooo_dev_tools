from __future__ import annotations
from typing import TYPE_CHECKING

from ooodev.adapter.chart2.error_bar_comp import ErrorBarComp
from ooodev.loader import lo as mLo
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.adapter.chart2.data.data_sink_partial import DataSinkPartial

if TYPE_CHECKING:
    from com.sun.star.beans import XPropertySet
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_doc import ChartDoc


class ChartErrorBar(LoInstPropsPartial, ErrorBarComp, DataSinkPartial, PropPartial, QiPartial, ServicePartial):
    """
    Class for managing Chart2 ErrorBar.
    """

    def __init__(
        self, chart_doc: ChartDoc, lo_inst: LoInst | None = None, component: XPropertySet | None = None
    ) -> None:
        """
        Constructor

        Args:
            lo_inst (mLo.LoInst): Lo Instance. This instance is used to create ``component`` is it is not provided.
            component (XPropertySet, optional): UNO Chart2 ErrorBar Component. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        ErrorBarComp.__init__(self, lo_inst=lo_inst, component=component)  # type: ignore
        DataSinkPartial.__init__(self, component=component)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        self._chart_doc = chart_doc

    @property
    def chart_doc(self) -> ChartDoc:
        """Chart Document."""
        return self._chart_doc
