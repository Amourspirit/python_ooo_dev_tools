from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.table.table_chart_comp import TableChartComp
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from .chart_doc import ChartDoc


if TYPE_CHECKING:
    from ooodev.calc import CalcSheet
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_draw_page import ChartDrawPage
    from .chart_shape import ChartShape


class TableChart(
    LoInstPropsPartial,
    TableChartComp,
    EventsPartial,
    PropPartial,
    QiPartial,
    ServicePartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
):
    """
    Class for managing table Chart.
    """

    def __init__(self, owner: CalcSheet, component: Any, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (XTableChart): UNO Component that supports ``com.sun.star.table.TableChart`` service.
            lo_inst (LoInst, optional): Lo Instance. Use when creating multiple documents. Defaults to None.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        TableChartComp.__init__(self, component)
        EventsPartial.__init__(self)
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        PropertyChangeImplement.__init__(self, component=component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=component, trigger_args=generic_args)
        self._owner = owner
        self._chart_doc = None
        self._shape = None
        self._draw_page = None

    # region Properties
    @property
    def calc_sheet(self) -> CalcSheet:
        """Sheet that owns this cell."""
        return self._owner

    @property
    def name(self) -> str:
        """Name of the chart."""
        return self.component.getName()

    @name.setter
    def name(self, value: str) -> None:
        """Sets the name of the chart."""
        self.component.setName(value)

    @property
    def chart_doc(self) -> ChartDoc:
        """Chart Document."""
        if self._chart_doc is None:
            embedded = self.get_property("EmbeddedObject")
            if embedded is None:
                raise mEx.PropertyNotFoundError("EmbeddedObject")
            self._chart_doc = ChartDoc(component=embedded, lo_inst=self.lo_inst)
        return self._chart_doc

    @property
    def shape(self) -> ChartShape:
        """OLE2 Shape."""
        if self._shape is None:
            from .chart_shape import ChartShape

            shape = mChart2.Chart2.get_chart_shape(sheet=self.calc_sheet.component, chart_name=self.name)
            self._shape = ChartShape(owner=self, component=shape, lo_inst=self.lo_inst)
        return self._shape  # type: ignore

    @property
    def draw_page(self) -> ChartDrawPage:
        """Draw Page."""
        if self._draw_page is None:
            from com.sun.star.drawing import XDrawPageSupplier
            from com.sun.star.embed import XComponentSupplier
            from .chart_draw_page import ChartDrawPage

            shape = self.shape
            embedded_chart = shape.get_property("EmbeddedObject")
            comp_supp = mLo.Lo.qi(XComponentSupplier, embedded_chart, True)
            x_closable = comp_supp.getComponent()
            supp_page = self.lo_inst.qi(XDrawPageSupplier, x_closable, True)
            self._draw_page = ChartDrawPage(owner=self, component=supp_page.getDrawPage(), lo_inst=self.lo_inst)
        return self._draw_page

    # endregion Properties
