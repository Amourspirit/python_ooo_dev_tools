from __future__ import annotations
from typing import Any, TYPE_CHECKING
import uno

from ooodev.mock import mock_g
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.table.table_chart_comp import TableChartComp
from ooodev.events.partial.events_partial import EventsPartial
from ooodev.exceptions import ex as mEx
from ooodev.loader import lo as mLo
from ooodev.office import chart2 as mChart2
from ooodev.gui import gui as mGui
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.the_dictionary_partial import TheDictionaryPartial
from ooodev.calc.partial.calc_doc_prop_partial import CalcDocPropPartial
from ooodev.calc.partial.calc_sheet_prop_partial import CalcSheetPropPartial
from ooodev.calc.chart2.chart_doc import ChartDoc


if TYPE_CHECKING:
    from ooodev.calc import CalcSheet
    from ooodev.loader.inst.lo_inst import LoInst
    from ooodev.calc.chart2.chart_draw_page import ChartDrawPage
    from ooodev.calc.chart2.chart_shape import ChartShape


class TableChart(
    LoInstPropsPartial,
    TableChartComp,
    EventsPartial,
    PropPartial,
    QiPartial,
    ServicePartial,
    TheDictionaryPartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    CalcSheetPropPartial,
    CalcDocPropPartial,
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
        TheDictionaryPartial.__init__(self)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore # pylint: disable=no-member
        PropertyChangeImplement.__init__(self, component=component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=component, trigger_args=generic_args)
        CalcSheetPropPartial.__init__(self, obj=owner)
        CalcDocPropPartial.__init__(self, obj=owner.calc_doc)
        self._chart_doc = None
        self._shape = None
        self._draw_page = None

    def copy_chart(self) -> None:
        """
        Copies the chart to the clipboard using a dispatch command.

        Raises:
            ChartError: If error occurs.

        Returns:
            None:
        """
        try:
            supp = mGui.GUI.get_selection_supplier(self.calc_sheet.calc_doc.component)
            supp.select(self.shape.component)
            self.lo_inst.dispatch_cmd("Copy")
        except Exception as e:
            raise mEx.ChartError("Error in attempt to copy chart") from e

    # region Properties

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
            self._chart_doc = ChartDoc(owner=self, component=embedded, lo_inst=self.lo_inst)
        return self._chart_doc

    @property
    def shape(self) -> ChartShape:
        """OLE2 Shape."""
        if self._shape is None:
            # pylint: disable=import-outside-toplevel
            from .chart_shape import ChartShape

            shape = mChart2.Chart2.get_chart_shape(sheet=self.calc_sheet.component, chart_name=self.name)
            self._shape = ChartShape(owner=self, component=shape, lo_inst=self.lo_inst)
        return self._shape  # type: ignore

    @property
    def draw_page(self) -> ChartDrawPage:
        """Draw Page."""
        if self._draw_page is None:
            # pylint: disable=import-outside-toplevel
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


if mock_g.FULL_IMPORT:
    from com.sun.star.drawing import XDrawPageSupplier
    from com.sun.star.embed import XComponentSupplier
    from .chart_shape import ChartShape
    from .chart_draw_page import ChartDrawPage
