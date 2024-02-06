from __future__ import annotations
from typing import Any, TYPE_CHECKING
from ooodev.loader import lo as mLo
from ooodev.adapter.chart2.chart_document_comp import ChartDocumentComp
from ooodev.utils.partial.lo_inst_props_partial import LoInstPropsPartial
from ooodev.utils.partial.qi_partial import QiPartial
from ooodev.utils.partial.service_partial import ServicePartial
from ooodev.utils.partial.prop_partial import PropPartial
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.util.close_events import CloseEvents
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.document.storage_change_event_events import StorageChangeEventEvents
from ooodev.office import chart2 as mChart2
from ooodev.utils.color import Color

if TYPE_CHECKING:
    from com.sun.star.chart2 import ChartDocument  # service
    from ooodev.loader.inst.lo_inst import LoInst
    from .chart_title import ChartTitle
    from .chart_axis import ChartAxis


class ChartDoc(
    LoInstPropsPartial,
    ChartDocumentComp,
    ModifyEvents,
    PropPartial,
    QiPartial,
    ServicePartial,
    PropertyChangeImplement,
    VetoableChangeImplement,
    CloseEvents,
    StorageChangeEventEvents,
):
    """
    Class for managing Chart2 ChartDocument Component.
    """

    def __init__(self, component: ChartDocument, lo_inst: LoInst | None = None) -> None:
        """
        Constructor

        Args:
            component (Any): UNO Chart2 ChartDocument Component.
        """
        if lo_inst is None:
            lo_inst = mLo.Lo.current_lo
        LoInstPropsPartial.__init__(self, lo_inst=lo_inst)
        ChartDocumentComp.__init__(self, component=component)  # type: ignore
        PropPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        QiPartial.__init__(self, component=component, lo_inst=self.lo_inst)
        ServicePartial.__init__(self, component=component, lo_inst=self.lo_inst)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        CloseEvents.__init__(self, trigger_args=generic_args, cb=self._on_close_events_add_remove)
        StorageChangeEventEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_storage_change_events_add_remove
        )
        PropertyChangeImplement.__init__(self, component=component, trigger_args=generic_args)  # type: ignore
        VetoableChangeImplement.__init__(self, component=component, trigger_args=generic_args)  # type: ignore
        self._axis_x = None
        self._axis_y = None

    # region Lazy Listeners

    def _on_modify_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addModifyListener(self.events_listener_modify)  # type: ignore
        event.remove_callback = True

    def _on_close_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addCloseListener(self.events_listener_close)  # type: ignore
        event.remove_callback = True

    def _on_storage_change_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addStorageChangeListener(self.events_listener_storage_change_event)  # type: ignore
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Methods
    def set_title(self, title: str) -> ChartTitle:
        """Adds a Chart Title."""
        from com.sun.star.chart2 import XTitled
        from com.sun.star.chart2 import XTitle
        from com.sun.star.chart2 import XFormattedString
        from .chart_title import ChartTitle

        x_title = self.lo_inst.create_instance_mcf(XTitle, "com.sun.star.chart2.Title", raise_err=True)
        x_title_str = self.lo_inst.create_instance_mcf(
            XFormattedString, "com.sun.star.chart2.FormattedString", raise_err=True
        )
        x_title_str.setString(title)

        title_arr = (x_title_str,)
        x_title.setText(title_arr)

        titled = self.qi(XTitled, True)
        titled.setTitleObject(x_title)
        return ChartTitle(owner=self, component=titled.getTitleObject(), lo_inst=self.lo_inst)

    def get_title(self) -> ChartTitle[ChartDoc] | None:
        """Gets the Chart Title Component."""
        from com.sun.star.chart2 import XTitled
        from .chart_title import ChartTitle

        titled = self.qi(XTitled, True)
        comp = titled.getTitleObject()
        if comp is None:
            return None
        return ChartTitle(owner=self, component=comp, lo_inst=self.lo_inst)

    def set_bg_color(self, color: Color) -> None:
        """Sets the background color."""
        mChart2.Chart2.set_background_colors(self.component, bg_color=color, wall_color=Color(-1))

    def set_wall_color(self, color: Color) -> None:
        """Sets the wall color."""
        mChart2.Chart2.set_background_colors(self.component, bg_color=Color(-1), wall_color=color)

    @property
    def axis_x(self) -> ChartAxis:
        """Gets the X Axis Component."""
        if self._axis_x is None:
            from .chart_axis import ChartAxis

            axis = mChart2.Chart2.get_x_axis(self.component)
            self._axis_x = ChartAxis(owner=self, component=axis, lo_inst=self.lo_inst)
        return self._axis_x

    @property
    def axis_y(self) -> ChartAxis:
        """Gets the Y Axis Component."""
        if self._axis_y is None:
            from .chart_axis import ChartAxis

            axis = mChart2.Chart2.get_x_axis(self.component)
            self._axis_y = ChartAxis(owner=self, component=axis, lo_inst=self.lo_inst)
        return self._axis_y

    # endregion Methods
