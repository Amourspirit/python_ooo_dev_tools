from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.chart.chart_data_change_event_events import ChartDataChangeEventEvents
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs


if TYPE_CHECKING:
    from com.sun.star.sheet import SheetCellRange  # service


class SheetCellRangeComp(
    ComponentBase, ModifyEvents, ChartDataChangeEventEvents, PropertyChangeImplement, VetoableChangeImplement
):
    """
    Class for managing Sheet Cell Range Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SheetCellRange) -> None:
        """
        Constructor

        Args:
            component (SheetCellRange): UNO Sheet Cell Range Component
        """
        ComponentBase.__init__(self, component)
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        ChartDataChangeEventEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_chart_data_change_event_add_remove
        )
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Lazy Listeners
    def _on_modify_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addModifyListener(self.events_listener_modify)
        event.remove_callback = True

    def _on_chart_data_change_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addChartDataChangeEventListener(self.events_listener_chart_data_change_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SheetCellRange",)

    # endregion Overrides

    # region Properties

    @property
    def component(self) -> SheetCellRange:
        """Sheet Cell Range Component"""
        return cast("SheetCellRange", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
