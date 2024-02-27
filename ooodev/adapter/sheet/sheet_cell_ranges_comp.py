from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.chart.chart_data_change_event_events import ChartDataChangeEventEvents
from ooodev.adapter.component_base import ComponentBase
from ooodev.events.args.listener_event_args import ListenerEventArgs


if TYPE_CHECKING:
    from com.sun.star.sheet import SheetCellRanges  # service


class SheetCellRangesComp(ComponentBase, ChartDataChangeEventEvents, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing Sheet Cell Ranges Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SheetCellRanges) -> None:
        """
        Constructor

        Args:
            component (SheetCellRanges): UNO Sheet Cell Ranges Component
        """
        ComponentBase.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ChartDataChangeEventEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_chart_data_change_event_add_remove
        )
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Lazy Listeners
    def _on_chart_data_change_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addChartDataChangeEventListener(self.events_listener_chart_data_change_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SheetCellRanges",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> SheetCellRanges:
        """Sheet Cell Ranges Component"""
        # pylint: disable=no-member
        return cast("SheetCellRanges", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
