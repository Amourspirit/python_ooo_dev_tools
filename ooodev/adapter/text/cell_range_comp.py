from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.chart.chart_data_change_event_events import ChartDataChangeEventEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase


if TYPE_CHECKING:
    from com.sun.star.text import CellRange  # service
    from com.sun.star.table import XCellRange


class CellRangeComp(ComponentBase, ChartDataChangeEventEvents):
    """
    Class for managing CellRange Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XCellRange) -> None:
        """
        Constructor

        Args:
            component (XCellRange): UNO Component that support ``com.sun.star.text.CellRange`` service.
        """

        ComponentBase.__init__(self, component)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ChartDataChangeEventEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_chart_data_change_event_add_remove
        )

    # region Lazy Listeners
    def _on_chart_data_change_event_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addChartDataChangeEventListener(self.events_listener_chart_data_change_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.text.CellRange",)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> CellRange:
        """CellRange Component"""
        # pylint: disable=no-member
        return cast("CellRange", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
