from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.chart.chart_data_change_event_events import ChartDataChangeEventEvents
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs


if TYPE_CHECKING:
    from com.sun.star.sheet import SheetCellCursor  # service


class SheetCellCursorComp(
    ComponentBase, ModifyEvents, ChartDataChangeEventEvents, PropertyChangeImplement, VetoableChangeImplement
):
    """
    Class for managing Sheet Cell Cursor Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SheetCellCursor) -> None:
        """
        Constructor

        Args:
            component (SheetCellCursor): UNO Sheet Cell Cursor Component
        """
        ComponentBase.__init__(self, component)
        # pylint: disable=no-member
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
        return ("com.sun.star.sheet.SheetCellCursor",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> SheetCellCursor:
        """Sheet Cell Cursor Component"""
        # pylint: disable=no-member
        return cast("SheetCellCursor", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
