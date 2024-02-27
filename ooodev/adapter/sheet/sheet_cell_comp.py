from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.adapter.beans.property_change_implement import PropertyChangeImplement
from ooodev.adapter.beans.vetoable_change_implement import VetoableChangeImplement
from ooodev.adapter.text.text_partial import TextPartial
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.util.modify_events import ModifyEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs


if TYPE_CHECKING:
    from com.sun.star.sheet import SheetCell  # service


class SheetCellComp(ComponentBase, TextPartial, ModifyEvents, PropertyChangeImplement, VetoableChangeImplement):
    """
    Class for managing Sheet Cell Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: SheetCell) -> None:
        """
        Constructor

        Args:
            component (SheetCell): UNO Sheet Cell Component
        """
        ComponentBase.__init__(self, component)
        TextPartial.__init__(self, component=component, interface=None)
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)
        PropertyChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)
        VetoableChangeImplement.__init__(self, component=self.component, trigger_args=generic_args)

    # region Lazy Listeners
    def _on_modify_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addModifyListener(self.events_listener_modify)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_supported_service_names(self) -> tuple[str, ...]:
        """Returns a tuple of supported service names."""
        return ("com.sun.star.sheet.SheetCell",)

    # endregion Overrides
    # region Properties
    @property
    def component(self) -> SheetCell:
        """Sheet Cell Component"""
        # pylint: disable=no-member
        return cast("SheetCell", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
