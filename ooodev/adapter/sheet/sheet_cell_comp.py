from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.util.modify_events import ModifyEvents


if TYPE_CHECKING:
    from com.sun.star.sheet import SheetCell  # service


class SheetCellComp(ComponentBase, ModifyEvents):
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
        generic_args = self._get_generic_args()
        ModifyEvents.__init__(self, trigger_args=generic_args, cb=self._on_modify_events_add_remove)

    # region Manage Events

    # region Lazy Listeners
    def _on_modify_events_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addModifyListener(self.events_listener_modify)
        event.remove_callback = True

    # endregion Lazy Listeners

    @property
    def component(self) -> SheetCell:
        """Tree Data Model Component"""
        return cast("SheetCell", self._get_component())

    # endregion Manage Events
