from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.adapter.lang.event_events import EventEvents
from .tree_data_model_events import TreeDataModelEvents

if TYPE_CHECKING:
    from com.sun.star.awt.tree import XTreeDataModel


class TreeDataModelComp(ComponentBase, EventEvents, TreeDataModelEvents):
    """
    Class for managing Tree Data Model Component.
    """

    # pylint: disable=unused-argument

    def __init__(self, component: XTreeDataModel) -> None:
        """
        Constructor

        Args:
            component (XTreeDataModel): Tree Data Model Component
        """
        ComponentBase.__init__(self, component)
        generic_args = self._get_generic_args()
        EventEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_event_events_listener_add_remove)
        TreeDataModelEvents.__init__(self, trigger_args=generic_args, cb=self._on_tree_model_listener_add_remove)

    # region Manage Events

    # region Lazy Listeners
    def _on_tree_model_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addTreeDataModelListener(self.events_listener_tree_data_model)
        event.remove_callback = True

    def _on_item_event_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addEventListener(self.events_listener_event)
        event.remove_callback = True

    # endregion Lazy Listeners

    @property
    def component(self) -> XTreeDataModel:
        """Tree Data Model Component"""
        return cast("XTreeDataModel", self._get_component())

    # endregion Manage Events
