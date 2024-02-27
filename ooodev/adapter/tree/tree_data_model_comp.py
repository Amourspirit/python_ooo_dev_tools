from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno
from com.sun.star.awt.tree import XTreeDataModel
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.adapter.component_base import ComponentBase
from ooodev.loader import lo as mLo
from ooodev.adapter.tree.tree_data_model_events import TreeDataModelEvents

if TYPE_CHECKING:
    from com.sun.star.lang import XComponent


class TreeDataModelComp(ComponentBase, TreeDataModelEvents):
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
        # pylint: disable=no-member
        generic_args = self._ComponentBase__get_generic_args()  # type: ignore
        TreeDataModelEvents.__init__(self, trigger_args=generic_args, cb=self._on_tree_model_listener_add_remove)

    # region Lazy Listeners
    def _on_tree_model_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.component.addTreeDataModelListener(self.events_listener_tree_data_model)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def _ComponentBase__get_is_supported(self, component: XComponent) -> bool:
        if not component:
            return False
        return mLo.Lo.is_uno_interfaces(component, XTreeDataModel)

    # endregion Overrides

    # region Properties
    @property
    def component(self) -> XTreeDataModel:
        """Tree Data Model Component"""
        # pylint: disable=no-member
        return cast("XTreeDataModel", self._ComponentBase__get_component())  # type: ignore

    # endregion Properties
