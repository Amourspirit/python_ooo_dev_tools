# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.view.selection_change_events import SelectionChangeEvents
from ooodev.adapter.awt.tree.tree_edit_events import TreeEditEvents
from ooodev.adapter.awt.tree.tree_expansion_events import TreeExpansionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt.tree import TreeControl  # service
    from com.sun.star.awt.tree import TreeControlModel  # service
# endregion imports


class CtlTree(CtlBase, SelectionChangeEvents, TreeEditEvents, TreeExpansionEvents):
    """Class for Tree Control"""

    # region init
    def __init__(self, ctl: TreeControl) -> None:
        """
        Constructor

        Args:
            ctl (TreeControl): Tree Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        SelectionChangeEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_selection_change_events_listener_add_remove
        )
        TreeEditEvents.__init__(self, trigger_args=generic_args, cb=self._on_tree_edit_events_listener_add_remove)
        TreeExpansionEvents.__init__(
            self, trigger_args=generic_args, cb=self._on_tree_expansion_events_listener_add_remove
        )

    # endregion init

    # region Lazy Listeners
    def _on_selection_change_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addSelectionChangeListener(self.events_listener_selection_change)
        self._add_listener(key)

    def _on_tree_edit_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addTreeEditListener(self.events_listener_tree_edit)
        self._add_listener(key)

    def _on_tree_expansion_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addTreeExpansionListener(self.events_listener_tree_expansion)
        self._add_listener(key)

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> TreeControl:
        return cast("TreeControl", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.TreeControl``"""
        return "com.sun.star.awt.TreeControl"

    def get_model(self) -> TreeControlModel:
        """Gets the Model for the control"""
        # Tree Control does have getModel() method even thought it is not documented
        return cast("TreeControlModel", self.get_view_ctl().getModel())  # type: ignore

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> TreeControl:
        return self.get_view_ctl()

    @property
    def model(self) -> TreeControlModel:
        return self.get_model()

    # endregion Properties
