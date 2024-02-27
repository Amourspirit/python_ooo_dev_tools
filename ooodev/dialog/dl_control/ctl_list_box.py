# region imports
from __future__ import annotations
from typing import Any, cast, Iterable, TYPE_CHECKING, Tuple
import contextlib
import uno

from ooodev.adapter.awt.uno_control_list_box_model_partial import UnoControlListBoxModelPartial
from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.adapter.awt.item_events import ItemEvents

from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind

from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlListBox  # service
    from com.sun.star.awt import UnoControlListBoxModel  # service
# endregion imports


class CtlListBox(DialogControlBase, UnoControlListBoxModelPartial, ActionEvents, ItemEvents):
    """Class for ListBox Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlListBox) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlListBox): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlListBoxModelPartial.__init__(self)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_action_events_listener_add_remove)
        ItemEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_action_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addActionListener(self.events_listener_action)
        event.remove_callback = True

    def _on_item_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addItemListener(self.events_listener_item)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlListBox:
        return cast("UnoControlListBox", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlListBox``"""
        return "com.sun.star.awt.UnoControlListBox"

    def get_model(self) -> UnoControlListBoxModel:
        """Gets the Model for the control"""
        return cast("UnoControlListBoxModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.LIST_BOX``"""
        return DialogControlKind.LIST_BOX

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.LIST_BOX``"""
        return DialogControlNamedKind.LIST_BOX

    # endregion Overrides

    # region Methods

    def set_list_data(self, data: Iterable[str]) -> None:
        """
        Sets the list data

        Args:
            data (Iterable[str]): List box entries
        """
        if not data:
            return
        ctl_props = self.get_control_props()
        if ctl_props is None:
            return
        uno_strings = uno.Any("[]string", tuple(data))  # type: ignore
        uno.invoke(ctl_props, "setPropertyValue", ("StringItemList", uno_strings))  # type: ignore

    # endregion Methods

    # region Properties

    @property
    def list_count(self) -> int:
        """Gets the number of items in the list"""
        with contextlib.suppress(Exception):
            return self.view.getItemCount()
        return 0

    @property
    def list_index(self) -> int:
        """
        Gets which item is selected

        Returns:
            Index of the first selected item or ``-1`` if no items are selected.
        """
        with contextlib.suppress(Exception):
            selected_items = self.selected_items
            if len(selected_items) > 0:
                return selected_items[0]
        return -1

    @property
    def model(self) -> UnoControlListBoxModel:
        # pylint: disable=no-member
        return cast("UnoControlListBoxModel", super().model)

    @property
    def row_source(self) -> Tuple[str, ...]:
        """
        Gets/Sets the row source.

        When setting the row source, the list box will be cleared and the new items will be added.

        The value passed in can be any iterable string such as a list or tuple of strings.
        """
        with contextlib.suppress(Exception):
            return self.model.StringItemList
        return ()

    @row_source.setter
    def row_source(self, value: Iterable[str]) -> None:
        """Sets the row source"""
        self.set_list_data(value)

    @property
    def view(self) -> UnoControlListBox:
        # pylint: disable=no-member
        return cast("UnoControlListBox", super().view)

    # item_count was renamed to list_count in 0.13.2
    item_count = list_count
    # endregion Properties


# ctl = CtlButton(None)
