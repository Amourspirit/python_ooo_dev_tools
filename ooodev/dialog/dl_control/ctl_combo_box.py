# region imports
from __future__ import annotations
from typing import Any, Iterable, Tuple, cast, TYPE_CHECKING
import contextlib

import uno
from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.adapter.awt.item_events import ItemEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind

from ooodev.events.args.listener_event_args import ListenerEventArgs

from .ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlComboBox  # service
    from com.sun.star.awt import UnoControlComboBoxModel  # service
# endregion imports


class CtlComboBox(DialogControlBase, ActionEvents, ItemEvents, TextEvents):
    """Class for ComboBox Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlComboBox) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlComboBox): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_action_events_listener_add_remove)
        ItemEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_events_listener_add_remove)
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)

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

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTextListener(self.events_listener_text)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlComboBox:
        return cast("UnoControlComboBox", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlComboBox``"""
        return "com.sun.star.awt.UnoControlComboBox"

    def get_model(self) -> UnoControlComboBoxModel:
        """Gets the Model for the control"""
        return cast("UnoControlComboBoxModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.COMBOBOX``"""
        return DialogControlKind.COMBOBOX

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.COMBOBOX``"""
        return DialogControlNamedKind.COMBOBOX

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
    def view(self) -> UnoControlComboBox:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlComboBoxModel:
        return self.get_model()

    @property
    def drop_down(self) -> bool:
        """Gets/Sets the combobox has a drop down button."""
        return self.model.Dropdown

    @drop_down.setter
    def drop_down(self, value: bool) -> None:
        self.model.Dropdown = value

    @property
    def max_text_len(self) -> int:
        """Gets/Sets the maximum character count."""
        return self.model.MaxTextLen

    @max_text_len.setter
    def max_text_len(self, value: int) -> None:
        self.model.MaxTextLen = value

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        with contextlib.suppress(Exception):
            return self.model.Text
        return ""

    @text.setter
    def text(self, value: str) -> None:
        self.model.Text = value

    @property
    def list_count(self) -> int:
        """Gets the number of items in the combo box"""
        with contextlib.suppress(Exception):
            items = self.model.StringItemList
            if items:
                return len(items)
        return 0

    @property
    def list_index(self) -> int:
        """
        Gets which item index is selected

        Returns:
            Index of the first selected item or ``-1`` if no items are selected.
        """
        text = self.text
        if not text:
            return -1
        with contextlib.suppress(Exception):
            items = self.model.StringItemList
            if items:
                return items.index(text)
        return -1

    @property
    def read_only(self) -> bool:
        """Gets/Sets the read-only property"""
        with contextlib.suppress(Exception):
            return self.model.ReadOnly
        return False

    @read_only.setter
    def read_only(self, value: bool) -> None:
        """Sets the read-only property"""
        with contextlib.suppress(Exception):
            self.model.ReadOnly = value

    @property
    def row_source(self) -> Tuple[str, ...]:
        """
        Gets/Sets the row source.

        When setting the row source, the combobox will be cleared and the new items will be added.

        The value passed in can be any iterable string such as a list or tuple of strings.
        """
        with contextlib.suppress(Exception):
            return self.model.StringItemList
        return ()

    @row_source.setter
    def row_source(self, value: Iterable[str]) -> None:
        """Sets the row source"""
        self.set_list_data(value)

    # endregion Properties


# ctl = CtlButton(None)
