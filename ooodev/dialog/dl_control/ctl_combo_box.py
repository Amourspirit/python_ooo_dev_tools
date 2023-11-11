# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.adapter.awt.item_events import ItemEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlComboBox  # service
    from com.sun.star.awt import UnoControlComboBoxModel  # service
# endregion imports


class CtlComboBox(CtlBase, ActionEvents, ItemEvents, TextEvents):
    """Class for ComboBox Control"""

    # region init
    def __init__(self, ctl: UnoControlComboBox) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlComboBox): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_action_events_listener_add_remove)
        ItemEvents.__init__(self, trigger_args=generic_args, cb=self._on_item_events_listener_add_remove)
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_action_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addActionListener(self.events_listener_action)
        self._add_listener(key)

    def _on_item_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addItemListener(self.events_listener_item)
        self._add_listener(key)

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addTextListener(self.events_listener_text)
        self._add_listener(key)

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

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlComboBox:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlComboBoxModel:
        return self.get_model()

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        return self.model.Text

    @text.setter
    def text(self, value: str) -> None:
        self.model.Text = value

    # endregion Properties


# ctl = CtlButton(None)
