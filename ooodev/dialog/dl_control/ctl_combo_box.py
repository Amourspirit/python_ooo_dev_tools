from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.adapter.awt.item_events import ItemEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.event_args import EventArgs as EventArgs

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlComboBox  # service
    from com.sun.star.awt import UnoControlComboBoxModel  # service


class CtlComboBox(CtlBase, ActionEvents, ItemEvents, TextEvents):
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
        ActionEvents.__init__(self, trigger_args=generic_args)
        ItemEvents.__init__(self, trigger_args=generic_args)
        TextEvents.__init__(self, trigger_args=generic_args)
        self.view.addActionListener(self.events_listener_action)
        self.view.addItemListener(self.events_listener_item)
        self.view.addTextListener(self.events_listener_text)

    def get_view_ctl(self) -> UnoControlComboBox:
        return cast("UnoControlComboBox", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlComboBox``"""
        return "com.sun.star.awt.UnoControlComboBox"

    @property
    def view(self) -> UnoControlComboBox:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlComboBoxModel:
        return cast("UnoControlComboBoxModel", self.get_view_ctl().getModel())

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        return self.model.Text

    @text.setter
    def text(self, value: str) -> None:
        self.model.Text = value


# ctl = CtlButton(None)
