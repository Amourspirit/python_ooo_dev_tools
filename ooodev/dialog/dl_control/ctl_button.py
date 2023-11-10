from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.events.args.event_args import EventArgs as EventArgs

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButton  # service
    from com.sun.star.awt import UnoControlButtonModel  # service


class CtlButton(CtlBase, ActionEvents):
    def __init__(self, ctl: UnoControlButton) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlButton): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        ActionEvents.__init__(self, trigger_args=generic_args)
        self.view.addActionListener(self.events_listener_action)

    def get_view_ctl(self) -> UnoControlButton:
        return cast("UnoControlButton", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlButton``"""
        return "com.sun.star.awt.UnoControlButton"

    @property
    def view(self) -> UnoControlButton:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlButtonModel:
        return cast("UnoControlButtonModel", self.get_view_ctl().getModel())


# ctl = CtlButton(None)
