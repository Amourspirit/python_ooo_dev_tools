# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.events.args.event_args import EventArgs as EventArgs
from ooodev.events.args.listener_event_args import ListenerEventArgs

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlButton  # service
    from com.sun.star.awt import UnoControlButtonModel  # service
# endregion imports


class CtlButton(CtlBase, ActionEvents):
    """Class for Button Control"""

    # region init
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
        ActionEvents.__init__(self, trigger_args=generic_args, cb=self._on_action_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_action_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addActionListener(self.events_listener_action)
        self._add_listener(key)

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlButton:
        return cast("UnoControlButton", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlButton``"""
        return "com.sun.star.awt.UnoControlButton"

    def get_model(self) -> UnoControlButtonModel:
        """Gets the Model for the control"""
        return cast("UnoControlButtonModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlButton:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlButtonModel:
        return self.get_model()

    # endregion Properties
