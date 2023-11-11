# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFileControl  # service
    from com.sun.star.awt import UnoControlFileControlModel  # service
# endregion imports


class CtlFile(CtlBase, TextEvents):
    """Class for file Control"""

    # region init
    def __init__(self, ctl: UnoControlFileControl) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFileControl): File Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addTextListener(self.events_listener_text)
        self._add_listener(key)

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlFileControl:
        return cast("UnoControlFileControl", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlFileControl``"""
        return "com.sun.star.awt.UnoControlFileControl"

    def get_model(self) -> UnoControlFileControlModel:
        """Gets the Model for the control"""
        return cast("UnoControlFileControlModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlFileControl:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlFileControlModel:
        return self.get_model()

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        return self.model.Text

    @text.setter
    def text(self, value: str) -> None:
        self.model.Text = value

    # endregion Properties
