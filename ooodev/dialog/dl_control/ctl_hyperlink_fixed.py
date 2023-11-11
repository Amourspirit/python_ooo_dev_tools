# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.awt.action_events import ActionEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedHyperlink  # service
    from com.sun.star.awt import UnoControlFixedHyperlinkModel  # service
# endregion imports


class CtlHyperlinkFixed(CtlBase, ActionEvents):
    """Class for Fixed Hyperlink Control"""

    # region init
    def __init__(self, ctl: UnoControlFixedHyperlink) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFixedHyperlink): Button Control
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
    def get_view_ctl(self) -> UnoControlFixedHyperlink:
        return cast("UnoControlFixedHyperlink", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlFixedHyperlink``"""
        return "com.sun.star.awt.UnoControlFixedHyperlink"

    def get_model(self) -> UnoControlFixedHyperlinkModel:
        """Gets the Model for the control"""
        return cast("UnoControlFixedHyperlinkModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlFixedHyperlink:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlFixedHyperlinkModel:
        return self.get_model()

    @property
    def label(self) -> str:
        """
        Gets/Sets the label of the control.
        """
        return self.model.Label

    @label.setter
    def label(self, value: str) -> None:
        self.model.Label = value

    @property
    def url(self) -> str:
        """
        Gets/Sets the URL to be opened when the hyperlink is activated.
        """
        return self.model.URL

    @url.setter
    def url(self, value: str) -> None:
        self.model.URL = value

    # endregion Properties
