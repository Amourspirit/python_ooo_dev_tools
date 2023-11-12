# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

from .ctl_base import CtlListenerBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFormattedField  # service
    from com.sun.star.awt import UnoControlFormattedFieldModel  # service
# endregion imports


class CtlFormattedField(CtlListenerBase, SpinEvents, TextEvents):
    """Class for Formatted Field Control"""

    # region init
    def __init__(self, ctl: UnoControlFormattedField) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFormattedField): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlListenerBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        SpinEvents.__init__(self, trigger_args=generic_args, cb=self._on_spin_events_listener_add_remove)
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_spin_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addSpinListener(self.events_listener_spin)
        self._add_listener(key)

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addTextListener(self.events_listener_text)
        self._add_listener(key)

    # endregion Lazy Listeners

    # region Overrides

    def get_view_ctl(self) -> UnoControlFormattedField:
        return cast("UnoControlFormattedField", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlFormattedField``"""
        return "com.sun.star.awt.UnoControlFormattedField"

    def get_model(self) -> UnoControlFormattedFieldModel:
        """Gets the Model for the control"""
        return cast("UnoControlFormattedFieldModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlFormattedField:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlFormattedFieldModel:
        return self.get_model()

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        return self.model.Text

    @text.setter
    def text(self, value: str) -> None:
        self.model.Text = value

    @property
    def value(self) -> float:
        """Gets/Sets the value"""
        return self.model.EffectiveValue

    @value.setter
    def value(self, value: float) -> None:
        self.model.EffectiveValue = value

    # endregion Properties


# ctl = CtlButton(None)
