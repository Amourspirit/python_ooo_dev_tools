from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.event_args import EventArgs as EventArgs

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlCurrencyField  # service
    from com.sun.star.awt import UnoControlCurrencyFieldModel  # service


class CtlCurrencyField(CtlBase, SpinEvents, TextEvents):
    def __init__(self, ctl: UnoControlCurrencyField) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlCurrencyField): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        SpinEvents.__init__(self, trigger_args=generic_args)
        TextEvents.__init__(self, trigger_args=generic_args)
        self.view.addSpinListener(self.events_listener_spin)
        self.view.addTextListener(self.events_listener_text)

    def get_view_ctl(self) -> UnoControlCurrencyField:
        return cast("UnoControlCurrencyField", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlCurrencyField``"""
        return "com.sun.star.awt.UnoControlCurrencyField"

    @property
    def view(self) -> UnoControlCurrencyField:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlCurrencyFieldModel:
        return cast("UnoControlCurrencyFieldModel", self.get_view_ctl().getModel())

    @property
    def value(self) -> float:
        """Gets/Sets the value"""
        return self.model.Value

    @value.setter
    def value(self, value: float) -> None:
        self.model.Value = value


# ctl = CtlButton(None)
