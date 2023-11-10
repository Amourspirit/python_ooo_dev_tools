from __future__ import annotations
from typing import cast, TYPE_CHECKING

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedText  # service
    from com.sun.star.awt import UnoControlFixedTextModel  # service


class CtlFixedText(CtlBase):
    def __init__(self, ctl: UnoControlFixedText) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFixedText): Fixed Text Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlBase.__init__(self, ctl)

    def get_view_ctl(self) -> UnoControlFixedText:
        return cast("UnoControlFixedText", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlFixedText``"""
        return "com.sun.star.awt.UnoControlFixedText"

    @property
    def view(self) -> UnoControlFixedText:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlFixedTextModel:
        return cast("UnoControlFixedTextModel", self.get_view_ctl().getModel())
