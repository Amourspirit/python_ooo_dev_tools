# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING

from .ctl_base import CtlListenerBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedText  # service
    from com.sun.star.awt import UnoControlFixedTextModel  # service
# endregion imports


class CtlFixedText(CtlListenerBase):
    """Class for Fixed Text Control"""

    # region init
    def __init__(self, ctl: UnoControlFixedText) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFixedText): Fixed Text Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlListenerBase.__init__(self, ctl)

    # endregion init

    # region Overrides
    def get_view_ctl(self) -> UnoControlFixedText:
        return cast("UnoControlFixedText", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlFixedText``"""
        return "com.sun.star.awt.UnoControlFixedText"

    def get_model(self) -> UnoControlFixedTextModel:
        """Gets the Model for the control"""
        return cast("UnoControlFixedTextModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlFixedText:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlFixedTextModel:
        return self.get_model()

    # endregion Properties
