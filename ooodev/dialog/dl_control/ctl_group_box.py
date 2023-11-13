# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING

from .ctl_base import CtlListenerBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlGroupBox  # service
    from com.sun.star.awt import UnoControlGroupBoxModel  # service
# endregion imports


class CtlGroupBox(CtlListenerBase):
    """Class for Group Box Control"""

    # region init
    def __init__(self, ctl: UnoControlGroupBox) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlGroupBox): Fixed Line Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlListenerBase.__init__(self, ctl)

    # endregion init

    # region Overrides
    def get_view_ctl(self) -> UnoControlGroupBox:
        return cast("UnoControlGroupBox", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlGroupBox``"""
        return "com.sun.star.awt.UnoControlGroupBox"

    def get_model(self) -> UnoControlGroupBoxModel:
        """Gets the Model for the control"""
        return cast("UnoControlGroupBoxModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlGroupBox:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlGroupBoxModel:
        return self.get_model()

    # endregion Properties
