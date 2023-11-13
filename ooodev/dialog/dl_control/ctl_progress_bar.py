# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING

from ooodev.utils.color import Color
from .ctl_base import CtlListenerBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlProgressBar  # service
    from com.sun.star.awt import UnoControlProgressBarModel  # service
# endregion imports


class CtlProgressBar(CtlListenerBase):
    """Class for Progress Bar Control"""

    # region init
    def __init__(self, ctl: UnoControlProgressBar) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlProgressBar): Progress Bar Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlListenerBase.__init__(self, ctl)

    # endregion init

    # region Overrides
    def get_view_ctl(self) -> UnoControlProgressBar:
        return cast("UnoControlProgressBar", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlProgressBar``"""
        return "com.sun.star.awt.UnoControlProgressBar"

    def get_model(self) -> UnoControlProgressBarModel:
        """Gets the Model for the control"""
        return cast("UnoControlProgressBarModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlProgressBar:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlProgressBarModel:
        return self.get_model()

    @property
    def value(self) -> int:
        """Gets or sets the current value of the progress bar"""
        return self.view.getValue()

    @value.setter
    def value(self, value: int) -> None:
        self.view.setValue(value)

    @property
    def fill_color(self) -> Color:
        """Gets or sets the fill color of the progress bar"""
        return Color(self.model.FillColor)

    @fill_color.setter
    def fill_color(self, value: Color) -> None:
        self.model.FillColor = value  # type: ignore

    # endregion Properties