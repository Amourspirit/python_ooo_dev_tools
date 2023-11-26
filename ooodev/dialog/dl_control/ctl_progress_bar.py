# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.utils.color import Color
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from .ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlProgressBar  # service
    from com.sun.star.awt import UnoControlProgressBarModel  # service
# endregion imports


class CtlProgressBar(DialogControlBase):
    """Class for Progress Bar Control"""

    # region init
    def __init__(self, ctl: UnoControlProgressBar) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlProgressBar): Progress Bar Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)

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

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.PROGRESS_BAR``"""
        return DialogControlKind.PROGRESS_BAR

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.PROGRESS_BAR``"""
        return DialogControlNamedKind.PROGRESS_BAR

    # endregion Overrides

    # region Properties
    @property
    def border(self) -> BorderKind:
        """Gets/Sets the border style"""
        return BorderKind(self.model.Border)

    @border.setter
    def border(self, value: BorderKind) -> None:
        self.model.Border = value.value

    @property
    def fill_color(self) -> Color:
        """Gets or sets the fill color of the progress bar"""
        return Color(self.model.FillColor)

    @fill_color.setter
    def fill_color(self, value: Color) -> None:
        self.model.FillColor = value  # type: ignore

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
    def view(self) -> UnoControlProgressBar:
        return self.get_view_ctl()

    # endregion Properties
