# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.uno_control_progress_bar_model_partial import UnoControlProgressBarModelPartial
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlProgressBar  # service
    from com.sun.star.awt import UnoControlProgressBarModel  # service
# endregion imports


class CtlProgressBar(DialogControlBase, UnoControlProgressBarModelPartial):
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
        UnoControlProgressBarModelPartial.__init__(self)

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
    def model(self) -> UnoControlProgressBarModel:
        # pylint: disable=no-member
        return cast("UnoControlProgressBarModel", super().model)

    @property
    def value(self) -> int:
        """Gets or sets the current value of the progress bar"""
        return self.view.getValue()

    @value.setter
    def value(self, value: int) -> None:
        self.view.setValue(value)

    @property
    def view(self) -> UnoControlProgressBar:
        # pylint: disable=no-member
        return cast("UnoControlProgressBar", super().view)

    # endregion Properties
