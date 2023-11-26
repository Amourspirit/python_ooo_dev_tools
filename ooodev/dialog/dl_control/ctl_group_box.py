# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from .ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlGroupBox  # service
    from com.sun.star.awt import UnoControlGroupBoxModel  # service
# endregion imports


class CtlGroupBox(DialogControlBase):
    """Class for Group Box Control"""

    # region init
    def __init__(self, ctl: UnoControlGroupBox) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlGroupBox): Fixed Line Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)

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

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.GROUP_BOX``"""
        return DialogControlKind.GROUP_BOX

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.GROUP_BOX``"""
        return DialogControlNamedKind.GROUP_BOX

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlGroupBox:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlGroupBoxModel:
        return self.get_model()

    @property
    def label(self) -> str:
        """Gets/Sets the label"""
        return self.model.Label

    @label.setter
    def label(self, value: str) -> None:
        self.model.Label = value

    # endregion Properties
