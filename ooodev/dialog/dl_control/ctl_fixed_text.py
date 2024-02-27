# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.adapter.awt.uno_control_fixed_text_model_partial import UnoControlFixedTextModelPartial
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedText  # service
    from com.sun.star.awt import UnoControlFixedTextModel  # service
# endregion imports


class CtlFixedText(DialogControlBase, UnoControlFixedTextModelPartial):
    """Class for Fixed Text Control"""

    # region init
    def __init__(self, ctl: UnoControlFixedText) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFixedText): Fixed Text Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlFixedTextModelPartial.__init__(self)

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

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.FIXED_TEXT``"""
        return DialogControlKind.FIXED_TEXT

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.FIXED_TEXT``"""
        return DialogControlNamedKind.FIXED_TEXT

    # endregion Overrides

    # region Properties

    @property
    def model(self) -> UnoControlFixedTextModel:
        # pylint: disable=no-member
        return cast("UnoControlFixedTextModel", super().model)

    @property
    def view(self) -> UnoControlFixedText:
        # pylint: disable=no-member
        return cast("UnoControlFixedText", super().view)

    # endregion Properties
