# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.adapter.awt.uno_control_fixed_text_model_partial import UnoControlFixedTextModelPartial
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from .ctl_base import DialogControlBase

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
        UnoControlFixedTextModelPartial.__init__(self, component=self.model)

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
    def border(self) -> BorderKind:
        """Gets/Sets the border style"""
        return BorderKind(self.model.Border)

    @border.setter
    def border(self, value: BorderKind) -> None:
        self.model.Border = value.value

    @property
    def label(self) -> str:
        """Gets/Sets the label"""
        return self.model.Label

    @label.setter
    def label(self, value: str) -> None:
        self.model.Label = value

    @property
    def model(self) -> UnoControlFixedTextModel:
        return self.get_model()

    @property
    def multi_line(self) -> bool:
        """Gets/Sets the multi line"""
        return self.model.MultiLine

    @multi_line.setter
    def multi_line(self, value: bool) -> None:
        self.model.MultiLine = value

    @property
    def view(self) -> UnoControlFixedText:
        return self.get_view_ctl()

    # endregion Properties
