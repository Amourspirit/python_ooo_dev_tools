# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.uno_control_fixed_line_model_partial import UnoControlFixedLineModelPartial
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFixedLine  # service
    from com.sun.star.awt import UnoControlFixedLineModel  # service
    from ooodev.dialog.dl_control.model.model_fixed_line import ModelFixedLine
    from ooodev.dialog.dl_control.view.view_fixed_line import ViewFixedLine

# endregion imports


class CtlFixedLine(DialogControlBase, UnoControlFixedLineModelPartial):
    """Class for Fixed Line Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlFixedLine) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFixedLine): Fixed Line Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlFixedLineModelPartial.__init__(self, component=self.get_model())
        self._model_ex = None
        self._view_ex = None

    # endregion init

    # region Overrides
    def get_view_ctl(self) -> UnoControlFixedLine:
        return cast("UnoControlFixedLine", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlFixedLine``"""
        return "com.sun.star.awt.UnoControlFixedLine"

    def get_model(self) -> UnoControlFixedLineModel:
        """Gets the Model for the control"""
        return cast("UnoControlFixedLineModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.FIXED_LINE``"""
        return DialogControlKind.FIXED_LINE

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.FIXED_LINE``"""
        return DialogControlNamedKind.FIXED_LINE

    # endregion Overrides

    # region Properties
    @property
    def model(self) -> UnoControlFixedLineModel:
        # pylint: disable=no-member
        return cast("UnoControlFixedLineModel", super().model)

    @property
    def model_ex(self) -> ModelFixedLine:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_fixed_line import ModelFixedLine

            self._model_ex = ModelFixedLine(self.model)
        return self._model_ex

    @property
    def view(self) -> UnoControlFixedLine:
        # pylint: disable=no-member
        return cast("UnoControlFixedLine", super().view)

    @property
    def view_ex(self) -> ViewFixedLine:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_fixed_line import ViewFixedLine

            self._view_ex = ViewFixedLine(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_fixed_line import ModelFixedLine
    from ooodev.dialog.dl_control.view.view_fixed_line import ViewFixedLine
