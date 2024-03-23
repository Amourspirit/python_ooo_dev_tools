# region imports
from __future__ import annotations
from typing import cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.uno_control_group_box_model_partial import UnoControlGroupBoxModelPartial
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlGroupBox  # service
    from com.sun.star.awt import UnoControlGroupBoxModel  # service
    from ooodev.dialog.dl_control.model.model_group_box import ModelGroupBox
# endregion imports


class CtlGroupBox(DialogControlBase, UnoControlGroupBoxModelPartial):
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
        UnoControlGroupBoxModelPartial.__init__(self, component=self.get_model())
        self._model_ex = None

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
    def model(self) -> UnoControlGroupBoxModel:
        # pylint: disable=no-member
        return cast("UnoControlGroupBoxModel", super().model)

    @property
    def model_ex(self) -> ModelGroupBox:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_group_box import ModelGroupBox

            self._model_ex = ModelGroupBox(self.model)
        return self._model_ex

    @property
    def view(self) -> UnoControlGroupBox:
        # pylint: disable=no-member
        return cast("UnoControlGroupBox", super().view)

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_group_box import ModelGroupBox
