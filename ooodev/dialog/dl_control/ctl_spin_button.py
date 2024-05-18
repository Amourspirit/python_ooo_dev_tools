# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.uno_control_spin_button_model_partial import UnoControlSpinButtonModelPartial
from ooodev.adapter.awt.spin_value_partial import SpinValuePartial
from ooodev.adapter.awt.adjustment_events import AdjustmentEvents
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlSpinButton  # service
    from com.sun.star.awt import UnoControlSpinButtonModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_spin_button import ModelSpinButton
    from ooodev.dialog.dl_control.view.view_spin_button import ViewSpinButton
# endregion imports


class CtlSpinButton(DialogControlBase, UnoControlSpinButtonModelPartial, SpinValuePartial, AdjustmentEvents):
    """Class for Spin Button Control"""

    # region init
    def __init__(self, ctl: UnoControlSpinButton) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlSpinButton): Progress Bar Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlSpinButtonModelPartial.__init__(self, component=self.get_model())
        SpinValuePartial.__init__(self, component=self.get_view())  # type: ignore
        generic_args = self._get_generic_args()
        AdjustmentEvents.__init__(self, trigger_args=generic_args, cb=self._on_adjustment_events_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlSpinButton({self.name})"
        return "CtlSpinButton"

    # region Lazy Listeners
    def _on_adjustment_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addAdjustmentListener(self.events_listener_adjustment)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides
    def get_view_ctl(self) -> UnoControlSpinButton:
        return cast("UnoControlSpinButton", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlSpinButton``"""
        return "com.sun.star.awt.UnoControlSpinButton"

    def get_model(self) -> UnoControlSpinButtonModel:
        """Gets the Model for the control"""
        return cast("UnoControlSpinButtonModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.SPIN_BUTTON``"""
        return DialogControlKind.SPIN_BUTTON

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.SPIN_BUTTON``"""
        return DialogControlNamedKind.SPIN_BUTTON

    # endregion Overrides

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlSpinButton":
        """
        Creates a new instance of the control.

        Keyword arguments are optional.
        Extra Keyword args are passed to the control as property values.

        Args:
            win (XWindowPeer): Parent Window

        Keyword Args:
            x (int, UnitT, optional): X Position in Pixels or UnitT.
            y (int, UnitT, optional): Y Position in Pixels or UnitT.
            width (int, UnitT, optional): Width in Pixels or UnitT.
            height (int, UnitT, optional): Height in Pixels or UnitT.

        Returns:
            CtlSpinButton: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlSpinButtonModel", win, **kwargs)
        return CtlSpinButton(ctl=ctrl)

    # endregion Static Methods

    # region Properties

    @property
    def model(self) -> UnoControlSpinButtonModel:
        # pylint: disable=no-member
        return cast("UnoControlSpinButtonModel", super().model)

    @property
    def model_ex(self) -> ModelSpinButton:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_spin_button import ModelSpinButton

            self._model_ex = ModelSpinButton(self.model)
        return self._model_ex

    @property
    def view(self) -> UnoControlSpinButton:
        # pylint: disable=no-member
        return cast("UnoControlSpinButton", super().view)

    @property
    def view_ex(self) -> ViewSpinButton:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_spin_button import ViewSpinButton

            self._view_ex = ViewSpinButton(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_spin_button import ModelSpinButton
    from ooodev.dialog.dl_control.view.view_spin_button import ViewSpinButton
