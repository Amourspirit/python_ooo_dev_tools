# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.uno_control_numeric_field_model_partial import UnoControlNumericFieldModelPartial
from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlNumericField  # service
    from com.sun.star.awt import UnoControlNumericFieldModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_numeric_field import ModelNumericField
    from ooodev.dialog.dl_control.view.view_numeric_field import ViewNumericField
# endregion imports


class CtlNumericField(DialogControlBase, UnoControlNumericFieldModelPartial, SpinEvents, TextEvents):
    """Class for Numeric Field Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlNumericField) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlNumericField): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlNumericFieldModelPartial.__init__(self, component=self.get_model())
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        SpinEvents.__init__(self, trigger_args=generic_args, cb=self._on_spin_events_listener_add_remove)
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init
    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlNumericField({self.name})"
        return "CtlNumericField"

    # region Lazy Listeners
    def _on_spin_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addSpinListener(self.events_listener_spin)
        event.remove_callback = True

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTextListener(self.events_listener_text)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides

    def get_view_ctl(self) -> UnoControlNumericField:
        return cast("UnoControlNumericField", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlNumericField``"""
        return "com.sun.star.awt.UnoControlNumericField"

    def get_model(self) -> UnoControlNumericFieldModel:
        """Gets the Model for the control"""
        return cast("UnoControlNumericFieldModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.NUMERIC``"""
        return DialogControlKind.NUMERIC

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.NUMERIC``"""
        return DialogControlNamedKind.NUMERIC

    # endregion Overrides

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlNumericField":
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
            CtlNumericField: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlNumericFieldModel", win, **kwargs)
        return CtlNumericField(ctl=ctrl)

    # endregion Static Methods

    # region Properties
    @property
    def accuracy(self) -> int:
        """Gets/Sets the accuracy. Same as ``decimal_accuracy`` property."""
        return self.decimal_accuracy

    @accuracy.setter
    def accuracy(self, value: int) -> None:
        self.decimal_accuracy = value

    @property
    def increment(self) -> float:
        """Gets/Sets the increment value. Same as ``value_step`` property."""
        return self.value_step

    @increment.setter
    def increment(self, value: float) -> None:
        self.value_step = value

    @property
    def max_value(self) -> float:
        """Gets/Sets the maximum value. Same as ``value_max`` property."""
        return self.value_max

    @max_value.setter
    def max_value(self, value: float) -> None:
        self.value_max = value

    @property
    def min_value(self) -> float:
        """Gets/Sets the minimum value. Same as ``value_min`` property."""
        return self.value_min

    @min_value.setter
    def min_value(self, value: float) -> None:
        self.value_min = value

    @property
    def model(self) -> UnoControlNumericFieldModel:
        # pylint: disable=no-member
        return cast("UnoControlNumericFieldModel", super().model)

    @property
    def model_ex(self) -> ModelNumericField:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_numeric_field import ModelNumericField

            self._model_ex = ModelNumericField(self.model)
        return self._model_ex

    @property
    def spin_button(self) -> bool:
        """Gets/Sets the spin button property. Same as ``spin`` property."""
        return self.spin

    @spin_button.setter
    def spin_button(self, value: bool) -> None:
        self.spin = value

    @property
    def view(self) -> UnoControlNumericField:
        # pylint: disable=no-member
        return cast("UnoControlNumericField", super().view)

    @property
    def view_ex(self) -> ViewNumericField:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        # pylint: disable=no-member
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_numeric_field import ViewNumericField

            self._view_ex = ViewNumericField(self.view)
        return self._view_ex

    # endregion Properties


if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_numeric_field import ModelNumericField
    from ooodev.dialog.dl_control.view.view_numeric_field import ViewNumericField
