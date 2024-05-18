# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.mock import mock_g
from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.adapter.awt.uno_control_currency_field_model_partial import UnoControlCurrencyFieldModelPartial

from ooodev.dialog.dl_control.ctl_base import DialogControlBase, _create_control

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlCurrencyField  # service
    from com.sun.star.awt import UnoControlCurrencyFieldModel  # service
    from com.sun.star.awt import XWindowPeer
    from ooodev.dialog.dl_control.model.model_currency_field import ModelCurrencyField
    from ooodev.dialog.dl_control.view.view_currency_field import ViewCurrencyField
# endregion imports


class CtlCurrencyField(DialogControlBase, UnoControlCurrencyFieldModelPartial, SpinEvents, TextEvents):
    """Class for Currency Field Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlCurrencyField) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlCurrencyField): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlCurrencyFieldModelPartial.__init__(self, component=self.get_model())
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        SpinEvents.__init__(self, trigger_args=generic_args, cb=self._on_spin_events_listener_add_remove)
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)
        self._model_ex = None
        self._view_ex = None

    # endregion init

    def __repr__(self) -> str:
        if hasattr(self, "name"):
            return f"CtlCurrencyField({self.name})"
        return "CtlCurrencyField"

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

    def get_view_ctl(self) -> UnoControlCurrencyField:
        return cast("UnoControlCurrencyField", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlCurrencyField``"""
        return "com.sun.star.awt.UnoControlCurrencyField"

    def get_model(self) -> UnoControlCurrencyFieldModel:
        """Gets the Model for the control"""
        return cast("UnoControlCurrencyFieldModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.CURRENCY``"""
        return DialogControlKind.CURRENCY

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.CURRENCY``"""
        return DialogControlNamedKind.CURRENCY

    # endregion Overrides

    # region Static Methods
    @staticmethod
    def create(win: XWindowPeer, **kwargs: Any) -> "CtlCurrencyField":
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
            CtlCurrencyField: New instance of the control.

        Note:
            The `UnoControlDialogElement <https://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1awt_1_1UnoControlDialogElement.html>`__
            interface is not included when creating the control with a window peer.
        """
        ctrl = _create_control("com.sun.star.awt.UnoControlCurrencyFieldModel", win, **kwargs)
        return CtlCurrencyField(ctl=ctrl)

    # endregion Static Methods

    # region Properties
    @property
    def accuracy(self) -> int:
        """Gets/Sets the accuracy"""
        return self.decimal_accuracy

    @accuracy.setter
    def accuracy(self, value: int) -> None:
        self.decimal_accuracy = value

    @property
    def increment(self) -> float:
        """Gets/Sets the increment value"""
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
    def model(self) -> UnoControlCurrencyFieldModel:
        # pylint: disable=no-member
        return cast("UnoControlCurrencyFieldModel", super().model)

    @property
    def model_ex(self) -> ModelCurrencyField:
        """
        Gets the extended Model for the control.

        This is a wrapped instance for the model property.
        It add some additional properties and methods to the model.
        """
        # pylint: disable=no-member
        if self._model_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.model.model_currency_field import ModelCurrencyField

            self._model_ex = ModelCurrencyField(self.model)
        return self._model_ex

    @property
    def spin_button(self) -> bool:
        """Gets/Sets the spin button property. Same as ``spin`` property."""
        return self.spin

    @spin_button.setter
    def spin_button(self, value: bool) -> None:
        self.spin = value

    @property
    def view(self) -> UnoControlCurrencyField:
        # pylint: disable=no-member
        return cast("UnoControlCurrencyField", super().view)

    @property
    def view_ex(self) -> ViewCurrencyField:
        """
        Gets the extended View for the control.

        This is a wrapped instance for the view property.
        It add some additional properties and methods to the view.
        """
        if self._view_ex is None:
            # pylint: disable=import-outside-toplevel
            # pylint: disable=redefined-outer-name
            from ooodev.dialog.dl_control.view.view_currency_field import ViewCurrencyField

            self._view_ex = ViewCurrencyField(self.view)
        return self._view_ex

    # endregion Properties


# ctl = CtlButton(None)

if mock_g.FULL_IMPORT:
    from ooodev.dialog.dl_control.model.model_currency_field import ModelCurrencyField
    from ooodev.dialog.dl_control.view.view_currency_field import ViewCurrencyField
