# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.uno_control_numeric_field_model_partial import UnoControlNumericFieldModelPartial
from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind

from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlNumericField  # service
    from com.sun.star.awt import UnoControlNumericFieldModel  # service
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
        UnoControlNumericFieldModelPartial.__init__(self)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        SpinEvents.__init__(self, trigger_args=generic_args, cb=self._on_spin_events_listener_add_remove)
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)

    # endregion init

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

    # endregion Properties
