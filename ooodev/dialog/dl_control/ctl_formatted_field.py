# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.uno_control_formatted_field_model_partial import UnoControlFormattedFieldModelPartial
from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind

from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlFormattedField  # service
    from com.sun.star.awt import UnoControlFormattedFieldModel  # service
# endregion imports


class CtlFormattedField(DialogControlBase, UnoControlFormattedFieldModelPartial, SpinEvents, TextEvents):
    """Class for Formatted Field Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlFormattedField) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlFormattedField): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlFormattedFieldModelPartial.__init__(self)
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

    def get_view_ctl(self) -> UnoControlFormattedField:
        return cast("UnoControlFormattedField", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlFormattedField``"""
        return "com.sun.star.awt.UnoControlFormattedField"

    def get_model(self) -> UnoControlFormattedFieldModel:
        """Gets the Model for the control"""
        return cast("UnoControlFormattedFieldModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.FORMATTED_TEXT``"""
        return DialogControlKind.FORMATTED_TEXT

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.FORMATTED_TEXT``"""
        return DialogControlNamedKind.FORMATTED_TEXT

    # endregion Overrides

    # region Properties

    @property
    def model(self) -> UnoControlFormattedFieldModel:
        # pylint: disable=no-member
        return cast("UnoControlFormattedFieldModel", super().model)

    @property
    def value(self) -> Any:
        """
        Gets/Sets the value.

        Same as ``effective_value`` property

        This may be a numeric value (float) or a string, depending on the formatting of the field.
        """
        return self.effective_value

    @value.setter
    def value(self, value: Any) -> None:
        self.effective_value = value

    @property
    def view(self) -> UnoControlFormattedField:
        # pylint: disable=no-member
        return cast("UnoControlFormattedField", super().view)

    # endregion Properties


# ctl = CtlButton(None)
