# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import datetime
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.uno_control_time_field_model_partial import UnoControlTimeFieldModelPartial
from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

# pylint: disable=useless-import-alias
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlTimeField  # service
    from com.sun.star.awt import UnoControlTimeFieldModel  # service
# endregion imports


class CtlTimeField(DialogControlBase, UnoControlTimeFieldModelPartial, SpinEvents, TextEvents):
    """Class for Time Field Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlTimeField) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlTimeField): Time Field Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlTimeFieldModelPartial.__init__(self)
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
    def get_view_ctl(self) -> UnoControlTimeField:
        return cast("UnoControlTimeField", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlTimeField``"""
        return "com.sun.star.awt.UnoControlTimeField"

    def get_model(self) -> UnoControlTimeFieldModel:
        """Gets the Model for the control"""
        return cast("UnoControlTimeFieldModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.TIME``"""
        return DialogControlKind.TIME

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.TIME``"""
        return DialogControlNamedKind.TIME

    # endregion Overrides

    # region Properties

    @property
    def model(self) -> UnoControlTimeFieldModel:
        # pylint: disable=no-member
        return cast("UnoControlTimeFieldModel", super().model)

    # region UnoControlTimeFieldModelPartial Overrides

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        val = super().text
        return "" if val is None else val

    @text.setter
    def text(self, value: str) -> None:
        super().text = value

    # endregion UnoControlTimeFieldModelPartial Overrides

    @property
    def view(self) -> UnoControlTimeField:
        # pylint: disable=no-member
        return cast("UnoControlTimeField", super().view)

    # endregion Properties
