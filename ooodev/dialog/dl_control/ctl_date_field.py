# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.adapter.awt.uno_control_date_field_model_partial import UnoControlDateFieldModelPartial
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind

# pylint: disable=useless-import-alias
from ooodev.utils.kind.date_format_kind import DateFormatKind as DateFormatKind
from ooodev.dialog.dl_control.ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDateField  # service
    from com.sun.star.awt import UnoControlDateFieldModel  # service
# endregion imports


class CtlDateField(DialogControlBase, UnoControlDateFieldModelPartial, SpinEvents, TextEvents):
    """Class for Date Field Control"""

    # pylint: disable=unused-argument

    # region init
    def __init__(self, ctl: UnoControlDateField) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlDateField): Date Field Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        DialogControlBase.__init__(self, ctl)
        UnoControlDateFieldModelPartial.__init__(self)
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
    def get_view_ctl(self) -> UnoControlDateField:
        return cast("UnoControlDateField", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlDateField``"""
        return "com.sun.star.awt.UnoControlDateField"

    def get_model(self) -> UnoControlDateFieldModel:
        """Gets the Model for the control"""
        return cast("UnoControlDateFieldModel", self.get_view_ctl().getModel())

    def get_control_kind(self) -> DialogControlKind:
        """Gets the control kind. Returns ``DialogControlKind.DATE_FIELD``"""
        return DialogControlKind.DATE_FIELD

    def get_control_named_kind(self) -> DialogControlNamedKind:
        """Gets the control named kind. Returns ``DialogControlNamedKind.DATE_FIELD``"""
        return DialogControlNamedKind.DATE_FIELD

    # endregion Overrides

    # region Properties

    @property
    def model(self) -> UnoControlDateFieldModel:
        # pylint: disable=no-member
        return cast("UnoControlDateFieldModel", super().model)

    @property
    def view(self) -> UnoControlDateField:
        # pylint: disable=no-member
        return cast("UnoControlDateField", super().view)

    # endregion Properties
