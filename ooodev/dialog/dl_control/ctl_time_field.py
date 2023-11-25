# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import datetime
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

# pylint: disable=useless-import-alias
from ooodev.utils.date_time_util import DateUtil
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.dialog_control_kind import DialogControlKind
from ooodev.utils.kind.dialog_control_named_kind import DialogControlNamedKind
from ooodev.utils.kind.time_format_kind import TimeFormatKind as TimeFormatKind
from .ctl_base import DialogControlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlTimeField  # service
    from com.sun.star.awt import UnoControlTimeFieldModel  # service
# endregion imports


class CtlTimeField(DialogControlBase, SpinEvents, TextEvents):
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
    def border(self) -> BorderKind:
        """Gets/Sets the border style"""
        return BorderKind(self.model.Border)

    @border.setter
    def border(self, value: BorderKind) -> None:
        self.model.Border = value.value

    @property
    def model(self) -> UnoControlTimeFieldModel:
        return self.get_model()

    @property
    def read_only(self) -> bool:
        """Gets/Sets the read-only property"""
        with contextlib.suppress(Exception):
            return self.model.ReadOnly
        return False

    @read_only.setter
    def read_only(self, value: bool) -> None:
        """Sets the read-only property"""
        with contextlib.suppress(Exception):
            self.model.ReadOnly = value

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        return self.model.Text

    @text.setter
    def text(self, value: str) -> None:
        self.model.Text = value

    @property
    def time(self) -> datetime.time:
        """Gets/Sets the time"""
        return DateUtil.uno_time_to_time(self.model.Time)

    @time.setter
    def time(self, value: datetime.time) -> None:
        self.model.Time = DateUtil.time_to_uno_time(value)

    @property
    def time_format(self) -> TimeFormatKind:
        """Gets/Sets the format"""
        return TimeFormatKind(self.model.TimeFormat)

    @time_format.setter
    def time_format(self, value: TimeFormatKind) -> None:
        self.model.TimeFormat = value.value

    @property
    def time_max(self) -> datetime.time:
        """Gets/Sets the min time"""
        return DateUtil.uno_time_to_time(self.model.TimeMax)

    @time_max.setter
    def time_max(self, value: datetime.time) -> None:
        self.model.TimeMax = DateUtil.time_to_uno_time(value)

    @property
    def time_min(self) -> datetime.time:
        """Gets/Sets the min time"""
        return DateUtil.uno_time_to_time(self.model.TimeMin)

    @time_min.setter
    def time_min(self, value: datetime.time) -> None:
        self.model.TimeMin = DateUtil.time_to_uno_time(value)

    @property
    def view(self) -> UnoControlTimeField:
        return self.get_view_ctl()

    # endregion Properties
