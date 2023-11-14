# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import datetime
import uno  # pylint: disable=unused-import

from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

# pylint: disable=useless-import-alias
from ooodev.utils.kind.time_format_kind import TimeFormatKind as TimeFormatKind
from ooodev.utils.date_time_util import DateUtil
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
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addSpinListener(self.events_listener_spin)
        self._add_listener(key)

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addTextListener(self.events_listener_text)
        self._add_listener(key)

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

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlTimeField:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlTimeFieldModel:
        return self.get_model()

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
    def time_min(self) -> datetime.time:
        """Gets/Sets the min time"""
        return DateUtil.uno_time_to_time(self.model.TimeMin)

    @time_min.setter
    def time_min(self, value: datetime.time) -> None:
        self.model.TimeMin = DateUtil.time_to_uno_time(value)

    @property
    def date_max(self) -> datetime.time:
        """Gets/Sets the min time"""
        return DateUtil.uno_time_to_time(self.model.TimeMax)

    @date_max.setter
    def date_max(self, value: datetime.time) -> None:
        self.model.TimeMax = DateUtil.time_to_uno_time(value)

    @property
    def time_format(self) -> TimeFormatKind:
        """Gets/Sets the format"""
        return TimeFormatKind(self.model.TimeFormat)

    @time_format.setter
    def time_format(self, value: TimeFormatKind) -> None:
        self.model.TimeFormat = value.value

    # endregion Properties
