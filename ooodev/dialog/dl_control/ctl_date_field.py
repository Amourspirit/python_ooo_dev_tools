# region imports
from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import datetime

from ooodev.adapter.awt.spin_events import SpinEvents
from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs
from ooodev.utils.kind.date_format_kind import DateFormatKind as DateFormatKind
from ooodev.utils.date_time_util import DateUtil
from .ctl_base import CtlListenerBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlDateField  # service
    from com.sun.star.awt import UnoControlDateFieldModel  # service
# endregion imports


class CtlDateField(CtlListenerBase, SpinEvents, TextEvents):
    """Class for Button Control"""

    # region init
    def __init__(self, ctl: UnoControlDateField) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlDateField): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlListenerBase.__init__(self, ctl)
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
    def get_view_ctl(self) -> UnoControlDateField:
        return cast("UnoControlDateField", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlDateField``"""
        return "com.sun.star.awt.UnoControlDateField"

    def get_model(self) -> UnoControlDateFieldModel:
        """Gets the Model for the control"""
        return cast("UnoControlDateFieldModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlDateField:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlDateFieldModel:
        return self.get_model()

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        return self.model.Text

    @text.setter
    def text(self, value: str) -> None:
        self.model.Text = value

    @property
    def date(self) -> datetime.date:
        """Gets/Sets the date"""
        return DateUtil.uno_date_to_date(self.model.Date)

    @date.setter
    def date(self, value: datetime.date) -> None:
        self.model.Date = DateUtil.date_to_uno_date(value)

    @property
    def date_min(self) -> datetime.date:
        """Gets/Sets the min date"""
        return DateUtil.uno_date_to_date(self.model.DateMin)

    @date_min.setter
    def date_min(self, value: datetime.date) -> None:
        self.model.DateMin = DateUtil.date_to_uno_date(value)

    @property
    def date_max(self) -> datetime.date:
        """Gets/Sets the min date"""
        return DateUtil.uno_date_to_date(self.model.DateMax)

    @date_max.setter
    def date_max(self, value: datetime.date) -> None:
        self.model.DateMax = DateUtil.date_to_uno_date(value)

    @property
    def date_format(self) -> DateFormatKind:
        """Gets/Sets the format"""
        return DateFormatKind(self.model.DateFormat)

    @date_format.setter
    def date_format(self, value: DateFormatKind) -> None:
        self.model.DateFormat = value.value

    # endregion Properties
