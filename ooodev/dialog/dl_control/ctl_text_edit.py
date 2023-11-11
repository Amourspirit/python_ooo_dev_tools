# region imports
from __future__ import annotations
import contextlib
from typing import Any, cast, TYPE_CHECKING

from ooodev.adapter.awt.text_events import TextEvents
from ooodev.events.args.listener_event_args import ListenerEventArgs

# com.sun.star.awt.LineEndFormat
from ooo.dyn.awt.line_end_format import LineEndFormatEnum as LineEndFormatEnum

from .ctl_base import CtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import UnoControlEdit  # service
    from com.sun.star.awt import UnoControlEditModel  # service
# endregion imports


class CtlTextEdit(CtlBase, TextEvents):
    """Class for Text Edit Control"""

    # region init
    def __init__(self, ctl: UnoControlEdit) -> None:
        """
        Constructor

        Args:
            ctl (UnoControlEdit): Button Control
        """
        # generally speaking EventArgs.event_data will contain the Event object for the UNO event raised.
        CtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        # EventArgs.event_data will contain the ActionEvent
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)

    # endregion init

    # region Lazy Listeners
    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        key = cast(str, event.source)
        if self._has_listener(key):
            return
        self.view.addTextListener(self.events_listener_text)
        self._add_listener(key)

    # endregion Lazy Listeners

    # region Overrides

    def get_view_ctl(self) -> UnoControlEdit:
        return cast("UnoControlEdit", super().get_view_ctl())

    def get_uno_srv_name(self) -> str:
        """Returns ``com.sun.star.awt.UnoControlEdit``"""
        return "com.sun.star.awt.UnoControlEdit"

    def get_model(self) -> UnoControlEditModel:
        """Gets the Model for the control"""
        return cast("UnoControlEditModel", self.get_view_ctl().getModel())

    # endregion Overrides

    # region Properties
    @property
    def view(self) -> UnoControlEdit:
        return self.get_view_ctl()

    @property
    def model(self) -> UnoControlEditModel:
        return self.get_model()

    @property
    def echo_char(self) -> str:
        """Gets/Sets the echo character as a string"""
        with contextlib.suppress(Exception):
            return chr(self.model.EchoChar)
        return ""

    @echo_char.setter
    def echo_char(self, value: str) -> None:
        if len(value) > 0:
            value = value[0]
        self.model.EchoChar = ord(value)

    @property
    def line_end_format(self) -> LineEndFormatEnum:
        """Gets/Sets the end line format"""
        return LineEndFormatEnum(self.model.LineEndFormat)

    @line_end_format.setter
    def line_end_format(self, value: LineEndFormatEnum) -> None:
        self.model.LineEndFormat = value.value

    @property
    def multi_line(self) -> bool:
        """Gets/Sets the multi line"""
        return self.model.MultiLine

    @multi_line.setter
    def multi_line(self, value: bool) -> None:
        self.model.MultiLine = value

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        return self.model.Text

    @text.setter
    def text(self, value: str) -> None:
        self.model.Text = value

    # endregion Properties


# ctl = CtlButton(None)
