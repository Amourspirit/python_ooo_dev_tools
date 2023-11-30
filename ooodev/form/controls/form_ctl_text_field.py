from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
import contextlib
import os

import uno
from ooo.dyn.awt.line_end_format import LineEndFormatEnum as LineEndFormatEnum
from ooo.dyn.awt.selection import Selection

from ooodev.adapter.awt.text_events import TextEvents
from ooodev.adapter.form.reset_events import ResetEvents
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.form_component_kind import FormComponentKind

from .form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.awt import XControl
    from com.sun.star.form.component import TextField as ControlModel  # service
    from com.sun.star.form.control import TextField as ControlView  # service
    from ooodev.events.args.listener_event_args import ListenerEventArgs


class FormCtlTextField(FormCtlBase, TextEvents, ResetEvents):
    """``com.sun.star.form.component.TextField`` control"""

    def __init__(self, ctl: XControl) -> None:
        FormCtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)
        ResetEvents.__init__(self, trigger_args=generic_args, cb=self._on_reset_add_remove)

    # region Lazy Listeners

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTextListener(self.events_listener_text)
        event.remove_callback = True

    def _on_reset_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.model.addResetListener(self.events_listener_reset)
        event.remove_callback = True

    # endregion Lazy Listeners

    # region Overrides

    if TYPE_CHECKING:
        # override the methods to provide type hinting
        def get_view(self) -> ControlView:
            """Gets the view of this control"""
            return cast("ControlView", super().get_view())

        def get_model(self) -> ControlModel:
            """Gets the model for this control"""
            return cast("ControlModel", super().get_model())

    def get_form_component_kind(self) -> FormComponentKind:
        """Gets the kind of form component this control is"""
        return FormComponentKind.TEXT_FIELD

    # endregion Overrides

    # region Text Methods
    def write_line(self, line: str = "") -> bool:
        """
        Add a new line to a multi-line text control

        Args:
            line (str, optional): Specifies a line to insert at the end of the text box
                a newline character will be inserted before the line, if relevant.

        Returns:
            bool: True if successful, False otherwise
        """
        # if not self.model.MultiLine:
        #     return False
        with contextlib.suppress(Exception):
            # will raise an exception if not multi-line
            self.model.HardLineBreaks = True
            sel = Selection()
            text_len = len(self.text)
            if text_len == 0:
                sel.Min = 0
                sel.Max = 0
                self.text = line
            else:
                # Put cursor at the end of the actual text
                sel.Min = text_len
                sel.Max = text_len
                self.view.insertText(sel, f"{os.linesep}{line}")
            # Put the cursor at the end of the inserted text
            sel.Max += len(os.linesep) + len(line)
            sel.Min = sel.Max
            self.view.setSelection(sel)
            return True
        return False

    # endregion Text Methods

    # region Properties
    @property
    def border(self) -> BorderKind:
        """Gets/Sets the border style"""
        return BorderKind(self.model.Border)

    @border.setter
    def border(self, value: BorderKind) -> None:
        self.model.Border = value.value

    @property
    def echo_char(self) -> str:
        """Gets/Sets the echo character as a string"""
        return chr(self.model.EchoChar)

    @echo_char.setter
    def echo_char(self, value: str) -> None:
        if len(value) > 0:
            value = value[0]
        self.model.EchoChar = ord(value)

    @property
    def enabled(self) -> bool:
        """Gets/Sets the enabled state for the control"""
        return self.model.Enabled

    @enabled.setter
    def enabled(self, value: bool) -> None:
        self.model.Enabled = value

    @property
    def help_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @help_text.setter
    def help_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def help_url(self) -> str:
        """Gets/Sets the help url"""
        return self.model.HelpURL

    @help_url.setter
    def help_url(self, value: str) -> None:
        self.model.HelpURL = value

    @property
    def line_end_format(self) -> LineEndFormatEnum:
        """Gets/Sets the end line format"""
        return LineEndFormatEnum(self.model.LineEndFormat)

    @line_end_format.setter
    def line_end_format(self, value: LineEndFormatEnum) -> None:
        self.model.LineEndFormat = value.value

    @property
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

    @property
    def multi_line(self) -> bool:
        """Gets/Sets the multi line"""
        return self.model.MultiLine

    @multi_line.setter
    def multi_line(self, value: bool) -> None:
        self.model.MultiLine = value

    @property
    def printable(self) -> bool:
        """Gets/Sets the printable property"""
        return self.model.Printable

    @printable.setter
    def printable(self, value: bool) -> None:
        self.model.Printable = value

    @property
    def read_only(self) -> bool:
        """Gets/Sets the read-only property"""
        return self.model.ReadOnly

    @read_only.setter
    def read_only(self, value: bool) -> None:
        """Sets the read-only property"""
        self.model.ReadOnly = value

    @property
    def step(self) -> int:
        """Gets/Sets the step"""
        return self.model.Step

    @step.setter
    def step(self, value: int) -> None:
        self.model.Step = value

    @property
    def tab_stop(self) -> bool:
        """Gets/Sets the tab stop property"""
        return self.model.Tabstop

    @tab_stop.setter
    def tab_stop(self, value: bool) -> None:
        self.model.Tabstop = value

    @property
    def text(self) -> str:
        """Gets/Sets the text"""
        return self.model.Text

    @text.setter
    def text(self, value: str) -> None:
        self.model.Text = value

    @property
    def tip_text(self) -> str:
        """Gets/Sets the tip text"""
        return self.model.HelpText

    @tip_text.setter
    def tip_text(self, value: str) -> None:
        self.model.HelpText = value

    @property
    def view(self) -> ControlView:
        """Gets the view of this control"""
        return self.get_view()

    # endregion Properties
