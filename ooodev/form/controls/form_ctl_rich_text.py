from __future__ import annotations
from typing import Any, cast, TYPE_CHECKING
from com.sun.star.awt import XControl

from ooodev.adapter.awt.text_events import TextEvents
from ooodev.utils.kind.border_kind import BorderKind as BorderKind
from ooodev.utils.kind.form_component_kind import FormComponentKind

from .form_ctl_base import FormCtlBase

if TYPE_CHECKING:
    from com.sun.star.form.component import RichTextControl as ControlModel  # service
    from com.sun.star.form.control import TextField as ControlView  # service
    from ooodev.events.args.listener_event_args import ListenerEventArgs


class FormCtlRichText(FormCtlBase, TextEvents):
    """``com.sun.star.form.component.RichTextControl`` control"""

    def __init__(self, ctl: XControl) -> None:
        FormCtlBase.__init__(self, ctl)
        generic_args = self._get_generic_args()
        TextEvents.__init__(self, trigger_args=generic_args, cb=self._on_text_events_listener_add_remove)

    # region Lazy Listeners

    def _on_text_events_listener_add_remove(self, source: Any, event: ListenerEventArgs) -> None:
        # will only ever fire once
        self.view.addTextListener(self.events_listener_text)
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
        return FormComponentKind.RICH_TEXT_CONTROL

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
    def model(self) -> ControlModel:
        """Gets the model for this control"""
        return self.get_model()

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
